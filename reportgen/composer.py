"""Composer v2: собирает отчёт из плана по канве Петровой/Порубова.

Шаги плана могут быть:
  - cover / final — обычное text-наполнение шаблонного слайда
  - slide_text — заголовок + список выводов (bullets) в body shape
  - slide_chart — заголовок + bullets в текстовом блоке СЛЕВА + chart
    из Sheets СПРАВА. body shape резервируется под bullets, поверх
    добавляем chart по правой половине слайда.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass

from googleapiclient.discovery import Resource

from .filler import (
    delete_slide,
    duplicate_with_ids,
    fill_slide_by_roles,
    move_slide_to_position,
    replace_text_keep_style,
)
from .planner import Step
from .template_index import TemplateEntry, build_template_index, fetch_slides

log = logging.getLogger(__name__)


@dataclass
class ComposedStep:
    step: Step
    new_slide_id: str


def compose_report(
    *,
    slides_svc: Resource,
    sheets_svc: Resource,
    drive_svc: Resource,
    template_id: str,
    plan: list[Step],
    title: str,
    parent_folder_id: str | None,
) -> str:
    pres_id = _copy_template(drive_svc, template_id, title, parent_folder_id)
    log.info("copy → %s", pres_id)

    _, slides = fetch_slides(slides_svc, pres_id)
    index = build_template_index(slides)

    composed: list[ComposedStep] = []
    used_template_slide_ids: set[str] = set()

    for i, step in enumerate(plan):
        entry = index.get(step.role)
        if not entry:
            log.warning("step %s skipped (нет эталона в шаблоне)", step.role)
            continue

        if step.role == "slide_chart":
            new_id = _compose_chart_step(
                slides_svc, sheets_svc, drive_svc, pres_id, parent_folder_id,
                i, step, entry,
            )
        else:
            new_id = _compose_text_step(slides_svc, pres_id, i, step, entry)

        if new_id:
            composed.append(ComposedStep(step, new_id))
            used_template_slide_ids.add(entry.slide.object_id)

    # Удаляем все оригинальные слайды шаблона
    _, slides_now = fetch_slides(slides_svc, pres_id)
    new_ids = {c.new_slide_id for c in composed}
    delete_reqs = [delete_slide(s.object_id) for s in slides_now
                   if s.object_id not in new_ids]
    if delete_reqs:
        _exec(slides_svc, pres_id, delete_reqs)

    # Переупорядочить по плану
    move_reqs = []
    for pos, c in enumerate(composed):
        move_reqs.append(move_slide_to_position(c.new_slide_id, pos))
    if move_reqs:
        _exec(slides_svc, pres_id, move_reqs)

    return pres_id


def _compose_text_step(slides_svc, pres_id, idx, step, entry: TemplateEntry) -> str:
    new_slide_id = _safe_id(f"s{idx}_{step.role}")
    pairs = [(sh.object_id, _safe_id(f"{new_slide_id}_{role}"))
             for role, sh in entry.shape_by_role.items()]
    reqs = [duplicate_with_ids(entry.slide.object_id, new_slide_id, pairs)]

    # Подмена object_id в локальной копии entry
    new_shape_by_role = {}
    for role, shape in entry.shape_by_role.items():
        new_shape = type(shape)(
            object_id=_safe_id(f"{new_slide_id}_{role}"),
            x_in=shape.x_in, y_in=shape.y_in,
            w_in=shape.w_in, h_in=shape.h_in,
            text=shape.text,
        )
        new_shape_by_role[role] = new_shape
    fake = TemplateEntry(role=entry.role, slide=entry.slide,
                          shape_by_role=new_shape_by_role)

    # Подготовка данных под slots
    data = {}
    if "title" in step.data:
        data["title"] = step.data["title"]
    if "bullets" in step.data:
        data["body"] = "\n".join(f"• {b}" for b in step.data["bullets"])
    if "body" in step.data and "body" not in data:
        data["body"] = step.data["body"]
    # cover/final поля
    for k in ("name", "subtitle", "url"):
        if k in step.data:
            data[k] = step.data[k]

    reqs += fill_slide_by_roles(fake, data)
    _exec(slides_svc, pres_id, reqs)
    return new_slide_id


def _compose_chart_step(slides_svc, sheets_svc, drive_svc, pres_id, parent_folder_id,
                        idx, step, entry: TemplateEntry) -> str:
    """Заголовок + bullets в body + chart справа сверху (поверх существующей картинки)."""
    new_slide_id = _compose_text_step(slides_svc, pres_id, idx, step, entry)

    df = step.data.get("dataframe")
    if df is None or df.empty:
        return new_slide_id

    try:
        ssid, chart_id = _create_chart_in_sheets(
            sheets_svc, drive_svc, parent_folder_id,
            title=step.data.get("chart_title", ""),
            df=df,
            categories_col=step.data.get("categories_col"),
            series_cols=step.data.get("series_cols", []),
            chart_kind=step.data.get("chart_kind", "column"),
        )
    except Exception as e:  # noqa: BLE001
        log.warning("chart skipped on slide %s: %s", new_slide_id, e)
        return new_slide_id

    chart_obj_id = _safe_id(f"{new_slide_id}_chart")
    # Chart размещается справа — там, где в шаблоне стоит дефолтная картинка
    # (правая половина слайда). Layout "Заголовок + текст" имеет картинку
    # в районе x=7.78, y=2.57, ~5×4.2". Перекрываем её живым графиком.
    _exec(slides_svc, pres_id, [{
        "createSheetsChart": {
            "objectId": chart_obj_id,
            "spreadsheetId": ssid, "chartId": chart_id,
            "linkingMode": "LINKED",
            "elementProperties": _elem(new_slide_id, 7.5, 2.4, 5.4, 4.4),
        }
    }])
    return new_slide_id


def _copy_template(drive_svc, template_id, title, parent_folder_id) -> str:
    body = {"name": title}
    if parent_folder_id:
        body["parents"] = [parent_folder_id]
    copy = drive_svc.files().copy(fileId=template_id, body=body,
                                  supportsAllDrives=True).execute()
    return copy["id"]


def _create_chart_in_sheets(sheets_svc, drive_svc, parent_folder_id, *,
                            title, df, categories_col, series_cols, chart_kind) -> tuple[str, int]:
    ss = sheets_svc.spreadsheets().create(
        body={"properties": {"title": f"data — {title[:60]}"}}
    ).execute()
    ssid = ss["spreadsheetId"]
    sheet_id = ss["sheets"][0]["properties"]["sheetId"]

    if parent_folder_id:
        meta = drive_svc.files().get(fileId=ssid, fields="parents",
                                     supportsAllDrives=True).execute()
        drive_svc.files().update(
            fileId=ssid, addParents=parent_folder_id,
            removeParents=",".join(meta.get("parents", [])),
            supportsAllDrives=True,
        ).execute()

    cols = [categories_col] + [c for c in series_cols if c != categories_col]
    cols = [c for c in cols if c is not None and c in df.columns]
    out_df = df[cols].copy()
    values = [list(out_df.columns)] + (
        out_df.astype(object).where(out_df.notna(), "").values.tolist()
    )
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=ssid, range="A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    n_rows = len(out_df) + 1
    n_series = len(cols) - 1
    series = []
    for i in range(1, n_series + 1):
        series.append({
            "series": {"sourceRange": {"sources": [{
                "sheetId": sheet_id,
                "startRowIndex": 0, "endRowIndex": n_rows,
                "startColumnIndex": i, "endColumnIndex": i + 1,
            }]}},
            "targetAxis": "LEFT_AXIS",
        })

    if chart_kind == "pie":
        spec = {
            "title": title,
            "pieChart": {
                "legendPosition": "RIGHT_LEGEND",
                "threeDimensional": False,
                "domain": {"sourceRange": {"sources": [{
                    "sheetId": sheet_id,
                    "startRowIndex": 0, "endRowIndex": n_rows,
                    "startColumnIndex": 0, "endColumnIndex": 1,
                }]}},
                "series": {"sourceRange": {"sources": [{
                    "sheetId": sheet_id,
                    "startRowIndex": 0, "endRowIndex": n_rows,
                    "startColumnIndex": 1, "endColumnIndex": 2,
                }]}},
            },
        }
    else:
        chart_type = {"column": "COLUMN", "bar": "BAR", "line": "LINE"}.get(chart_kind, "COLUMN")
        spec = {
            "title": title,
            "basicChart": {
                "chartType": chart_type,
                "legendPosition": "BOTTOM_LEGEND",
                "axis": [
                    {"position": "BOTTOM_AXIS"},
                    {"position": "LEFT_AXIS"},
                ],
                "domains": [{
                    "domain": {"sourceRange": {"sources": [{
                        "sheetId": sheet_id,
                        "startRowIndex": 0, "endRowIndex": n_rows,
                        "startColumnIndex": 0, "endColumnIndex": 1,
                    }]}}
                }],
                "series": series,
                "headerCount": 1,
            },
        }

    resp = sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=ssid,
        body={"requests": [{
            "addChart": {"chart": {
                "spec": spec, "position": {"newSheet": True},
            }}
        }]},
    ).execute()
    chart_id = resp["replies"][0]["addChart"]["chart"]["chartId"]
    return ssid, chart_id


def _exec(slides_svc, pres_id, reqs):
    if not reqs:
        return None
    return slides_svc.presentations().batchUpdate(
        presentationId=pres_id, body={"requests": reqs}
    ).execute()


def _elem(slide_id, x_in, y_in, w_in, h_in):
    return {
        "pageObjectId": slide_id,
        "size": {
            "width": {"magnitude": w_in * 914400, "unit": "EMU"},
            "height": {"magnitude": h_in * 914400, "unit": "EMU"},
        },
        "transform": {
            "scaleX": 1, "scaleY": 1,
            "translateX": x_in * 914400, "translateY": y_in * 914400,
            "unit": "EMU",
        },
    }


def _safe_id(s: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9_-]", "_", s)[:50]
    return cleaned if len(cleaned) >= 5 else cleaned + "_xxxxx"[:5 - len(cleaned)]
