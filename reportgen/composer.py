"""Composer: оркестратор сборки отчёта по plan + шаблону Контура.

Алгоритм:
  1. copy(template) → новая презентация
  2. read структуру → TemplateEntry для каждой роли
  3. для каждого Step плана:
       - если роль есть в индексе → duplicateObject(template_slide) с
         фиксированными object_ids + filler заполняет shape'ы
       - если роль table/chart → берём blank-эталон, дублируем и сверху
         createTable / createSheetsChart с данными
  4. удаляем ВСЕ оригинальные слайды шаблона (наши копии остаются)
  5. переупорядочиваем оставшиеся в порядке плана

Возвращаем ID готовой презентации.
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
    missing = [s.role for s in plan
               if s.role not in index and s.role not in ("table", "chart_column")]
    if missing:
        log.warning("в шаблоне нет эталонов для ролей: %s", missing)

    text_requests: list[dict] = []
    composed: list[ComposedStep] = []
    used_template_slide_ids: set[str] = set()

    # двухпроход: сначала text-роли в одном batchUpdate (быстро и атомарно),
    # потом table/chart — каждый отдельным батчем (нужны промежуточные ID).
    for i, step in enumerate(plan):
        if step.role in ("table", "chart_column"):
            continue
        entry = index.get(step.role)
        if not entry:
            log.warning("step %s skipped (нет эталона)", step.role)
            continue
        new_slide_id, fill_reqs = _build_text_step(i, step, entry)
        text_requests.append(_dup_request(entry, new_slide_id))
        text_requests.extend(fill_reqs)
        composed.append(ComposedStep(step, new_slide_id))
        used_template_slide_ids.add(entry.slide.object_id)

    if text_requests:
        _exec(slides_svc, pres_id, text_requests)

    # Графики/таблицы — отдельно
    blank_entry = index.get("blank")
    for i, step in enumerate(plan):
        if step.role == "table":
            new_slide_id = _add_table_step(slides_svc, pres_id, i, step, blank_entry, used_template_slide_ids)
            if new_slide_id:
                composed.append(ComposedStep(step, new_slide_id))
        elif step.role == "chart_column":
            new_slide_id = _add_chart_step(
                slides_svc, sheets_svc, drive_svc, pres_id, parent_folder_id,
                i, step, blank_entry, used_template_slide_ids,
            )
            if new_slide_id:
                composed.append(ComposedStep(step, new_slide_id))

    # Удаляем все оригинальные слайды шаблона
    _, slides_now = fetch_slides(slides_svc, pres_id)
    new_ids = {c.new_slide_id for c in composed}
    delete_reqs = [delete_slide(s.object_id) for s in slides_now
                   if s.object_id not in new_ids]
    if delete_reqs:
        _exec(slides_svc, pres_id, delete_reqs)

    # Переупорядочить в порядке плана
    move_reqs = []
    plan_order = {id(c.step): pos for pos, c in enumerate(composed)}
    for c in composed:
        move_reqs.append(move_slide_to_position(c.new_slide_id, plan_order[id(c.step)]))
    if move_reqs:
        _exec(slides_svc, pres_id, move_reqs)

    return pres_id


def _copy_template(drive_svc, template_id, title, parent_folder_id) -> str:
    body = {"name": title}
    if parent_folder_id:
        body["parents"] = [parent_folder_id]
    copy = drive_svc.files().copy(fileId=template_id, body=body,
                                  supportsAllDrives=True).execute()
    return copy["id"]


def _dup_request(entry: TemplateEntry, new_slide_id: str) -> dict:
    pairs = []
    for role, shape in entry.shape_by_role.items():
        pairs.append((shape.object_id, _safe_id(f"{new_slide_id}_{role}")))
    return duplicate_with_ids(entry.slide.object_id, new_slide_id, pairs)


def _build_text_step(idx: int, step: Step, entry: TemplateEntry) -> tuple[str, list[dict]]:
    new_slide_id = _safe_id(f"s{idx}_{step.role}")
    # подменяем shape.object_id на новые предсказуемые ID для filler
    new_shape_by_role = {}
    for role, shape in entry.shape_by_role.items():
        new_shape = type(shape)(
            object_id=_safe_id(f"{new_slide_id}_{role}"),
            x_in=shape.x_in, y_in=shape.y_in,
            w_in=shape.w_in, h_in=shape.h_in,
            text=shape.text,
        )
        new_shape_by_role[role] = new_shape
    fake_entry = TemplateEntry(role=entry.role, slide=entry.slide,
                                shape_by_role=new_shape_by_role)
    return new_slide_id, fill_slide_by_roles(fake_entry, step.data)


def _add_table_step(slides_svc, pres_id, idx, step, blank_entry, used) -> str | None:
    if not blank_entry:
        log.warning("table step skipped — нет blank эталона")
        return None
    new_slide_id = _safe_id(f"s{idx}_table")
    df = step.data.get("dataframe")
    if df is None or df.empty:
        return None

    used.add(blank_entry.slide.object_id)
    _exec(slides_svc, pres_id, [_dup_request(blank_entry, new_slide_id)])

    # Очищаем blank-слайд: все его pageElements (которые остались после dup)
    # нам не нужны (это title-плейсхолдер шаблона). Удалим их по факту.
    _, slides_now = fetch_slides(slides_svc, pres_id)
    target = next((s for s in slides_now if s.object_id == new_slide_id), None)
    if target:
        cleanup = [{"deleteObject": {"objectId": sh.object_id}} for sh in target.shapes]
        if cleanup:
            _exec(slides_svc, pres_id, cleanup)

    # Заголовок + таблица
    title_id = _safe_id(f"{new_slide_id}_title")
    table_id = _safe_id(f"{new_slide_id}_tbl")
    title_text = step.data.get("title", "")
    rows = len(df) + 1
    cols = len(df.columns)

    reqs = []
    reqs.append({"createShape": {
        "objectId": title_id, "shapeType": "TEXT_BOX",
        "elementProperties": _elem(new_slide_id, 0.64, 0.64, 12.05, 1.0),
    }})
    if title_text:
        reqs.append({"insertText": {"objectId": title_id, "text": title_text}})
        reqs.append({"updateTextStyle": {
            "objectId": title_id, "textRange": {"type": "ALL"},
            "style": {"fontSize": _pt(28), "bold": True,
                      "fontFamily": "Montserrat"},
            "fields": "fontSize,bold,fontFamily",
        }})

    reqs.append({"createTable": {
        "objectId": table_id, "rows": rows, "columns": cols,
        "elementProperties": _elem(new_slide_id, 0.64, 1.95, 12.05, 5.0),
    }})
    _exec(slides_svc, pres_id, reqs)

    # Наполнение ячеек
    fill = []
    headers = [str(c) for c in df.columns]
    for c, h in enumerate(headers):
        fill.append({"insertText": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 0, "columnIndex": c},
            "text": h}})
        fill.append({"updateTextStyle": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 0, "columnIndex": c},
            "textRange": {"type": "ALL"},
            "style": {"bold": True, "fontFamily": "Montserrat",
                      "fontSize": _pt(11),
                      "foregroundColor": _color(1, 1, 1)},
            "fields": "bold,fontFamily,fontSize,foregroundColor"}})
    fill.append({"updateTableCellProperties": {
        "objectId": table_id,
        "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0},
                       "rowSpan": 1, "columnSpan": cols},
        "tableCellProperties": {
            "tableCellBackgroundFill": {"solidFill": {"color": _color(0.21, 0.42, 0.95)}}},
        "fields": "tableCellBackgroundFill.solidFill.color"}})

    for r, row in enumerate(df.itertuples(index=False), start=1):
        for c, val in enumerate(row):
            s = _format_cell(val)
            fill.append({"insertText": {
                "objectId": table_id,
                "cellLocation": {"rowIndex": r, "columnIndex": c},
                "text": s}})
            fill.append({"updateTextStyle": {
                "objectId": table_id,
                "cellLocation": {"rowIndex": r, "columnIndex": c},
                "textRange": {"type": "ALL"},
                "style": {"fontFamily": "Montserrat", "fontSize": _pt(11)},
                "fields": "fontFamily,fontSize"}})
    _exec(slides_svc, pres_id, fill)
    return new_slide_id


def _add_chart_step(slides_svc, sheets_svc, drive_svc, pres_id, parent_folder_id,
                    idx, step, blank_entry, used) -> str | None:
    if not blank_entry:
        log.warning("chart step skipped — нет blank эталона")
        return None
    df = step.data.get("dataframe")
    if df is None or df.empty:
        return None
    categories_col = step.data.get("categories_col", df.columns[0])
    series_cols = step.data.get("series_cols", [c for c in df.columns if c != categories_col])

    ssid, chart_id = _create_chart_in_sheets(
        sheets_svc, drive_svc, parent_folder_id,
        title=step.data.get("title", "Chart"),
        df=df, categories_col=categories_col, series_cols=series_cols,
    )

    new_slide_id = _safe_id(f"s{idx}_chart")
    used.add(blank_entry.slide.object_id)
    _exec(slides_svc, pres_id, [_dup_request(blank_entry, new_slide_id)])

    # Чистим всё на дубликате
    _, slides_now = fetch_slides(slides_svc, pres_id)
    target = next((s for s in slides_now if s.object_id == new_slide_id), None)
    if target:
        cleanup = [{"deleteObject": {"objectId": sh.object_id}} for sh in target.shapes]
        if cleanup:
            _exec(slides_svc, pres_id, cleanup)

    title_id = _safe_id(f"{new_slide_id}_title")
    chart_obj_id = _safe_id(f"{new_slide_id}_chart")
    reqs = [
        {"createShape": {
            "objectId": title_id, "shapeType": "TEXT_BOX",
            "elementProperties": _elem(new_slide_id, 0.64, 0.64, 12.05, 1.0)}},
        {"insertText": {"objectId": title_id, "text": step.data.get("title", "")}},
        {"updateTextStyle": {
            "objectId": title_id, "textRange": {"type": "ALL"},
            "style": {"fontSize": _pt(28), "bold": True,
                      "fontFamily": "Montserrat"},
            "fields": "fontSize,bold,fontFamily"}},
        {"createSheetsChart": {
            "objectId": chart_obj_id,
            "spreadsheetId": ssid, "chartId": chart_id,
            "linkingMode": "LINKED",
            "elementProperties": _elem(new_slide_id, 0.64, 1.95, 12.05, 5.0)}},
    ]
    _exec(slides_svc, pres_id, reqs)
    return new_slide_id


def _create_chart_in_sheets(sheets_svc, drive_svc, parent_folder_id, *,
                            title, df, categories_col, series_cols) -> tuple[str, int]:
    ss = sheets_svc.spreadsheets().create(
        body={"properties": {"title": f"data — {title}"}}
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

    # лояльный порядок колонок: category первая, серии — далее
    cols = [categories_col] + list(series_cols)
    out_df = df[cols].copy()
    values = [list(out_df.columns)] + out_df.astype(object).where(out_df.notna(), "").values.tolist()
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=ssid, range="A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    n_rows = len(out_df) + 1
    series = []
    for i, _ in enumerate(series_cols, start=1):
        series.append({
            "series": {"sourceRange": {"sources": [{
                "sheetId": sheet_id,
                "startRowIndex": 0, "endRowIndex": n_rows,
                "startColumnIndex": i, "endColumnIndex": i + 1,
            }]}},
            "targetAxis": "LEFT_AXIS",
        })

    resp = sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=ssid,
        body={"requests": [{
            "addChart": {"chart": {
                "spec": {
                    "title": title,
                    "basicChart": {
                        "chartType": "COLUMN",
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
                },
                "position": {"newSheet": True},
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


def _pt(v):
    return {"magnitude": v, "unit": "PT"}


def _color(r, g, b):
    return {"opaqueColor": {"rgbColor": {"red": r, "green": g, "blue": b}}}


def _format_cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        if abs(v - round(v)) < 0.001:
            return f"{int(round(v)):,}".replace(",", " ")
        return f"{v:,.1f}".replace(",", " ")
    return str(v)


def _safe_id(s: str) -> str:
    # Slides objectId: a-zA-Z0-9_- , 5-50 символов, уникальный в презентации
    cleaned = re.sub(r"[^a-zA-Z0-9_-]", "_", s)[:50]
    return cleaned if len(cleaned) >= 5 else cleaned + "_xxxxx"[:5 - len(cleaned)]
