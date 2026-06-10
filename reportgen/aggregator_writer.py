"""Записывает AggregateResult в новый Google Sheet с 10 листами.

Структура:
  1. Обзор — план/факт квартала + изменения QoQ/YoY
  2-5. Розница / Общепит / Кассы / ОФД — направления
  6. Маркет — тотал + топ тарифов
  7. Контекст реклама — помесячно по продуктам
  8. Помесячная динамика БЮ
  9. Топ-20 тарифов
  10. Источники
"""
from __future__ import annotations

import pandas as pd
from googleapiclient.discovery import Resource

from .aggregator import AggregateResult


def write_aggregate(sheets_svc: Resource, drive_svc: Resource,
                    rd: AggregateResult, *, title: str,
                    parent_folder_id: str | None) -> str:
    ss = sheets_svc.spreadsheets().create(
        body={"properties": {"title": title}}
    ).execute()
    ssid = ss["spreadsheetId"]
    default_sheet_id = ss["sheets"][0]["properties"]["sheetId"]

    if parent_folder_id:
        meta = drive_svc.files().get(fileId=ssid, fields="parents",
                                     supportsAllDrives=True).execute()
        drive_svc.files().update(
            fileId=ssid, addParents=parent_folder_id,
            removeParents=",".join(meta.get("parents", [])),
            supportsAllDrives=True,
        ).execute()

    pages: list[tuple[str, list[list]]] = []
    pages.append(("1. Обзор", _page_overview(rd)))
    for d_name in ("Розница", "Общепит", "Кассы", "ОФД"):
        ds = rd.directions.get(d_name)
        if ds:
            pages.append((f"2. {d_name}", _page_direction(ds, rd)))
    if rd.market_top_tariffs is not None:
        pages.append(("6. Маркет тарифы", _page_market(rd)))
    if rd.ads_by_product:
        pages.append(("7. Контекст реклама", _page_ads(rd)))
    if rd.plan_fact_monthly is not None:
        pages.append(("8. Помесячно БЮ", _df_to_rows(rd.plan_fact_monthly)))
    if rd.top_tariffs is not None:
        pages.append(("9. Топ-20 тарифов", _df_to_rows(rd.top_tariffs)))
    pages.append(("10. Источники", _page_sources(rd)))

    # Создаём все листы + переименовываем дефолтный → первая страница
    requests = [{"updateSheetProperties": {
        "properties": {"sheetId": default_sheet_id, "title": pages[0][0]},
        "fields": "title",
    }}]
    for name, _ in pages[1:]:
        requests.append({"addSheet": {"properties": {"title": name}}})
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=ssid, body={"requests": requests}).execute()

    # Заливаем значения
    data = []
    for name, rows in pages:
        data.append({
            "range": f"'{name}'!A1",
            "values": [_stringify(r) for r in rows],
        })
    sheets_svc.spreadsheets().values().batchUpdate(
        spreadsheetId=ssid,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()

    return ssid


# ── страницы ──────────────────────────────────────────────────────

def _page_overview(rd: AggregateResult) -> list[list]:
    rows = [[f"Сводка отчёта — {rd.period_label}"], []]
    pf = rd.plan_fact_quarter
    if pf:
        cur = pf["cur"]; prev = pf["prev"]; yoy = pf["yoy_base"]
        rows += [
            ["Метрика", rd.period_label, rd.prev_period_label,
             f"Q{rd.cur_q} {rd.cur_y - 1}", "QoQ %", "YoY %"],
            ["Выручка факт",
             _fmt(cur["fact_rev"]), _fmt(prev["fact_rev"]), _fmt(yoy["fact_rev"]),
             f"{pf['qoq_delta']:+.1f}%", f"{pf['yoy_delta']:+.1f}%"],
            ["Выручка план", _fmt(cur["plan_rev"]), "—", "—",
             f"% выполнения: {pf['cur_pct']:.0f}%", ""],
            ["Кол-во оплат",
             f"{int(cur['fact_qty']):,}".replace(",", " "),
             f"{int(prev['fact_qty']):,}".replace(",", " "),
             f"{int(yoy['fact_qty']):,}".replace(",", " "), "", ""],
            ["Ср.чек факт",
             _fmt(cur["fact_rev"] / cur["fact_qty"] if cur["fact_qty"] else 0),
             _fmt(prev["fact_rev"] / prev["fact_qty"] if prev["fact_qty"] else 0),
             "", "", ""],
            [],
        ]
    rows += [["Направление", "Выручка факт", f"vs {rd.prev_period_label}",
              "vs " + f"Q{rd.cur_q} {rd.cur_y - 1}", "Кол-во"]]
    for d in rd.directions.values():
        qoq = ((d.revenue_cur - d.revenue_prev) / d.revenue_prev * 100) if d.revenue_prev else 0
        yoy = ((d.revenue_cur - d.revenue_yoy) / d.revenue_yoy * 100) if d.revenue_yoy else 0
        rows.append([d.name, _fmt(d.revenue_cur),
                     f"{qoq:+.1f}%", f"{yoy:+.1f}%",
                     f"{d.qty_cur:,}".replace(",", " ")])
    return rows


def _page_direction(ds, rd: AggregateResult) -> list[list]:
    rows = [[f"Направление: {ds.name} — {rd.period_label}"], []]
    qoq = ((ds.revenue_cur - ds.revenue_prev) / ds.revenue_prev * 100) if ds.revenue_prev else 0
    yoy = ((ds.revenue_cur - ds.revenue_yoy) / ds.revenue_yoy * 100) if ds.revenue_yoy else 0
    rows += [
        ["Метрика", rd.period_label, rd.prev_period_label,
         f"Q{rd.cur_q} {rd.cur_y - 1}", "QoQ %", "YoY %"],
        ["Выручка", _fmt(ds.revenue_cur), _fmt(ds.revenue_prev), _fmt(ds.revenue_yoy),
         f"{qoq:+.1f}%", f"{yoy:+.1f}%"],
        ["Кол-во оплат",
         f"{ds.qty_cur:,}".replace(",", " "),
         f"{ds.qty_prev:,}".replace(",", " "), "", "", ""],
        ["Средний чек", _fmt(ds.avg_check_cur), "", "", "", ""],
        [],
    ]
    if ds.by_month is not None:
        rows += [["Помесячно"], ["Месяц", "Выручка"]]
        rows += _df_to_rows(ds.by_month, with_header=False)
        rows.append([])
    if ds.top_tariffs is not None:
        rows += [["Топ-10 тарифов"], ["Тариф", "Выручка"]]
        rows += _df_to_rows(ds.top_tariffs, with_header=False)
    return rows


def _page_market(rd: AggregateResult) -> list[list]:
    rows = [[f"Маркет — {rd.period_label}"], [],
            ["Топ-20 тарифов Маркета (рекламные продажи)"]]
    rows += _df_to_rows(rd.market_top_tariffs)
    return rows


def _page_ads(rd: AggregateResult) -> list[list]:
    rows = [[f"Контекст реклама — {rd.period_label}"], []]
    for product, df in rd.ads_by_product.items():
        rows += [[f"Продукт: {product}"]]
        if df is None or df.empty:
            rows.append(["(нет данных)"])
            continue
        # помесячный свод выручки и оплат
        if "_year" in df.columns:
            sel = df[df["_year"] == rd.cur_y]
            g = sel.groupby("_month")[["Выручка", "Оплаты"]].sum().reset_index()
            g["Месяц"] = g["_month"].apply(lambda m: f"месяц {int(m)}")
            rows.append(["Месяц", "Выручка", "Оплаты"])
            for _, r in g.iterrows():
                rows.append([r["Месяц"], _fmt(r["Выручка"]), f"{int(r['Оплаты'])}"])
        rows.append([])
    return rows


def _page_sources(rd: AggregateResult) -> list[list]:
    return [
        ["Источники данных"], [],
        ["Лист", "Источник", "Фильтр / преобразование"],
        ["1. Обзор", "Царь свод + Царь продажи", f"Период = {rd.period_label}"],
        ["2-5. Направления", "Царь продажи",
         "Розница: Сегментный тариф=Госсистемы для розницы; "
         "Общепит: =Госсистемы для общепита; Кассы: =Кассовики; "
         "ОФД: Тариф содержит «ОФД»"],
        ["6. Маркет тарифы", "data (16).xlsx", "Все рекламные продажи продукта Маркет"],
        ["7. Контекст реклама", "data (15) ОФД / data (16) Маркет / data (17) Бандл",
         "Помесячная свёртка Выручка и Оплаты"],
        ["8. Помесячно БЮ", "Царь свод все разрезы.xlsx",
         f"Помесячно за квартал {rd.period_label}"],
        ["9. Топ-20 тарифов", "Царь продажи", f"Group by Тариф, Q1-{rd.cur_y}"],
        [],
        ["Что НЕ автоматизировано:"],
        ["• Тексты «Где отстаём» — нужны от аналитика"],
        ["• CPL / CPO / показы — на скриншотах PNG, OCR ненадёжен"],
        ["• Проекты и задачи Q2 — план команды, не из данных"],
        ["• Поисковый спрос Wordstat — нужна отдельная выгрузка"],
    ]


# ── утилиты ───────────────────────────────────────────────────────

def _df_to_rows(df: pd.DataFrame, with_header: bool = True) -> list[list]:
    rows: list[list] = []
    if with_header:
        rows.append(list(df.columns))
    for _, r in df.iterrows():
        rows.append([_cell(v) for v in r.values])
    return rows


def _cell(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if isinstance(v, (int,)):
        return v
    if isinstance(v, float):
        if abs(v - round(v)) < 0.001:
            return int(round(v))
        return round(v, 2)
    return str(v)


def _stringify(row: list) -> list:
    return [("" if v is None else v) for v in row]


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and v != v):
        return "—"
    v = float(v)
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.2f} млн ₽"
    if abs(v) >= 1_000:
        return f"{v/1_000:.0f} тыс ₽"
    return f"{v:.0f} ₽"
