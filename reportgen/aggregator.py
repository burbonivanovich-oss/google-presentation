"""Агрегатор данных Контура: читает 4 xlsx из папки-источника, строит
набор aggregated DataFrame под все нужные разрезы отчёта.

Источники в папке `1DbAM_3ldLBVMGAxUVDtEeSniF4_97nTx`:
  * Царь свод все разрезы.xlsx  — план-факт по всем разрезам
  * Царь-данные по продажам.xlsx — транзакции 2025-2026
  * data (15).xlsx — реклама / ОФД помесячно по тарифам
  * data (16).xlsx — реклама / Маркет
  * data (17).xlsx — реклама / Бандл Маркет-ОФД
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import pandas as pd
from googleapiclient.http import MediaIoBaseDownload

from .drive import DriveClient

# Транзакционная таблица (Царь продажи)
TX_DATE = "Дата оплаты"
TX_REVENUE = "Оплата факт"
TX_QTY = "Кол-во факт"
TX_SEGMENT_TARIFF = "Сегментный тариф"
TX_BUSINESS_UNIT = "Бизнес-юнит"
TX_TARIFF = "Тариф"
TX_ONLINE = "Онлайн"
TX_ONLINE_TYPE = "Тип онлайна"
TX_SALE_METHOD = "Способ продажи"
TX_REGION = "Название региона клиента"
TX_MRC = "МРЦ"

# План-факт (Царь свод)
PF_MONTH = "Месяц"
PF_QTY_PLAN = "Кол-во план"
PF_REV_PLAN = "Оплата план"
PF_QTY_FACT = "Кол-во факт"
PF_REV_FACT = "Оплата факт"
PF_SEGMENT = "Сегмент плана"
PF_PROJECT = "Проект"

# Маппинг направлений на фильтр по «Сегментный тариф»
DIRECTIONS = {
    "Розница": ["Госсистемы для розницы"],
    "Общепит": ["Госсистемы для общепита"],
    "Кассы": ["Кассовики"],
    # ОФД определяем не по сегменту, а по продукту в Тарифе (содержит "ОФД")
    "ОФД": None,
}

MONTHS_RU = {1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
             5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
             9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"}


@dataclass
class DirectionSummary:
    """Метрики по одному направлению (Розница / Общепит / Кассы / ОФД)."""
    name: str
    revenue_cur: float = 0.0
    revenue_prev: float = 0.0
    revenue_yoy: float = 0.0
    qty_cur: int = 0
    qty_prev: int = 0
    avg_check_cur: float = 0.0
    top_tariffs: pd.DataFrame | None = None
    by_month: pd.DataFrame | None = None


@dataclass
class AggregateResult:
    period_label: str       # "Q1 2026"
    prev_period_label: str  # "Q4 2025"
    cur_q: int = 1
    cur_y: int = 2026
    prev_q: int = 4
    prev_y: int = 2025

    # Общие
    plan_fact_monthly: pd.DataFrame | None = None  # помесячно весь БЮ
    plan_fact_quarter: dict = field(default_factory=dict)  # сводные числа
    top_tariffs: pd.DataFrame | None = None
    by_region: pd.DataFrame | None = None
    by_online: pd.DataFrame | None = None

    # По направлениям
    directions: dict[str, DirectionSummary] = field(default_factory=dict)

    # Маркет
    market_monthly: pd.DataFrame | None = None
    market_top_tariffs: pd.DataFrame | None = None

    # Реклама (контекст) — словарь по продуктам
    ads_by_product: dict[str, pd.DataFrame] = field(default_factory=dict)


def aggregate(
    drive: DriveClient, sources_folder_id: str,
    current_period: str, previous_period: str,
) -> AggregateResult:
    cur_q, cur_y = _parse_period(current_period)
    prev_q, prev_y = _parse_period(previous_period)
    out = AggregateResult(
        period_label=f"Q{cur_q} {cur_y}",
        prev_period_label=f"Q{prev_q} {prev_y}",
        cur_q=cur_q, cur_y=cur_y, prev_q=prev_q, prev_y=prev_y,
    )

    # Скачиваем все источники
    files = _list_xlsx(drive, sources_folder_id)
    car_svod = _find(files, [r"Царь свод", r"свод все разрезы"])
    car_sales = _find(files, [r"Царь.*продаж", r"данные по продажам"])
    data15 = _find(files, [r"data \(15\)"])  # ОФД
    data16 = _find(files, [r"data \(16\)"])  # Маркет
    data17 = _find(files, [r"data \(17\)"])  # Бандл

    pf_df = _download_xlsx(drive, car_svod["id"]) if car_svod else None
    tx_df = _download_xlsx(drive, car_sales["id"]) if car_sales else None
    ofd_df = _download_xlsx(drive, data15["id"]) if data15 else None
    market_df = _download_xlsx(drive, data16["id"]) if data16 else None
    bundle_df = _download_xlsx(drive, data17["id"]) if data17 else None

    # Очистка
    if pf_df is not None:
        pf_df = _clean_plan_fact(pf_df)
    if tx_df is not None:
        tx_df = _clean_transactions(tx_df)

    # ── общие агрегации ───────────────────────────────────────
    if pf_df is not None and not pf_df.empty:
        out.plan_fact_monthly = _pf_monthly(pf_df, cur_q, cur_y)
        out.plan_fact_quarter = _pf_totals(pf_df, cur_q, cur_y, prev_q, prev_y)

    if tx_df is not None and not tx_df.empty:
        cur_tx = _filter_period(tx_df, cur_q, cur_y)
        prev_tx = _filter_period(tx_df, prev_q, prev_y)
        yoy_tx = _filter_period(tx_df, cur_q, cur_y - 1)
        out.top_tariffs = _top_tariffs_cmp(cur_tx, prev_tx, yoy_tx, out)
        out.by_region = _agg_top(cur_tx, TX_REGION, "Регион", 10)
        out.by_online = _agg_top(cur_tx, TX_ONLINE, "Канал", 0)

        # По направлениям
        for name, segs in DIRECTIONS.items():
            ds = _summarize_direction(cur_tx, prev_tx, yoy_tx, name, segs, out)
            out.directions[name] = ds

    # Маркет — из data (16)
    if market_df is not None:
        out.market_monthly = _market_monthly(market_df, cur_q, cur_y, prev_q, prev_y)
        out.market_top_tariffs = _market_top_tariffs(market_df, cur_q, cur_y)

    # Реклама
    for product, df in [("ОФД", ofd_df), ("Маркет", market_df), ("Бандл", bundle_df)]:
        if df is not None:
            out.ads_by_product[product] = _market_monthly(df, cur_q, cur_y, prev_q, prev_y)

    return out


# ── загрузка xlsx ─────────────────────────────────────────────────

def _list_xlsx(drive: DriveClient, folder_id: str) -> list[dict]:
    XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    resp = (
        drive._drive.files()  # noqa: SLF001
        .list(
            q=f"'{folder_id}' in parents and trashed = false and mimeType = '{XLSX}'",
            fields="files(id,name)", pageSize=200,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
    )
    return resp.get("files", [])


def _find(files: list[dict], patterns: list[str]) -> dict | None:
    for f in files:
        for p in patterns:
            if re.search(p, f["name"], re.IGNORECASE):
                return f
    return None


def _download_xlsx(drive: DriveClient, file_id: str) -> pd.DataFrame | None:
    req = drive._drive.files().get_media(fileId=file_id, supportsAllDrives=True)  # noqa: SLF001
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    try:
        return pd.read_excel(buf, sheet_name=0, header=None)
    except Exception:  # noqa: BLE001
        return None


# ── очистка ───────────────────────────────────────────────────────

def _parse_period(s: str) -> tuple[int, int]:
    m = re.match(r"\s*Q?(\d)\W+(\d{4})\s*", s)
    if not m:
        raise ValueError(f"bad period {s!r}")
    return int(m.group(1)), int(m.group(2))


def _to_num(x):
    if x is None or (isinstance(x, float) and x != x):
        return float("nan")
    s = str(x).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    s = s.replace("%", "").replace("₽", "")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def _promote_header(df: pd.DataFrame, anchor: str) -> pd.DataFrame:
    """Если первая строка — служебная ("Примененные фильтры...") и хедер
    лежит во 2-3-й строке, найдём его по якорю и поднимем."""
    for i in range(min(len(df), 5)):
        row = df.iloc[i]
        if any(str(v).strip() == anchor for v in row.values):
            df = df.copy()
            df.columns = [str(v).strip() if v == v else "" for v in row.values]
            return df.iloc[i + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.iloc[0].values]
    return df.iloc[1:].reset_index(drop=True)


def _clean_transactions(df: pd.DataFrame) -> pd.DataFrame:
    df = _promote_header(df, TX_DATE)
    out = df.copy()
    if TX_REVENUE in out.columns:
        out[TX_REVENUE] = out[TX_REVENUE].apply(_to_num)
    if TX_QTY in out.columns:
        out[TX_QTY] = out[TX_QTY].apply(_to_num)
    if TX_DATE in out.columns:
        dt = pd.to_datetime(out[TX_DATE], errors="coerce", format="mixed")
        out["_year"] = dt.dt.year.astype("Int64")
        out["_month"] = dt.dt.month.astype("Int64")
        out["_quarter"] = ((dt.dt.month - 1) // 3 + 1).astype("Int64")
    return out[out.get(TX_REVENUE, 0).notna()]


def _clean_plan_fact(df: pd.DataFrame) -> pd.DataFrame:
    df = _promote_header(df, PF_MONTH)
    out = df.copy()
    if PF_MONTH in out.columns:
        dt = pd.to_datetime(out[PF_MONTH], errors="coerce", format="mixed")
        out["_year"] = dt.dt.year.astype("Int64")
        out["_month"] = dt.dt.month.astype("Int64")
        out["_quarter"] = ((dt.dt.month - 1) // 3 + 1).astype("Int64")
    for col in (PF_QTY_PLAN, PF_REV_PLAN, PF_QTY_FACT, PF_REV_FACT):
        if col in out.columns:
            out[col] = out[col].apply(_to_num)
    return out


def _filter_period(df: pd.DataFrame, q: int, y: int) -> pd.DataFrame:
    return df[(df["_quarter"] == q) & (df["_year"] == y)]


# ── агрегации общие ───────────────────────────────────────────────

def _pf_monthly(pf: pd.DataFrame, q: int, y: int) -> pd.DataFrame | None:
    months = [(q - 1) * 3 + i for i in (1, 2, 3)]
    sel = pf[(pf["_year"] == y) & (pf["_month"].isin(months))]
    if sel.empty:
        return None
    g = sel.groupby("_month")[[PF_REV_PLAN, PF_REV_FACT, PF_QTY_PLAN, PF_QTY_FACT]].sum()
    g = g.reset_index()
    g["Месяц"] = g["_month"].apply(lambda m: f"{MONTHS_RU.get(int(m), str(m))} {y}")
    g["% выполнения"] = (g[PF_REV_FACT] / g[PF_REV_PLAN] * 100).round(0)
    return g[["Месяц", PF_REV_PLAN, PF_REV_FACT, "% выполнения", PF_QTY_PLAN, PF_QTY_FACT]]


def _pf_totals(pf: pd.DataFrame, cur_q: int, cur_y: int, prev_q: int, prev_y: int) -> dict:
    def total(q, y):
        sel = pf[(pf["_quarter"] == q) & (pf["_year"] == y)]
        return {
            "plan_rev": float(sel[PF_REV_PLAN].sum()) if PF_REV_PLAN in sel.columns else 0,
            "fact_rev": float(sel[PF_REV_FACT].sum()) if PF_REV_FACT in sel.columns else 0,
            "plan_qty": float(sel[PF_QTY_PLAN].sum()) if PF_QTY_PLAN in sel.columns else 0,
            "fact_qty": float(sel[PF_QTY_FACT].sum()) if PF_QTY_FACT in sel.columns else 0,
        }
    cur = total(cur_q, cur_y)
    prev = total(prev_q, prev_y)
    yoy = total(cur_q, cur_y - 1)
    pct = lambda f, p: (f / p * 100) if p else 0  # noqa: E731
    return {
        "cur": cur, "prev": prev, "yoy_base": yoy,
        "cur_pct": pct(cur["fact_rev"], cur["plan_rev"]),
        "qoq_delta": pct(cur["fact_rev"] - prev["fact_rev"], prev["fact_rev"]),
        "yoy_delta": pct(cur["fact_rev"] - yoy["fact_rev"], yoy["fact_rev"]),
    }


def _agg_top(df: pd.DataFrame, key: str, label: str, top_n: int) -> pd.DataFrame | None:
    if key not in df.columns:
        return None
    g = df.groupby(key)[TX_REVENUE].sum().sort_values(ascending=False)
    if g.empty:
        return None
    if top_n:
        g = g.head(top_n)
    return pd.DataFrame({label: g.index, "Выручка": g.values})


def _top_tariffs_cmp(cur: pd.DataFrame, prev: pd.DataFrame, yoy: pd.DataFrame,
                     out: AggregateResult) -> pd.DataFrame | None:
    if TX_TARIFF not in cur.columns:
        return None
    cur_g = cur.groupby(TX_TARIFF)[TX_REVENUE].sum().rename(out.period_label)
    prev_g = prev.groupby(TX_TARIFF)[TX_REVENUE].sum().rename(out.prev_period_label) \
        if not prev.empty else pd.Series(dtype=float, name=out.prev_period_label)
    yoy_g = yoy.groupby(TX_TARIFF)[TX_REVENUE].sum().rename(f"Q{out.cur_q} {out.cur_y - 1}") \
        if not yoy.empty else pd.Series(dtype=float, name=f"Q{out.cur_q} {out.cur_y - 1}")
    df = pd.concat([prev_g, yoy_g, cur_g], axis=1).fillna(0).reset_index()
    import numpy as np
    df["Δ QoQ %"] = ((df[out.period_label] - df[out.prev_period_label]) /
                     df[out.prev_period_label].replace(0, np.nan) * 100).round(0)
    df["Δ YoY %"] = ((df[out.period_label] - df[f"Q{out.cur_q} {out.cur_y - 1}"]) /
                     df[f"Q{out.cur_q} {out.cur_y - 1}"].replace(0, np.nan) * 100).round(0)
    df = df.rename(columns={TX_TARIFF: "Тариф"})
    return df.sort_values(out.period_label, ascending=False).head(20)


def _summarize_direction(cur: pd.DataFrame, prev: pd.DataFrame, yoy: pd.DataFrame,
                         name: str, segments: list[str] | None,
                         out: AggregateResult) -> DirectionSummary:
    if name == "ОФД":
        m = lambda df: df[df[TX_TARIFF].astype(str).str.contains("ОФД", na=False)]  # noqa: E731
    elif segments:
        m = lambda df: df[df[TX_SEGMENT_TARIFF].isin(segments)]  # noqa: E731
    else:
        m = lambda df: df  # noqa: E731

    cur_d = m(cur)
    prev_d = m(prev)
    yoy_d = m(yoy)

    ds = DirectionSummary(name=name)
    ds.revenue_cur = float(cur_d[TX_REVENUE].sum())
    ds.revenue_prev = float(prev_d[TX_REVENUE].sum())
    ds.revenue_yoy = float(yoy_d[TX_REVENUE].sum())
    ds.qty_cur = int(cur_d[TX_QTY].sum()) if TX_QTY in cur_d.columns else len(cur_d)
    ds.qty_prev = int(prev_d[TX_QTY].sum()) if TX_QTY in prev_d.columns else len(prev_d)
    ds.avg_check_cur = ds.revenue_cur / ds.qty_cur if ds.qty_cur else 0

    if TX_TARIFF in cur_d.columns:
        top = (cur_d.groupby(TX_TARIFF)[TX_REVENUE].sum()
               .sort_values(ascending=False).head(10))
        ds.top_tariffs = pd.DataFrame({"Тариф": top.index, "Выручка": top.values})
    if "_month" in cur_d.columns:
        gm = cur_d.groupby("_month")[TX_REVENUE].sum()
        ds.by_month = pd.DataFrame({
            "Месяц": [f"{MONTHS_RU.get(int(m), str(m))} {out.cur_y}" for m in gm.index],
            "Выручка": gm.values,
        })
    return ds


def _market_monthly(df: pd.DataFrame, cur_q, cur_y, prev_q, prev_y) -> pd.DataFrame | None:
    """data (15/16/17) формат: 'Период / Продукт учета / Тариф / Прайс' | ...
    Парсим период из первой колонки (e.g. 'янв 2025')."""
    if df is None or df.empty:
        return None
    df = _promote_header(df, "Оплаты")
    if "Оплаты" not in df.columns or "Выручка" not in df.columns:
        return None
    # первая колонка содержит период в формате 'мес год'
    period_col = df.columns[0]
    out = df.copy()
    out["Выручка"] = out["Выручка"].apply(_to_num)
    out["Оплаты"] = out["Оплаты"].apply(_to_num)
    # извлекаем (месяц, год) из строки 'янв 2025'
    parsed = out[period_col].astype(str).str.extract(r"(\w+)\s+(\d{4})")
    months_ru_short = {"янв": 1, "фев": 2, "мар": 3, "апр": 4, "май": 5, "июн": 6,
                       "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12}
    out["_month"] = parsed[0].str.lower().map(months_ru_short)
    out["_year"] = pd.to_numeric(parsed[1], errors="coerce")
    return out


def _market_top_tariffs(df: pd.DataFrame, cur_q: int, cur_y: int) -> pd.DataFrame | None:
    if df is None:
        return None
    df = _market_monthly(df, cur_q, cur_y, 0, 0)
    if df is None or "Тариф" not in df.columns:
        return None
    months = [(cur_q - 1) * 3 + i for i in (1, 2, 3)]
    sel = df[(df["_year"] == cur_y) & (df["_month"].isin(months))]
    if sel.empty:
        return None
    g = sel.groupby("Тариф")["Выручка"].sum().sort_values(ascending=False).head(20)
    return pd.DataFrame({"Тариф": g.index, "Выручка": g.values})
