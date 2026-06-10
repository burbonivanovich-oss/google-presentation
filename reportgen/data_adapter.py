"""Адаптер данных Контура — мультиисточник.

Скрипт работает с тремя источниками в папке отчёта:
  * data (14) — транзакции 2026 (по строке на оплату)
  * data (13) — помесячный план/факт по всему БЮ
  * Данные — сводка CAC / CPL / Конверсия по кварталам

Адаптер сам ищет нужные файлы в Drive по name pattern (без жёсткого
указания ID в конфиге), парсит их, агрегирует под слайды Петровой.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

from .drive import MIME_SHEET, DriveClient
from .sheets import SheetsClient

# ── транзакционная таблица (data (14)) ─────────────────────────────
COL_QUARTER = "Квартал"
COL_MONTH = "Месяц"
COL_DATE = "Дата оплаты"
COL_REVENUE = "Оплата факт"
COL_BUSINESS_UNIT = "Бизнес-юнит"
COL_SEGMENT_TARIFF = "Сегментный тариф"
COL_TARIFF = "Тариф"
COL_SEGMENT_PLAN = "Сегмент плана"
COL_ONLINE = "Онлайн"
COL_ONLINE_TYPE = "Тип онлайна"
COL_SALE_METHOD = "Способ продажи"
COL_REGION = "Название региона клиента"

# ── план/факт (data (13)) ───────────────────────────────────────────
PF_MONTH = "Месяц"
PF_QTY_PLAN = "Кол-во план"
PF_REV_PLAN = "Оплата план"
PF_QTY_FACT = "Кол-во факт"
PF_REV_FACT = "Оплата факт"

MONTHS_RU = {1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
             5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
             9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"}


@dataclass
class MultiSourceData:
    period_label: str       # "Q1 2026"
    prev_period_label: str  # "Q4 2025"
    cur_q: int = 1
    cur_y: int = 2026
    prev_q: int = 4
    prev_y: int = 2025

    # Сырые таблицы
    transactions: pd.DataFrame | None = None  # data (14)
    plan_fact: pd.DataFrame | None = None     # data (13)
    cac_cpl: pd.DataFrame | None = None       # Данные

    # Агрегации
    pf_monthly: pd.DataFrame | None = None    # помесячно за квартал
    pf_quarter: dict = field(default_factory=dict)  # суммарно за квартал
    by_product: pd.DataFrame | None = None
    by_month: pd.DataFrame | None = None
    top_tariffs: pd.DataFrame | None = None
    online_vs_offline: pd.DataFrame | None = None
    by_region: pd.DataFrame | None = None


def load_and_adapt(
    drive: DriveClient, sheets: SheetsClient, folder_id: str,
    current_period: str, previous_period: str,
) -> MultiSourceData:
    cur_q, cur_y = _parse_period(current_period)
    prev_q, prev_y = _parse_period(previous_period)
    rd = MultiSourceData(
        period_label=f"Q{cur_q} {cur_y}",
        prev_period_label=f"Q{prev_q} {prev_y}",
        cur_q=cur_q, cur_y=cur_y, prev_q=prev_q, prev_y=prev_y,
    )

    # 1) Транзакции — data (14) — первый Google Sheet с "(14)" в имени
    tx_file = _find_first(drive, folder_id, [r"data \(14\)", r"data\s*\(?14\)?"])
    if tx_file:
        try:
            rd.transactions = _clean_transactions(
                sheets.read_table(tx_file["id"], "Sheet1!A1:AZ100000")
            )
        except Exception:  # noqa: BLE001
            pass

    # 2) План/факт помесячно — data (13)
    pf_file = _find_first(drive, folder_id, [r"data \(13\)", r"data\s*\(?13\)?"])
    if pf_file:
        try:
            rd.plan_fact = _clean_plan_fact(
                sheets.read_table(pf_file["id"], "Sheet1!A1:M100")
            )
        except Exception:  # noqa: BLE001
            pass

    # 3) CAC/CPL — Данные
    cc_file = _find_first(drive, folder_id, [r"^Данные$", r"^Данные "])
    if cc_file:
        try:
            rd.cac_cpl = sheets.read_table(cc_file["id"], "Лист1!A1:E20")
        except Exception:  # noqa: BLE001
            pass

    # Агрегации по транзакциям
    if rd.transactions is not None and not rd.transactions.empty:
        cur_tx = rd.transactions[
            (rd.transactions["_quarter"] == cur_q) & (rd.transactions["_year"] == cur_y)
        ]
        prev_tx = rd.transactions[
            (rd.transactions["_quarter"] == prev_q) & (rd.transactions["_year"] == prev_y)
        ]
        rd.by_product = _agg_by(cur_tx, prev_tx, COL_BUSINESS_UNIT, rd)
        rd.top_tariffs = _agg_by(cur_tx, prev_tx, COL_TARIFF, rd, top_n=10)
        rd.online_vs_offline = _agg_by(cur_tx, prev_tx, COL_ONLINE, rd)
        rd.by_region = _agg_by(cur_tx, pd.DataFrame(), COL_REGION, rd, top_n=10)
        rd.by_month = _agg_by_month(cur_tx, cur_q, cur_y)

    # Агрегации по плану-факту
    if rd.plan_fact is not None and not rd.plan_fact.empty:
        rd.pf_monthly = _pf_quarter_months(rd.plan_fact, cur_q, cur_y)
        rd.pf_quarter = _pf_quarter_totals(rd.plan_fact, cur_q, cur_y, prev_q, prev_y)

    return rd


# ── вспомогательное ────────────────────────────────────────────────

def _find_first(drive: DriveClient, folder_id: str, patterns: list[str]) -> dict | None:
    resp = (
        drive._drive.files()  # noqa: SLF001
        .list(
            q=f"'{folder_id}' in parents and trashed = false and mimeType = '{MIME_SHEET}'",
            fields="files(id,name)", pageSize=200,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
    )
    for f in resp.get("files", []):
        for p in patterns:
            if re.search(p, f["name"], re.IGNORECASE):
                return f
    return None


def _parse_period(s: str) -> tuple[int, int]:
    m = re.match(r"\s*Q?(\d)\W+(\d{4})\s*", s)
    if not m:
        raise ValueError(f"Не разобрать период: {s!r}")
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


def _clean_transactions(df: pd.DataFrame) -> pd.DataFrame:
    # У data (14) первая строка может быть "Примененные фильтры: ..." — тогда
    # реальный header на 2-й строке. Поищем строку с "Дата оплаты" в столбцах.
    if COL_DATE not in df.columns:
        for i, row in df.iterrows():
            if any(str(v).strip() == COL_DATE for v in row.values):
                df.columns = row.values
                df = df.iloc[i + 1:].reset_index(drop=True)
                break
    out = df.copy()
    if COL_REVENUE in out.columns:
        out[COL_REVENUE] = out[COL_REVENUE].apply(_to_num)
    if COL_DATE in out.columns:
        dt = pd.to_datetime(out[COL_DATE], errors="coerce", format="mixed")
        out["_year"] = dt.dt.year.astype("Int64")
        out["_month"] = dt.dt.month.astype("Int64")
        out["_quarter"] = ((dt.dt.month - 1) // 3 + 1).astype("Int64")
    if COL_REVENUE in out.columns:
        out = out[out[COL_REVENUE].notna()]
    return out


def _clean_plan_fact(df: pd.DataFrame) -> pd.DataFrame:
    if PF_MONTH not in df.columns:
        for i, row in df.iterrows():
            if any(str(v).strip() == PF_MONTH for v in row.values):
                df.columns = row.values
                df = df.iloc[i + 1:].reset_index(drop=True)
                break
    out = df.copy()
    # Даты вида "1/1/2026"
    dt = pd.to_datetime(out[PF_MONTH], errors="coerce", format="mixed")
    out["_year"] = dt.dt.year.astype("Int64")
    out["_month"] = dt.dt.month.astype("Int64")
    out["_quarter"] = ((dt.dt.month - 1) // 3 + 1).astype("Int64")
    for col in (PF_QTY_PLAN, PF_REV_PLAN, PF_QTY_FACT, PF_REV_FACT):
        if col in out.columns:
            out[col] = out[col].apply(_to_num)
    return out


def _agg_by(cur: pd.DataFrame, prev: pd.DataFrame, key: str,
            rd: MultiSourceData, top_n: int = 0) -> pd.DataFrame | None:
    if cur is None or cur.empty or key not in cur.columns:
        return None
    cur_g = cur.groupby(key)[COL_REVENUE].sum().rename(rd.period_label)
    if prev is not None and not prev.empty and key in prev.columns:
        prev_g = prev.groupby(key)[COL_REVENUE].sum().rename(rd.prev_period_label)
        out = pd.concat([prev_g, cur_g], axis=1).fillna(0).reset_index()
    else:
        out = cur_g.reset_index()
    out = out.rename(columns={key: key})
    out = out.sort_values(rd.period_label, ascending=False)
    if top_n:
        out = out.head(top_n)
    return out


def _agg_by_month(cur: pd.DataFrame, quarter: int, year: int) -> pd.DataFrame | None:
    if cur is None or cur.empty or "_month" not in cur.columns:
        return None
    months_in_q = [(quarter - 1) * 3 + i for i in (1, 2, 3)]
    g = cur[cur["_month"].isin(months_in_q)].groupby("_month")[COL_REVENUE].sum()
    if g.empty:
        return None
    return pd.DataFrame({
        "Месяц": [f"{MONTHS_RU.get(int(m), str(m))} {year}" for m in g.index],
        "Выручка": g.values,
    })


def _pf_quarter_months(pf: pd.DataFrame, quarter: int, year: int) -> pd.DataFrame | None:
    months_in_q = [(quarter - 1) * 3 + i for i in (1, 2, 3)]
    sel = pf[(pf["_quarter"] == quarter) & (pf["_year"] == year)]
    if sel.empty:
        return None
    out = pd.DataFrame({
        "Месяц": [MONTHS_RU.get(int(m), str(m)) for m in sel["_month"]],
        "План, ₽": sel[PF_REV_PLAN].values,
        "Факт, ₽": sel[PF_REV_FACT].values,
    })
    return out


def _pf_quarter_totals(pf: pd.DataFrame, cur_q: int, cur_y: int,
                       prev_q: int, prev_y: int) -> dict:
    def total(q: int, y: int) -> dict:
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
