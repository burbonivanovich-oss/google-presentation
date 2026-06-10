"""Адаптер транзакционных данных Контура (sheet `data (24)`).

На входе — sheet с двумя секциями:
  1. «Деньги Медийка» — сводная таблица (мы её игнорируем).
  2. «Сырые данные» — транзакционная таблица оплат, по строке на сделку.

На выходе — набор aggregated DataFrame и метрик для planner-а:
  - totals: суммарная выручка/оплаты/лиды за период
  - by_product: выручка по продуктам (бизнес-юнит / сегментный тариф)
  - by_month: помесячная выручка
  - top_tariffs: топ-N тарифов
  - online_vs_offline: разбивка по способу продажи (онлайн/офлайн)
  - by_region: топ регионов
  - yoy: сравнение текущий vs прошлогодний квартал
  - qoq: сравнение текущий vs прошлый квартал
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

# Названия колонок в транзакционной таблице (как в data (24))
COL_QUARTER = "Квартал"
COL_MONTH = "Месяц"
COL_DATE = "Дата оплаты"
COL_REVENUE = "Оплата факт"
COL_BUSINESS_UNIT = "Бизнес-юнит"
COL_SEGMENT_TARIFF = "Сегментный тариф"
COL_TARIFF = "Тариф"
COL_SEGMENT_PLAN = "Сегмент плана"
COL_ONLINE_TYPE = "Тип онлайна"
COL_ONLINE = "Онлайн"
COL_SALE_METHOD = "Способ продажи"
COL_REGION = "Название региона клиента"
COL_MRC = "МРЦ"

MONTHS_RU = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}


@dataclass
class ReportData:
    """Все агрегации готовые к рендерингу в слайды."""
    period_label: str            # "Q1 2026"
    prev_period_label: str       # "Q4 2025"
    totals: dict[str, float] = field(default_factory=dict)
    totals_prev: dict[str, float] = field(default_factory=dict)
    by_product: pd.DataFrame | None = None       # ["Продукт", cur, prev]
    by_month: pd.DataFrame | None = None         # ["Месяц", "Выручка"]
    top_tariffs: pd.DataFrame | None = None      # ["Тариф", cur, prev, "Δ%"]
    online_vs_offline: pd.DataFrame | None = None # ["Канал", cur, prev]
    by_region: pd.DataFrame | None = None        # топ-10 регионов
    by_sale_method: pd.DataFrame | None = None   # ["Способ", "Выручка"]


def adapt(
    raw: pd.DataFrame,
    *,
    current_period: str,   # "Q1-2026"
    previous_period: str,  # "Q4-2025"
) -> ReportData:
    """Главная точка входа. raw — то, что вернул SheetsClient.read_table."""
    if raw is None or raw.empty:
        return ReportData(period_label=current_period, prev_period_label=previous_period)

    df = _clean(raw)
    cur_q, cur_y = _parse_period(current_period)   # (1, 2026)
    prev_q, prev_y = _parse_period(previous_period)

    cur = df[(df["_quarter"] == cur_q) & (df["_year"] == cur_y)]
    prev = df[(df["_quarter"] == prev_q) & (df["_year"] == prev_y)]

    rd = ReportData(
        period_label=f"Q{cur_q} {cur_y}",
        prev_period_label=f"Q{prev_q} {prev_y}",
        totals=_totals(cur),
        totals_prev=_totals(prev),
        by_product=_by_product(cur, prev, rd_labels=(f"Q{prev_q} {prev_y}", f"Q{cur_q} {cur_y}")),
        by_month=_by_month(cur, cur_q, cur_y),
        top_tariffs=_top_tariffs(cur, prev, rd_labels=(f"Q{prev_q} {prev_y}", f"Q{cur_q} {cur_y}")),
        online_vs_offline=_online_vs_offline(cur, prev, rd_labels=(f"Q{prev_q} {prev_y}", f"Q{cur_q} {cur_y}")),
        by_region=_by_region(cur),
        by_sale_method=_by_sale_method(cur),
    )
    return rd


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    """Чистим типы и нормализуем числовые колонки."""
    out = df.copy()
    # Выручка как float
    if COL_REVENUE in out.columns:
        out[COL_REVENUE] = out[COL_REVENUE].apply(_to_num)
    # Месяц как int
    if COL_MONTH in out.columns:
        out["_month"] = pd.to_numeric(out[COL_MONTH], errors="coerce").astype("Int64")
    # Квартал как int
    if COL_QUARTER in out.columns:
        out["_quarter"] = pd.to_numeric(out[COL_QUARTER], errors="coerce").astype("Int64")
    # Год — из даты или из «Дата оплаты»
    if COL_DATE in out.columns:
        out["_year"] = pd.to_datetime(out[COL_DATE], errors="coerce",
                                       format="mixed", dayfirst=False).dt.year.astype("Int64")
    else:
        out["_year"] = pd.NA
    # фильтруем мусорные строки (без выручки и без квартала)
    if COL_REVENUE in out.columns:
        out = out[out[COL_REVENUE].notna()]
    return out


def _parse_period(s: str) -> tuple[int, int]:
    """'Q1-2026' → (1, 2026). Принимает 'Q1 2026', 'Q1.2026' и т.п."""
    import re
    m = re.match(r"\s*Q?(\d)\W+(\d{4})\s*", s)
    if not m:
        raise ValueError(f"Не разобрать период: {s!r}. Ожидаю формат Q1-2026.")
    return int(m.group(1)), int(m.group(2))


def _to_num(x):
    if x is None or (isinstance(x, float) and x != x):
        return float("nan")
    s = str(x).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def _totals(df: pd.DataFrame) -> dict[str, float]:
    if df.empty:
        return {"revenue": 0.0, "deals": 0}
    return {
        "revenue": float(df[COL_REVENUE].sum()) if COL_REVENUE in df.columns else 0.0,
        "deals": int(len(df)),
    }


def _by_product(cur: pd.DataFrame, prev: pd.DataFrame, *, rd_labels) -> pd.DataFrame | None:
    """Группировка по бизнес-юниту или, если его нет — по сегментному тарифу."""
    key = COL_BUSINESS_UNIT if COL_BUSINESS_UNIT in cur.columns else COL_SEGMENT_TARIFF
    if key not in cur.columns or COL_REVENUE not in cur.columns:
        return None
    prev_lbl, cur_lbl = rd_labels
    cur_g = cur.groupby(key)[COL_REVENUE].sum().rename(cur_lbl)
    prev_g = prev.groupby(key)[COL_REVENUE].sum().rename(prev_lbl) if not prev.empty else pd.Series(name=prev_lbl, dtype=float)
    out = pd.concat([prev_g, cur_g], axis=1).fillna(0).reset_index()
    out = out.rename(columns={key: "Продукт"})
    out = out.sort_values(cur_lbl, ascending=False).head(10)
    return out


def _by_month(df: pd.DataFrame, quarter: int, year: int) -> pd.DataFrame | None:
    if df.empty or "_month" not in df.columns:
        return None
    months_in_q = [(quarter - 1) * 3 + i for i in (1, 2, 3)]
    g = df[df["_month"].isin(months_in_q)].groupby("_month")[COL_REVENUE].sum()
    if g.empty:
        return None
    out = pd.DataFrame({
        "Месяц": [f"{MONTHS_RU.get(int(m), str(m))} {year}" for m in g.index],
        "Выручка": g.values,
    })
    return out


def _top_tariffs(cur: pd.DataFrame, prev: pd.DataFrame, *, rd_labels) -> pd.DataFrame | None:
    if COL_TARIFF not in cur.columns:
        return None
    prev_lbl, cur_lbl = rd_labels
    cur_g = cur.groupby(COL_TARIFF)[COL_REVENUE].sum().rename(cur_lbl)
    prev_g = prev.groupby(COL_TARIFF)[COL_REVENUE].sum().rename(prev_lbl) if not prev.empty else pd.Series(name=prev_lbl, dtype=float)
    out = pd.concat([prev_g, cur_g], axis=1).fillna(0).reset_index()
    out = out.rename(columns={COL_TARIFF: "Тариф"})
    out["Δ%"] = out.apply(
        lambda r: f"{(r[cur_lbl]-r[prev_lbl])/r[prev_lbl]*100:+.0f}%" if r[prev_lbl] else "новый",
        axis=1,
    )
    out = out.sort_values(cur_lbl, ascending=False).head(10)
    return out


def _online_vs_offline(cur: pd.DataFrame, prev: pd.DataFrame, *, rd_labels) -> pd.DataFrame | None:
    key = COL_ONLINE
    if key not in cur.columns:
        return None
    prev_lbl, cur_lbl = rd_labels
    cur_g = cur.groupby(key)[COL_REVENUE].sum().rename(cur_lbl)
    prev_g = prev.groupby(key)[COL_REVENUE].sum().rename(prev_lbl) if not prev.empty else pd.Series(name=prev_lbl, dtype=float)
    out = pd.concat([prev_g, cur_g], axis=1).fillna(0).reset_index()
    out = out.rename(columns={key: "Канал"})
    return out


def _by_region(df: pd.DataFrame) -> pd.DataFrame | None:
    if COL_REGION not in df.columns:
        return None
    g = df.groupby(COL_REGION)[COL_REVENUE].sum().sort_values(ascending=False).head(10)
    if g.empty:
        return None
    return pd.DataFrame({"Регион": g.index, "Выручка": g.values})


def _by_sale_method(df: pd.DataFrame) -> pd.DataFrame | None:
    if COL_SALE_METHOD not in df.columns:
        return None
    g = df.groupby(COL_SALE_METHOD)[COL_REVENUE].sum().sort_values(ascending=False)
    if g.empty:
        return None
    return pd.DataFrame({"Способ продажи": g.index, "Выручка": g.values})
