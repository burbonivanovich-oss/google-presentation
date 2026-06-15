"""Прогноз выполнения годового плана 2026 по Маркету и ОФД.

Источник плана и факта — «Царь свод все разрезы.xlsx» (план-факт помесячно
по всем разрезам). Источник рекламного факта (целевые лиды) — data (15).xlsx.

Логика:
  1. Найти, в каких строках плана-факта «живут» Маркет и ОФД. План задаётся
     по «Сегмент плана» (напр. «ФН от плана ОФД на 36 мес.»), поэтому
     фильтруем сначала по продуктовому столбцу, если он есть, иначе по
     ключевым словам в сегменте/проекте/бизнес-юните.
  2. Собрать помесячный план и факт за 2025 и 2026.
  3. Спрогнозировать остаток 2026: базовый сценарий — оставшиеся месяцы
     выполняются с тем же % выполнения, что и YTD; сезонность берётся из
     самого плана (план уже её закладывает). Консервативный/оптимистичный —
     ±коридор от темпа.
  4. Посчитать прогнозный год-факт и % выполнения годового плана.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from .aggregator import (
    PF_MONTH, PF_QTY_FACT, PF_QTY_PLAN, PF_REV_FACT, PF_REV_PLAN,
    MONTHS_RU, _clean_plan_fact, _download_xlsx, _find, _list_xlsx,
)
from .drive import DriveClient

# Кандидаты на продуктовый столбец и ключевые слова направлений
PRODUCT_COLS = ["Продукт учета", "Продукт учёта", "Продукт", "Макропродукт",
                "Бизнес-юнит", "Сегмент плана групп", "Сегмент плана", "Проект",
                "Направление", "Сегмент Маркет"]

MARKET_KEYS = ["контур.маркет", "маркет"]
OFD_KEYS = ["контур.офд", "офд"]


@dataclass
class MonthCell:
    month: int
    plan_rev: float = 0.0
    fact_rev: float = 0.0
    plan_qty: float = 0.0
    fact_qty: float = 0.0


@dataclass
class ProductForecast:
    name: str
    filter_desc: str = ""
    months_2026: list[MonthCell] = field(default_factory=list)
    months_2025: list[MonthCell] = field(default_factory=list)
    completed_months: int = 0
    # годовые числа
    annual_plan_rev: float = 0.0
    ytd_plan_rev: float = 0.0
    ytd_fact_rev: float = 0.0
    annual_plan_qty: float = 0.0
    ytd_plan_qty: float = 0.0
    ytd_fact_qty: float = 0.0
    # прогноз выручки (3 сценария) — год целиком
    proj_rev_base: float = 0.0
    proj_rev_low: float = 0.0
    proj_rev_high: float = 0.0
    proj_qty_base: float = 0.0
    # производные
    ytd_fulfil: float = 0.0          # % выполнения YTD
    yoy_rev: float = 0.0             # YTD 2026 vs тот же период 2025


def run_forecast(drive: DriveClient, folder_id: str, year: int = 2026,
                 completed_through: int = 5) -> dict:
    """completed_through — последний завершённый месяц 2026 (по умолчанию май=5)."""
    files = _list_xlsx(drive, folder_id)
    svod = _find(files, [r"Царь свод", r"свод все разрезы"])
    if not svod:
        raise RuntimeError("Не найден Царь свод все разрезы.xlsx")
    raw = _download_xlsx(drive, svod["id"])
    df = _clean_plan_fact(raw)

    diagnostics = _diagnose(df)

    out = {"diagnostics": diagnostics, "products": {}}
    for name, keys in (("Маркет", MARKET_KEYS), ("ОФД", OFD_KEYS)):
        pf = _forecast_product(df, name, keys, year, completed_through)
        out["products"][name] = pf
    return out


def _diagnose(df: pd.DataFrame) -> dict:
    """Где в плане-факте есть Маркет/ОФД: по каждому кандидат-столбцу
    показываем уникальные значения с суммой плановой выручки."""
    info = {"columns": list(df.columns), "by_col": {}}
    for col in PRODUCT_COLS:
        if col not in df.columns:
            continue
        if PF_REV_PLAN not in df.columns:
            break
        g = df.groupby(col)[PF_REV_PLAN].sum().sort_values(ascending=False)
        rows = [(str(k), float(v)) for k, v in g.items()
                if any(t in str(k).lower() for t in ("маркет", "офд", "контур"))]
        if rows:
            info["by_col"][col] = rows
    return info


def _pick_filter(df: pd.DataFrame, keys: list[str]) -> tuple[pd.Series | None, str]:
    """Выбираем лучший столбец+значения, где совокупный плановый объём
    под ключевые слова максимален."""
    best_mask = None
    best_desc = ""
    best_total = -1.0
    for col in PRODUCT_COLS:
        if col not in df.columns or PF_REV_PLAN not in df.columns:
            continue
        col_vals = df[col].astype(str).str.lower()
        mask = col_vals.apply(lambda s: any(k in s for k in keys))
        if not mask.any():
            continue
        total = df.loc[mask, PF_REV_PLAN].sum()
        # предпочитаем продуктовые столбцы (точное "контур.маркет/офд")
        exact_bonus = 1e12 if col.lower().startswith("продукт") else 0
        score = float(total) + exact_bonus
        if score > best_total:
            best_total = score
            best_mask = mask
            matched = sorted(df.loc[mask, col].astype(str).unique())[:8]
            best_desc = f"{col} ∈ {{{', '.join(matched)}}}"
    return best_mask, best_desc


def _forecast_product(df: pd.DataFrame, name: str, keys: list[str],
                      year: int, completed: int) -> ProductForecast:
    pf = ProductForecast(name=name)
    mask, desc = _pick_filter(df, keys)
    pf.filter_desc = desc or "(не найдено)"
    if mask is None:
        return pf
    sub = df[mask]

    def month_series(y: int) -> list[MonthCell]:
        cells = []
        for m in range(1, 13):
            sel = sub[(sub["_year"] == y) & (sub["_month"] == m)]
            cells.append(MonthCell(
                month=m,
                plan_rev=float(sel[PF_REV_PLAN].sum()) if PF_REV_PLAN in sel else 0,
                fact_rev=float(sel[PF_REV_FACT].sum()) if PF_REV_FACT in sel else 0,
                plan_qty=float(sel[PF_QTY_PLAN].sum()) if PF_QTY_PLAN in sel else 0,
                fact_qty=float(sel[PF_QTY_FACT].sum()) if PF_QTY_FACT in sel else 0,
            ))
        return cells

    pf.months_2025 = month_series(year - 1)
    pf.months_2026 = month_series(year)
    pf.completed_months = completed

    pf.annual_plan_rev = sum(c.plan_rev for c in pf.months_2026)
    pf.annual_plan_qty = sum(c.plan_qty for c in pf.months_2026)
    pf.ytd_plan_rev = sum(c.plan_rev for c in pf.months_2026[:completed])
    pf.ytd_fact_rev = sum(c.fact_rev for c in pf.months_2026[:completed])
    pf.ytd_plan_qty = sum(c.plan_qty for c in pf.months_2026[:completed])
    pf.ytd_fact_qty = sum(c.fact_qty for c in pf.months_2026[:completed])

    pf.ytd_fulfil = (pf.ytd_fact_rev / pf.ytd_plan_rev * 100) if pf.ytd_plan_rev else 0
    prev_ytd = sum(c.fact_rev for c in pf.months_2025[:completed])
    pf.yoy_rev = ((pf.ytd_fact_rev - prev_ytd) / prev_ytd * 100) if prev_ytd else 0

    # Прогноз остатка: оставшиеся месяцы = план остатка × коэффициент темпа
    remain_plan_rev = sum(c.plan_rev for c in pf.months_2026[completed:])
    remain_plan_qty = sum(c.plan_qty for c in pf.months_2026[completed:])
    rate = pf.ytd_fact_rev / pf.ytd_plan_rev if pf.ytd_plan_rev else 1.0
    pf.proj_rev_base = pf.ytd_fact_rev + remain_plan_rev * rate
    pf.proj_rev_low = pf.ytd_fact_rev + remain_plan_rev * rate * 0.9
    pf.proj_rev_high = pf.ytd_fact_rev + remain_plan_rev * min(rate * 1.1, 1.0) \
        if rate < 1 else pf.ytd_fact_rev + remain_plan_rev
    qrate = pf.ytd_fact_qty / pf.ytd_plan_qty if pf.ytd_plan_qty else 1.0
    pf.proj_qty_base = pf.ytd_fact_qty + remain_plan_qty * qrate
    return pf


def print_report(result: dict, console) -> None:
    diag = result["diagnostics"]
    console.print("\n[bold]Столбцы Царь свода:[/bold]")
    console.print(", ".join(diag["columns"]))
    console.print("\n[bold]Где встречаются Маркет/ОФД/Контур (плановая выручка):[/bold]")
    for col, rows in diag["by_col"].items():
        console.print(f"\n  [cyan]{col}[/cyan]:")
        for k, v in rows[:15]:
            console.print(f"    {v:>16,.0f} ₽   {k}")

    for name, pf in result["products"].items():
        console.print(f"\n[bold green]══ {name} ══[/bold green]")
        console.print(f"Фильтр: {pf.filter_desc}")
        console.print(f"Годовой план 2026:   {pf.annual_plan_rev:>16,.0f} ₽ "
                      f"({pf.annual_plan_qty:,.0f} оплат)")
        console.print(f"YTD план ({pf.completed_months} мес): {pf.ytd_plan_rev:>14,.0f} ₽")
        console.print(f"YTD факт ({pf.completed_months} мес): {pf.ytd_fact_rev:>14,.0f} ₽ "
                      f"→ выполнение {pf.ytd_fulfil:.0f}% · YoY {pf.yoy_rev:+.0f}%")
        console.print(f"Прогноз года (база): {pf.proj_rev_base:>14,.0f} ₽ "
                      f"→ {pf.proj_rev_base / pf.annual_plan_rev * 100 if pf.annual_plan_rev else 0:.0f}% плана")
        console.print(f"  коридор: {pf.proj_rev_low:,.0f} … {pf.proj_rev_high:,.0f} ₽")
        console.print("  помесячно 2026 (план / факт), ₽:")
        for c in pf.months_2026:
            fact = f"{c.fact_rev:,.0f}" if c.month <= pf.completed_months else "—"
            console.print(f"    {MONTHS_RU[c.month]:<9} {c.plan_rev:>14,.0f} / {fact:>14}")
