"""Автогенератор текстовых тезисов и выводов для каждой секции отчёта.

Получает aggregated DataFrame, выдаёт короткую формулировку (для
заголовка-тезиса) и список из 3-5 буллитов (для блока «Выводы»).
"""
from __future__ import annotations

import pandas as pd

from .data_adapter import ReportData


def fmt_money(v: float) -> str:
    if v != v:
        return "—"
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.2f} млн ₽"
    if abs(v) >= 1_000:
        return f"{v/1_000:.0f} тыс ₽"
    return f"{v:.0f} ₽"


def pct_change(cur: float, prev: float) -> str:
    if not prev:
        return "новое"
    delta = (cur - prev) / prev * 100
    return f"{delta:+.1f}%"


def overall_dynamics(rd: ReportData) -> tuple[str, list[str]]:
    cur = rd.totals.get("revenue", 0)
    prev = rd.totals_prev.get("revenue", 0)
    deals = rd.totals.get("deals", 0)
    delta = pct_change(cur, prev)
    headline = (
        f"Общая динамика: выручка {fmt_money(cur)} в {rd.period_label}"
    )
    if prev:
        headline += f", {delta} к {rd.prev_period_label}"
    bullets = [
        f"Суммарная выручка за {rd.period_label}: {fmt_money(cur)}",
        f"Количество оплат: {deals:,}".replace(",", " "),
    ]
    if prev:
        bullets.append(
            f"К {rd.prev_period_label}: {delta} (было {fmt_money(prev)})"
        )
    return headline, bullets


def by_product_dynamics(rd: ReportData) -> tuple[str, list[str]] | None:
    df = rd.by_product
    if df is None or df.empty:
        return None
    cur_col = rd.period_label
    prev_col = rd.prev_period_label
    if cur_col not in df.columns:
        return None
    top = df.iloc[0]
    headline = (
        f"По продуктам: лидер — {top['Продукт']} ({fmt_money(top[cur_col])})"
    )
    bullets = []
    for _, row in df.head(5).iterrows():
        prev_v = row.get(prev_col, 0)
        cur_v = row[cur_col]
        delta = pct_change(cur_v, prev_v) if prev_col in df.columns else ""
        line = f"{row['Продукт']}: {fmt_money(cur_v)}"
        if delta:
            line += f" ({delta})"
        bullets.append(line)
    return headline, bullets


def monthly_dynamics(rd: ReportData) -> tuple[str, list[str]] | None:
    df = rd.by_month
    if df is None or df.empty:
        return None
    total = df["Выручка"].sum()
    biggest = df.loc[df["Выручка"].idxmax()]
    smallest = df.loc[df["Выручка"].idxmin()]
    headline = f"Помесячная динамика: пик в {biggest['Месяц']} ({fmt_money(biggest['Выручка'])})"
    bullets = []
    for _, row in df.iterrows():
        share = (row["Выручка"] / total * 100) if total else 0
        bullets.append(f"{row['Месяц']}: {fmt_money(row['Выручка'])} ({share:.0f}% от квартала)")
    if biggest['Месяц'] != smallest['Месяц']:
        delta = (biggest["Выручка"] - smallest["Выручка"]) / smallest["Выручка"] * 100 if smallest["Выручка"] else 0
        bullets.append(f"Разрыв пик/минимум: {delta:+.0f}%")
    return headline, bullets


def top_tariffs(rd: ReportData) -> tuple[str, list[str]] | None:
    df = rd.top_tariffs
    if df is None or df.empty:
        return None
    cur_col = rd.period_label
    headline = f"Топ-10 тарифов по выручке за {rd.period_label}"
    bullets = []
    for _, row in df.head(5).iterrows():
        bullets.append(
            f"{row['Тариф']}: {fmt_money(row[cur_col])} ({row.get('Δ%', '')})"
        )
    return headline, bullets


def online_vs_offline(rd: ReportData) -> tuple[str, list[str]] | None:
    df = rd.online_vs_offline
    if df is None or df.empty:
        return None
    cur_col = rd.period_label
    if cur_col not in df.columns:
        return None
    total = df[cur_col].sum()
    headline = "Соотношение онлайн и офлайн каналов"
    bullets = []
    for _, row in df.iterrows():
        share = (row[cur_col] / total * 100) if total else 0
        bullets.append(f"{row['Канал']}: {fmt_money(row[cur_col])} ({share:.0f}%)")
    return headline, bullets


def by_region(rd: ReportData) -> tuple[str, list[str]] | None:
    df = rd.by_region
    if df is None or df.empty:
        return None
    total = df["Выручка"].sum()
    headline = f"Топ-10 регионов: лидер — {df.iloc[0]['Регион']}"
    bullets = []
    for _, row in df.head(5).iterrows():
        share = (row["Выручка"] / total * 100) if total else 0
        bullets.append(f"{row['Регион']}: {fmt_money(row['Выручка'])} ({share:.0f}%)")
    return headline, bullets


def by_sale_method(rd: ReportData) -> tuple[str, list[str]] | None:
    df = rd.by_sale_method
    if df is None or df.empty:
        return None
    total = df["Выручка"].sum()
    headline = "Распределение по способам продажи"
    bullets = []
    for _, row in df.head(5).iterrows():
        share = (row["Выручка"] / total * 100) if total else 0
        bullets.append(f"{row['Способ продажи']}: {fmt_money(row['Выручка'])} ({share:.0f}%)")
    return headline, bullets
