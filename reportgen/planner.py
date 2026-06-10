"""Planner: на основе данных и конфига собирает «программу» слайдов.

Каждый Step — это инструкция для композитора: «возьми эталон роли X и
заполни его этими значениями». Композитор разворачивает их по очереди.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd


@dataclass
class Step:
    role: str  # cover | section | kpi_6 | big_quote | facts_3 | table | chart_column | final | ...
    data: dict[str, Any] = field(default_factory=dict)


def build_plan(
    *,
    report_name: str,
    period: str,
    previous_period: str,
    data: dict[str, pd.DataFrame],
    author_name: str = "",
    author_role: str = "",
) -> list[Step]:
    """Базовая логика выбора слайдов под квартальный отчёт.

    Сценарий: cover → section('Итоги квартала') → KPI → section('По
    каналам') → table или chart → section('Инсайты') → big_quote +
    facts → final.
    """
    plan: list[Step] = []
    plan.append(Step("cover", {
        "title": f"{report_name}\n{period}",
        "name": author_name or "Команда аналитики",
        "subtitle": author_role or f"Отчёт за {period}",
        "url": "kontur.ru",
    }))

    channels = _numeric(data.get("channels"))

    # Раздел 1: KPI
    kpis = _aggregate_kpis(channels, period, previous_period)
    if kpis:
        plan.append(Step("section", {
            "title": "Итоги квартала",
            "body": f"Сводные показатели за {period} в сравнении с {previous_period}.",
        }))
        if len(kpis) <= 2:
            payload = {"title": "Ключевые показатели"}
            for i, (label, value, _delta) in enumerate(kpis[:2], 1):
                payload[f"value_{i}"] = value
                payload[f"desc_{i}"] = label
            plan.append(Step("kpi_2", payload))
        else:
            payload = {}
            for i, (label, value, _delta) in enumerate(kpis[:6], 1):
                payload[f"value_{i}"] = value
                payload[f"desc_{i}"] = label
            plan.append(Step("kpi_6", payload))

    # Раздел 2: каналы
    if channels is not None and not channels.empty:
        plan.append(Step("section", {
            "title": "Динамика по каналам",
            "body": f"Сравнение {previous_period} → {period} в разрезе каналов.",
        }))
        # Таблица по каналам (если есть метрики)
        cmp = _channel_compare(channels, period, previous_period)
        if cmp is not None and not cmp.empty:
            plan.append(Step("table", {
                "title": f"Каналы — {previous_period} vs {period}",
                "dataframe": cmp,
            }))
            # И график расходов
            plan.append(Step("chart_column", {
                "title": f"Расходы по каналам — {previous_period} vs {period}",
                "categories_col": "channel",
                "series_cols": [previous_period, period],
                "dataframe": cmp.copy(),
            }))

    # Раздел 3: инсайты
    insights = data.get("_insights") or []
    if insights:
        plan.append(Step("section", {
            "title": "Ключевые инсайты",
            "body": f"Что мы видим в данных за {period}.",
        }))
        top = insights[0]
        plan.append(Step("big_quote", {
            "preface": top.headline,
            "value": _short_metric(top),
            "unit": top.unit if hasattr(top, "unit") else "",
        }))
        rest = insights[1:4]
        if rest:
            payload = {"title": "Дополнительные находки"}
            for i, ins in enumerate(rest, 1):
                payload[f"num_{i}"] = str(i)
                payload[f"desc_{i}"] = f"{ins.headline}\n{ins.detail}"
            plan.append(Step("facts_3", payload))

    plan.append(Step("final", {
        "title": "Спасибо!\nВопросы?",
        "name": author_name or "Команда аналитики",
        "subtitle": author_role or f"Отчёт за {period}",
    }))
    return plan


def _numeric(df: pd.DataFrame | None) -> pd.DataFrame | None:
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        if col in ("channel", "period"):
            continue
        out[col] = out[col].apply(_to_num)
    return out


def _to_num(x):
    try:
        return float(str(x).replace(" ", "").replace(",", ".").replace("\xa0", ""))
    except (ValueError, TypeError):
        return float("nan")


def _aggregate_kpis(df: pd.DataFrame | None, cur: str, prev: str) -> list[tuple[str, str, float]]:
    if df is None or df.empty or "period" not in df.columns:
        return []
    cur_df = df[df["period"] == cur]
    prev_df = df[df["period"] == prev]
    out = []
    for col in ("revenue", "spend", "leads", "roas"):
        if col not in df.columns:
            continue
        cur_total = cur_df[col].sum()
        prev_total = prev_df[col].sum()
        if cur_total != cur_total:  # NaN
            continue
        delta = ((cur_total - prev_total) / prev_total * 100) if prev_total else 0
        sign = "+" if delta >= 0 else ""
        label = {
            "revenue": f"Выручка ({sign}{delta:.1f}%)",
            "spend": f"Расходы ({sign}{delta:.1f}%)",
            "leads": f"Лиды ({sign}{delta:.1f}%)",
            "roas": f"ROAS среднее ({sign}{delta:.1f}%)",
        }[col]
        out.append((label, _fmt(cur_total), delta))
    return out


def _channel_compare(df: pd.DataFrame, cur: str, prev: str) -> pd.DataFrame | None:
    if "channel" not in df.columns or "period" not in df.columns:
        return None
    metric = next((m for m in ("spend", "revenue", "leads") if m in df.columns), None)
    if not metric:
        return None
    cur_df = df[df["period"] == cur][["channel", metric]].rename(columns={metric: cur})
    prev_df = df[df["period"] == prev][["channel", metric]].rename(columns={metric: prev})
    merged = prev_df.merge(cur_df, on="channel", how="outer").fillna(0)
    merged["Δ %"] = merged.apply(
        lambda r: f"{(r[cur]-r[prev])/r[prev]*100:+.1f}%" if r[prev] else "—", axis=1
    )
    return merged


def _short_metric(insight) -> str:
    # пытаемся вынуть число из headline / details — иначе ставим вопрос
    import re
    for src in (getattr(insight, "headline", ""), getattr(insight, "detail", "")):
        m = re.search(r"[-+]?\d+[.,]?\d*\s*%?", src)
        if m:
            return m.group(0)
    return "—"


def _fmt(x: float) -> str:
    if x != x:
        return "—"
    if abs(x) >= 1_000_000:
        return f"{x/1_000_000:.1f}M"
    if abs(x) >= 1_000:
        return f"{x/1_000:.0f}K"
    return f"{x:.1f}"
