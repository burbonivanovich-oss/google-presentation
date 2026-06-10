"""Planner v2: канва квартального отчёта по образцу Петровой/Порубова.

Базовый кирпич — секция «slide_chart» = тезис-заголовок + график (LINKED
chart из Sheets) + список выводов. Эта секция повторяется для каждого
аналитического разреза: общая динамика, по продуктам, помесячно,
топ тарифов, онлайн vs офлайн, по регионам, по способу продажи.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from . import conclusions
from .data_adapter import ReportData


@dataclass
class Step:
    role: str  # cover | slide_chart | slide_text | final
    data: dict[str, Any] = field(default_factory=dict)


def build_plan(
    *,
    report_name: str,
    period: str,
    previous_period: str,
    rd: ReportData,
    author_name: str = "",
    author_role: str = "",
) -> list[Step]:
    plan: list[Step] = []

    plan.append(Step("cover", {
        "title": f"{report_name}\n{rd.period_label}",
        "name": author_name or "Команда аналитики",
        "subtitle": author_role or f"Квартальный отчёт {rd.period_label}",
        "url": "kontur.ru",
    }))

    # 1. Общая динамика
    headline, bullets = conclusions.overall_dynamics(rd)
    plan.append(Step("slide_chart", {
        "title": headline,
        "bullets": bullets,
        "chart_kind": "column",
        "chart_title": f"Выручка по продуктам — {rd.period_label} vs {rd.prev_period_label}",
        "dataframe": rd.by_product,
        "categories_col": "Продукт",
        "series_cols": [rd.prev_period_label, rd.period_label],
    }))

    # 2. Помесячная динамика
    res = conclusions.monthly_dynamics(rd)
    if res and rd.by_month is not None:
        headline, bullets = res
        plan.append(Step("slide_chart", {
            "title": headline,
            "bullets": bullets,
            "chart_kind": "column",
            "chart_title": f"Выручка по месяцам {rd.period_label}",
            "dataframe": rd.by_month,
            "categories_col": "Месяц",
            "series_cols": ["Выручка"],
        }))

    # 3. По продуктам — отдельный детальный слайд
    res = conclusions.by_product_dynamics(rd)
    if res and rd.by_product is not None:
        headline, bullets = res
        plan.append(Step("slide_chart", {
            "title": headline,
            "bullets": bullets,
            "chart_kind": "bar",  # горизонтальные полосы для рейтинга
            "chart_title": f"Выручка по продуктам — {rd.period_label}",
            "dataframe": rd.by_product,
            "categories_col": "Продукт",
            "series_cols": [rd.period_label],
        }))

    # 4. Топ тарифов
    res = conclusions.top_tariffs(rd)
    if res and rd.top_tariffs is not None:
        headline, bullets = res
        plan.append(Step("slide_chart", {
            "title": headline,
            "bullets": bullets,
            "chart_kind": "bar",
            "chart_title": f"Топ-10 тарифов — {rd.period_label}",
            "dataframe": rd.top_tariffs,
            "categories_col": "Тариф",
            "series_cols": [rd.period_label],
        }))

    # 5. Онлайн vs офлайн
    res = conclusions.online_vs_offline(rd)
    if res and rd.online_vs_offline is not None:
        headline, bullets = res
        plan.append(Step("slide_chart", {
            "title": headline,
            "bullets": bullets,
            "chart_kind": "pie",
            "chart_title": f"Онлайн vs Офлайн — {rd.period_label}",
            "dataframe": rd.online_vs_offline,
            "categories_col": "Канал",
            "series_cols": [rd.period_label],
        }))

    # 6. По регионам
    res = conclusions.by_region(rd)
    if res and rd.by_region is not None:
        headline, bullets = res
        plan.append(Step("slide_chart", {
            "title": headline,
            "bullets": bullets,
            "chart_kind": "bar",
            "chart_title": f"Топ-10 регионов — {rd.period_label}",
            "dataframe": rd.by_region,
            "categories_col": "Регион",
            "series_cols": ["Выручка"],
        }))

    # 7. По способу продажи
    res = conclusions.by_sale_method(rd)
    if res and rd.by_sale_method is not None:
        headline, bullets = res
        plan.append(Step("slide_chart", {
            "title": headline,
            "bullets": bullets,
            "chart_kind": "pie",
            "chart_title": f"Способы продажи — {rd.period_label}",
            "dataframe": rd.by_sale_method,
            "categories_col": "Способ продажи",
            "series_cols": ["Выручка"],
        }))

    # 8. Итоговый «Задачи на следующий квартал» — пустой текст-слайд
    plan.append(Step("slide_text", {
        "title": "Что делаем дальше",
        "bullets": [
            "(заполняется руками после прогона) — приоритеты команды на следующий квартал",
            "Цели по выручке по продуктам",
            "Эксперименты с каналами привлечения",
            "Тарифные изменения, маркетинговые активности",
        ],
    }))

    plan.append(Step("final", {
        "title": "Спасибо!\nВопросы?",
        "name": author_name or "Команда аналитики",
        "subtitle": author_role or f"Отчёт за {rd.period_label}",
    }))
    return plan
