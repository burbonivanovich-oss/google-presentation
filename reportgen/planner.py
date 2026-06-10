"""Planner: 10 автоматических слайдов в стилистике Петровой.

Список слайдов фиксирован, под каждый — конкретный источник данных.
Если данных не хватает — слайд пропускается, чтобы не было пустых.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from .data_adapter import MultiSourceData


@dataclass
class Step:
    role: str  # cover | slide_chart | slide_text | slide_table | final
    data: dict[str, Any] = field(default_factory=dict)


def build_plan(*, report_name: str, rd: MultiSourceData,
               author_name: str = "", author_role: str = "") -> list[Step]:
    plan: list[Step] = []

    # 1. Обложка
    plan.append(Step("cover", {
        "title": f"{report_name}\n{rd.period_label}",
        "name": author_name or "Команда аналитики",
        "subtitle": author_role or f"Квартальный отчёт {rd.period_label}",
        "url": "kontur.ru",
    }))

    # 2. План-Факт всего за квартал
    if rd.pf_quarter:
        cur = rd.pf_quarter["cur"]
        plan.append(Step("slide_text", {
            "title": f"План-Факт БЮ — {rd.period_label}",
            "bullets": [
                f"Выручка: факт {_money(cur['fact_rev'])} / план {_money(cur['plan_rev'])}",
                f"% выполнения плана: {rd.pf_quarter['cur_pct']:.0f}%",
                f"К {rd.prev_period_label}: {rd.pf_quarter['qoq_delta']:+.1f}%",
                f"YoY (к Q{rd.cur_q} {rd.cur_y - 1}): {rd.pf_quarter['yoy_delta']:+.1f}%",
                f"Кол-во оплат: {int(cur['fact_qty']):,}".replace(",", " "),
            ],
        }))

    # 3. План-Факт помесячно — column chart
    if rd.pf_monthly is not None and not rd.pf_monthly.empty:
        plan.append(Step("slide_chart", {
            "title": f"План-Факт по месяцам — {rd.period_label}",
            "bullets": [
                f"{row['Месяц']}: факт {_money(row['Факт, ₽'])} / план {_money(row['План, ₽'])}"
                for _, row in rd.pf_monthly.iterrows()
            ],
            "chart_kind": "column",
            "chart_title": "План vs Факт по месяцам",
            "dataframe": rd.pf_monthly,
            "categories_col": "Месяц",
            "series_cols": ["План, ₽", "Факт, ₽"],
        }))

    # 4. Выручка по продуктам (бизнес-юнитам)
    if rd.by_product is not None and not rd.by_product.empty:
        plan.append(Step("slide_chart", {
            "title": f"Выручка по продуктам — {rd.period_label}",
            "bullets": _diff_bullets(rd.by_product, rd, key=_first_col(rd.by_product)),
            "chart_kind": "column",
            "chart_title": "Выручка по бизнес-юнитам",
            "dataframe": rd.by_product,
            "categories_col": _first_col(rd.by_product),
            "series_cols": _other_cols(rd.by_product),
        }))

    # 5. Помесячная выручка по транзакциям
    if rd.by_month is not None and not rd.by_month.empty:
        biggest = rd.by_month.loc[rd.by_month["Выручка"].idxmax()]
        plan.append(Step("slide_chart", {
            "title": f"Помесячная динамика — {rd.period_label}",
            "bullets": [
                f"Пик: {biggest['Месяц']} — {_money(biggest['Выручка'])}",
                *(f"{row['Месяц']}: {_money(row['Выручка'])}" for _, row in rd.by_month.iterrows()),
            ],
            "chart_kind": "column",
            "chart_title": "Выручка по месяцам",
            "dataframe": rd.by_month,
            "categories_col": "Месяц",
            "series_cols": ["Выручка"],
        }))

    # 6. Топ-10 тарифов
    if rd.top_tariffs is not None and not rd.top_tariffs.empty:
        plan.append(Step("slide_chart", {
            "title": f"Топ-10 тарифов — {rd.period_label}",
            "bullets": _diff_bullets(rd.top_tariffs, rd, key=_first_col(rd.top_tariffs), top=5),
            "chart_kind": "bar",
            "chart_title": "Топ-10 тарифов по выручке",
            "dataframe": rd.top_tariffs,
            "categories_col": _first_col(rd.top_tariffs),
            "series_cols": [rd.period_label],
        }))

    # 7. Онлайн vs офлайн
    if rd.online_vs_offline is not None and not rd.online_vs_offline.empty:
        plan.append(Step("slide_chart", {
            "title": "Соотношение онлайн / офлайн",
            "bullets": _share_bullets(rd.online_vs_offline, rd.period_label),
            "chart_kind": "pie",
            "chart_title": "Доли каналов продаж",
            "dataframe": rd.online_vs_offline,
            "categories_col": _first_col(rd.online_vs_offline),
            "series_cols": [rd.period_label],
        }))

    # 8. Топ регионов
    if rd.by_region is not None and not rd.by_region.empty:
        plan.append(Step("slide_chart", {
            "title": "Топ-10 регионов по выручке",
            "bullets": _share_bullets(rd.by_region, rd.period_label, top=5),
            "chart_kind": "bar",
            "chart_title": "Регионы — топ-10",
            "dataframe": rd.by_region,
            "categories_col": _first_col(rd.by_region),
            "series_cols": [rd.period_label],
        }))

    # 9. CAC / CPL / Конверсия (если есть)
    if rd.cac_cpl is not None and not rd.cac_cpl.empty:
        plan.append(Step("slide_table", {
            "title": "Стоимость клиента и конверсия",
            "dataframe": rd.cac_cpl,
        }))

    # 10. Финал
    plan.append(Step("final", {
        "title": "Спасибо!\nВопросы?",
        "name": author_name or "Команда аналитики",
        "subtitle": author_role or f"Отчёт за {rd.period_label}",
    }))
    return plan


# ── вспомогательное ──────────────────────────────────────────────────

def _money(v: float) -> str:
    if v != v or v is None:
        return "—"
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.2f} млн ₽"
    if abs(v) >= 1_000:
        return f"{v/1_000:.0f} тыс ₽"
    return f"{v:.0f} ₽"


def _first_col(df) -> str:
    return df.columns[0]


def _other_cols(df) -> list[str]:
    return [c for c in df.columns[1:]]


def _diff_bullets(df, rd: MultiSourceData, *, key: str, top: int = 5) -> list[str]:
    out = []
    cur = rd.period_label
    prev = rd.prev_period_label
    for _, row in df.head(top).iterrows():
        if prev in df.columns and row.get(prev, 0):
            delta = (row[cur] - row[prev]) / row[prev] * 100
            out.append(f"{row[key]}: {_money(row[cur])} ({delta:+.1f}%)")
        else:
            out.append(f"{row[key]}: {_money(row[cur])}")
    return out


def _share_bullets(df, value_col: str, top: int = 5) -> list[str]:
    total = df[value_col].sum() if value_col in df.columns else df.iloc[:, 1].sum()
    col = value_col if value_col in df.columns else df.columns[1]
    key = df.columns[0]
    out = []
    for _, row in df.head(top).iterrows():
        share = (row[col] / total * 100) if total else 0
        out.append(f"{row[key]}: {_money(row[col])} ({share:.0f}%)")
    return out
