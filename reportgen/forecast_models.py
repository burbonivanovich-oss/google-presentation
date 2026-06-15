"""Модели прогноза временного ряда для помесячных продаж.

Данные ограничены (~17 полных месяцев: янв 2025 – май 2026), поэтому
используем АНСАМБЛЬ из нескольких устойчивых методов и честно показываем
разброс между ними как сценарный коридор.

Методы:
  1. plan_pacing   — оставшиеся месяцы = план месяца × (факт YTD / план YTD).
                     Самый бизнес-обоснованный: план уже закладывает сезонность,
                     а текущий темп выполнения переносится на остаток года.
  2. seasonal_yoy  — оставшийся месяц = факт того же месяца 2025 × (1 + g),
                     где g — прирост YTD 2026 к тому же периоду 2025.
  3. ets_holt      — экспоненциальное сглаживание с трендом (statsmodels),
                     сезонное (period=12), если хватает данных, иначе демпф. тренд.
  4. linreg_season — линейная регрессия по индексу времени + сезонные дамми
                     (scikit-learn).

Ансамбль (база) — взвешенное среднее доступных методов.
Коридор: низ = min по методам, верх = max по методам (помесячно).
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

MONTHS = list(range(1, 13))


@dataclass
class SeriesForecast:
    completed: int
    annual_plan: float
    ytd_plan: float
    ytd_fact: float
    prev_ytd_fact: float          # факт тех же месяцев 2025
    fulfil_ytd: float             # % выполнения YTD
    yoy: float                    # прирост YTD к 2025
    # помесячный прогноз остатка (месяц → значение)
    base: dict = field(default_factory=dict)
    low: dict = field(default_factory=dict)
    high: dict = field(default_factory=dict)
    by_method: dict = field(default_factory=dict)   # метод → {месяц: значение}
    # годовые итоги
    year_base: float = 0.0
    year_low: float = 0.0
    year_high: float = 0.0
    year_by_method: dict = field(default_factory=dict)
    pct_plan_base: float = 0.0
    pct_plan_low: float = 0.0
    pct_plan_high: float = 0.0
    notes: list = field(default_factory=list)


def _clip(x: float) -> float:
    return max(0.0, float(x))


def forecast_series(fact_2025: dict[int, float], fact_2026: dict[int, float],
                    plan_2026: dict[int, float], completed: int) -> SeriesForecast:
    annual_plan = sum(plan_2026.get(m, 0.0) for m in MONTHS)
    ytd_plan = sum(plan_2026.get(m, 0.0) for m in MONTHS[:completed])
    ytd_fact = sum(fact_2026.get(m, 0.0) for m in MONTHS[:completed])
    prev_ytd = sum(fact_2025.get(m, 0.0) for m in MONTHS[:completed])
    rem = MONTHS[completed:]

    fc = SeriesForecast(
        completed=completed, annual_plan=annual_plan, ytd_plan=ytd_plan,
        ytd_fact=ytd_fact, prev_ytd_fact=prev_ytd,
        fulfil_ytd=(ytd_fact / ytd_plan * 100) if ytd_plan else 0.0,
        yoy=((ytd_fact - prev_ytd) / prev_ytd * 100) if prev_ytd else 0.0,
    )

    methods: dict[str, dict[int, float]] = {}

    # 1) plan pacing
    rate = (ytd_fact / ytd_plan) if ytd_plan else 1.0
    methods["plan_pacing"] = {m: _clip(plan_2026.get(m, 0.0) * rate) for m in rem}

    # 2) seasonal YoY
    g = (ytd_fact / prev_ytd - 1.0) if prev_ytd else 0.0
    methods["seasonal_yoy"] = {m: _clip(fact_2025.get(m, 0.0) * (1 + g)) for m in rem}

    # сплошной ряд фактов для моделей
    series = []
    for y in (2025, 2026):
        src = fact_2025 if y == 2025 else fact_2026
        last = 12 if y == 2025 else completed
        for m in range(1, last + 1):
            series.append(src.get(m, 0.0))
    series = np.array(series, dtype=float)
    horizon = len(rem)

    # 3) ETS / Holt (statsmodels)
    try:
        methods["ets_holt"] = _ets(series, horizon, rem)
    except Exception as e:  # noqa: BLE001
        fc.notes.append(f"ETS недоступен ({type(e).__name__}), пропущен")

    # 4) линейная регрессия с сезонными дамми
    try:
        methods["linreg_season"] = _linreg(series, horizon, rem)
    except Exception as e:  # noqa: BLE001
        fc.notes.append(f"Регрессия недоступна ({type(e).__name__}), пропущена")

    fc.by_method = methods

    # веса ансамбля (нормируются по доступным методам)
    weights = {"plan_pacing": 0.40, "seasonal_yoy": 0.25,
               "ets_holt": 0.15, "linreg_season": 0.20}
    avail = {k: weights[k] for k in methods}
    wsum = sum(avail.values()) or 1.0
    for m in rem:
        vals = [methods[k][m] for k in methods]
        fc.base[m] = sum(methods[k][m] * avail[k] for k in methods) / wsum
        fc.low[m] = min(vals)
        fc.high[m] = max(vals)

    fc.year_base = ytd_fact + sum(fc.base.values())
    fc.year_low = ytd_fact + sum(fc.low.values())
    fc.year_high = ytd_fact + sum(fc.high.values())
    fc.year_by_method = {k: ytd_fact + sum(v.values()) for k, v in methods.items()}
    if annual_plan:
        fc.pct_plan_base = fc.year_base / annual_plan * 100
        fc.pct_plan_low = fc.year_low / annual_plan * 100
        fc.pct_plan_high = fc.year_high / annual_plan * 100
    return fc


def _ets(series: np.ndarray, horizon: int, rem: list[int]) -> dict[int, float]:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    if len(series) >= 24:
        model = ExponentialSmoothing(series, trend="add", damped_trend=True,
                                     seasonal="add", seasonal_periods=12)
    else:
        model = ExponentialSmoothing(series, trend="add", damped_trend=True)
    fit = model.fit(optimized=True)
    pred = fit.forecast(horizon)
    return {m: _clip(pred[i]) for i, m in enumerate(rem)}


def _linreg(series: np.ndarray, horizon: int, rem: list[int]) -> dict[int, float]:
    from sklearn.linear_model import LinearRegression
    n = len(series)
    t = np.arange(n)
    # сезонные дамми по месяцу (индекс начинается с января 2025)
    month_idx = [(i % 12) for i in range(n)]
    X = np.column_stack([t] + [(np.array(month_idx) == k).astype(float)
                               for k in range(1, 12)])
    reg = LinearRegression().fit(X, series)
    out = {}
    for i, m in enumerate(rem):
        gi = n + i
        mi = gi % 12
        row = [gi] + [(1.0 if (mi == k) else 0.0) for k in range(1, 12)]
        out[m] = _clip(reg.predict(np.array([row]))[0])
    return out
