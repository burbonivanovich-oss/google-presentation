from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass
class Insight:
    """Один найденный факт, который заслуживает отдельного слайда."""

    kind: str            # "qoq_jump" / "anomaly" / "roas_below" / ...
    entity: str          # канал, метрика и т.п.
    headline: str        # краткая формулировка для заголовка слайда
    detail: str          # текст для подписи
    value: float | None = None


def qoq_changes(
    df: pd.DataFrame,
    *,
    entity_col: str,
    metric_col: str,
    current_period: str,
    previous_period: str,
    period_col: str = "period",
    threshold_pct: float = 20.0,
) -> list[Insight]:
    """Изменения метрики между двумя кварталами выше порога в %."""
    cur = df[df[period_col] == current_period].set_index(entity_col)[metric_col]
    prev = df[df[period_col] == previous_period].set_index(entity_col)[metric_col]
    joined = pd.concat({"cur": cur, "prev": prev}, axis=1).dropna()
    joined = joined[joined["prev"] != 0]
    joined["delta_pct"] = (joined["cur"] - joined["prev"]) / joined["prev"] * 100

    out: list[Insight] = []
    for entity, row in joined[joined["delta_pct"].abs() >= threshold_pct].iterrows():
        direction = "вырос" if row["delta_pct"] > 0 else "упал"
        out.append(
            Insight(
                kind="qoq_jump",
                entity=str(entity),
                headline=f"{entity}: {metric_col} {direction} на {abs(row['delta_pct']):.0f}%",
                detail=(
                    f"Прошлый период: {row['prev']:.2f}, текущий: {row['cur']:.2f}. "
                    f"Изменение Q/Q: {row['delta_pct']:+.1f}%."
                ),
                value=float(row["delta_pct"]),
            )
        )
    return out


def sigma_anomalies(
    df: pd.DataFrame,
    *,
    entity_col: str,
    metric_col: str,
    sigma: float = 3.0,
) -> list[Insight]:
    """Выбросы по правилу N-сигма по последнему периоду."""
    values = pd.to_numeric(df[metric_col], errors="coerce").dropna()
    if len(values) < 3:
        return []
    mean, std = values.mean(), values.std()
    if std == 0:
        return []
    out: list[Insight] = []
    for _, row in df.iterrows():
        v = pd.to_numeric(row[metric_col], errors="coerce")
        if pd.isna(v):
            continue
        z = (v - mean) / std
        if abs(z) >= sigma:
            out.append(
                Insight(
                    kind="anomaly",
                    entity=str(row[entity_col]),
                    headline=f"Аномалия: {row[entity_col]} по {metric_col} (z={z:+.1f})",
                    detail=(
                        f"Значение {v:.2f} отклоняется от среднего {mean:.2f} "
                        f"на {z:+.1f}σ (std={std:.2f})."
                    ),
                    value=float(z),
                )
            )
    return out


def roas_below_benchmark(
    df: pd.DataFrame,
    *,
    entity_col: str,
    roas_col: str = "roas",
    benchmark: float = 3.0,
) -> list[Insight]:
    out: list[Insight] = []
    for _, row in df.iterrows():
        v = pd.to_numeric(row[roas_col], errors="coerce")
        if pd.isna(v) or v >= benchmark:
            continue
        out.append(
            Insight(
                kind="roas_below",
                entity=str(row[entity_col]),
                headline=f"{row[entity_col]}: ROAS {v:.2f} ниже бенчмарка {benchmark}",
                detail=f"Канал не окупается по целевому ROAS={benchmark}. Текущий: {v:.2f}.",
                value=float(v),
            )
        )
    return out
