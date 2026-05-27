from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml


@dataclass
class SheetSource:
    spreadsheet_id: str
    range: str          # например "channels!A1:H100"


@dataclass
class InsightRules:
    qoq_threshold_pct: float = 20.0
    sigma: float = 3.0
    roas_benchmark: float = 3.0


@dataclass
class ReportConfig:
    name: str
    presentation_template_id: str
    # Идентификатор слайда в шаблоне, который дублируется под каждый инсайт.
    # В шаблоне в этом слайде должны быть плейсхолдеры {{insight_headline}} и {{insight_detail}}.
    insight_slide_id: str
    sources: dict[str, SheetSource]
    insights: InsightRules
    static_placeholders: dict[str, str]

    @classmethod
    def load(cls, path: str | Path) -> "ReportConfig":
        data = yaml.safe_load(Path(path).read_text())
        sources = {
            key: SheetSource(**val) for key, val in data.get("sources", {}).items()
        }
        return cls(
            name=data["name"],
            presentation_template_id=data["presentation_template_id"],
            insight_slide_id=data["insight_slide_id"],
            sources=sources,
            insights=InsightRules(**data.get("insights", {})),
            static_placeholders=data.get("static_placeholders", {}),
        )
