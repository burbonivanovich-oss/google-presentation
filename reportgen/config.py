from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml


@dataclass
class SheetSource:
    range: str                          # "A1:F1000" или "channels!A1:F1000"
    spreadsheet_id: str | None = None   # явный ID
    name_pattern: str | None = None     # либо паттерн имени в folder_id

    def __post_init__(self) -> None:
        if not self.spreadsheet_id and not self.name_pattern:
            raise ValueError(
                "У источника должен быть либо spreadsheet_id, либо name_pattern"
            )


@dataclass
class InsightRules:
    qoq_threshold_pct: float = 20.0
    sigma: float = 3.0
    roas_benchmark: float = 3.0


@dataclass
class ReportConfig:
    name: str
    insight_slide_id: str
    sources: dict[str, SheetSource]
    insights: InsightRules
    static_placeholders: dict[str, str]
    # Папка-приёмник в Drive: все таблицы и шаблон лежат тут.
    folder_id: str | None = None
    presentation_template_id: str | None = None
    presentation_template_name: str | None = None  # ищется в folder_id

    def __post_init__(self) -> None:
        if not self.presentation_template_id and not self.presentation_template_name:
            raise ValueError("Нужен presentation_template_id или presentation_template_name")
        needs_folder = self.presentation_template_name or any(
            s.name_pattern for s in self.sources.values()
        )
        if needs_folder and not self.folder_id:
            raise ValueError(
                "Используете name_pattern / presentation_template_name — "
                "тогда обязателен folder_id"
            )

    @classmethod
    def load(cls, path: str | Path) -> "ReportConfig":
        data = yaml.safe_load(Path(path).read_text())
        sources = {
            key: SheetSource(**val) for key, val in data.get("sources", {}).items()
        }
        return cls(
            name=data["name"],
            insight_slide_id=data["insight_slide_id"],
            sources=sources,
            insights=InsightRules(**data.get("insights", {})),
            static_placeholders=data.get("static_placeholders", {}),
            folder_id=data.get("folder_id"),
            presentation_template_id=data.get("presentation_template_id"),
            presentation_template_name=data.get("presentation_template_name"),
        )
