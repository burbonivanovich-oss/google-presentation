"""Карта layout-ов брендбука Контура → роли.

После анализа реальных отчётов (Петрова/Порубов) видно: основная
рабочая лошадка отчёта — layout «Заголовок в 1 строку + текст». В нём
располагается заголовок-тезис, рядом график/таблица, под ним выводы.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ShapeSlot:
    role: str
    near_x: float  # дюймы
    near_y: float
    text_hint: str | None = None


@dataclass
class LayoutSpec:
    role: str
    layout_name_patterns: list[str]
    slots: list[ShapeSlot] = field(default_factory=list)
    content_signatures: list[str] = field(default_factory=list)


LAYOUTS: list[LayoutSpec] = [
    LayoutSpec(
        role="cover",
        layout_name_patterns=["Титул 1", "Title", "Cover"],
        content_signatures=["Тема выступления"],
        slots=[
            ShapeSlot("title", 0.65, 0.64, "Тема выступления"),
            ShapeSlot("name", 0.65, 4.61, "Имя Фамилия"),
            ShapeSlot("subtitle", 0.65, 5.03, "должность"),
            ShapeSlot("url", 8.95, 6.44, "kontur.ru"),
        ],
    ),
    # Основной слайд: заголовок + большой текст. Сюда наш plan кладёт
    # тезис-заголовок и выводы списком (через \n). Если нужен график —
    # composer добавит chart рядом с текстовым блоком.
    LayoutSpec(
        role="slide_chart",
        layout_name_patterns=["Заголовок в 1 строку + текст"],
        content_signatures=["Короткий", "поясняющий"],
        slots=[
            ShapeSlot("title", 0.64, 0.64, "Короткий"),
            ShapeSlot("body", 0.64, 2.77, "поясняющий"),
        ],
    ),
    # Текст без графика. Используем тот же layout, просто без вставки chart.
    LayoutSpec(
        role="slide_text",
        layout_name_patterns=["Заголовок в 1 строку + текст"],
        content_signatures=["Короткий", "поясняющий"],
        slots=[
            ShapeSlot("title", 0.64, 0.64, "Короткий"),
            ShapeSlot("body", 0.64, 2.77, "поясняющий"),
        ],
    ),
    LayoutSpec(
        role="final",
        layout_name_patterns=["Финальный слайд_1", "Финальный слайд", "Final"],
        content_signatures=["Благодарю", "за внимание"],
        slots=[
            ShapeSlot("title", 0.64, 0.64, "Благодарю"),
            ShapeSlot("name", 0.64, 4.56, "Имя Фамилия"),
            ShapeSlot("subtitle", 0.64, 4.98, "должность"),
        ],
    ),
]


def by_role(role: str) -> LayoutSpec | None:
    for spec in LAYOUTS:
        if spec.role == role:
            return spec
    return None
