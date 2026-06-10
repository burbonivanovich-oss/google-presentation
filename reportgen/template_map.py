"""Карта известных layout-ов брендбука Контура → роли shape'ов.

Координаты — в дюймах (как в PPTX), берутся из исходного шаблона.
Идентификация shape'а в слайде идёт по близости (top-left) к ожидаемой
точке + опциональному совпадению "сигнатурного" текста.

При импорте PPTX в Google Slides структура shape'ов сохраняется, имена
layout-ов — тоже (Google показывает их как displayName). Поэтому первый
слайд с подходящим layoutName используется как эталон.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ShapeSlot:
    """Один shape в слайде с его ролью."""
    role: str
    near_x: float  # дюймы
    near_y: float
    # подсказка для матчинга: если текст shape содержит эту подстроку — приоритет
    text_hint: str | None = None


@dataclass
class LayoutSpec:
    """Тип слайда → как его искать в шаблоне + какие в нём роли."""
    role: str  # внутреннее имя типа: cover, section, kpi_2, kpi_6 и т.д.
    layout_name_patterns: list[str]  # совпадение по displayName layout
    slots: list[ShapeSlot] = field(default_factory=list)


# Все типы слайдов, которые planner может использовать.
# Координаты сверены с разбором template.pptx.
LAYOUTS: list[LayoutSpec] = [
    LayoutSpec(
        role="cover",
        layout_name_patterns=["Титул 1"],
        slots=[
            ShapeSlot("title", 0.65, 0.64, "Тема выступления"),
            ShapeSlot("name", 0.65, 4.61, "Имя Фамилия"),
            ShapeSlot("subtitle", 0.65, 5.03, "должность"),
            ShapeSlot("url", 8.95, 6.44, "kontur.ru"),
        ],
    ),
    LayoutSpec(
        role="section",
        layout_name_patterns=["Отбивка"],
        slots=[
            ShapeSlot("title", 0.64, 0.64, "Гайдлайн"),
            ShapeSlot("body", 1.04, 3.24, "Чтобы начать"),
        ],
    ),
    LayoutSpec(
        role="kpi_2",
        layout_name_patterns=["2 цифры крупно"],
        slots=[
            ShapeSlot("title", 0.64, 0.64, "цифры"),
            ShapeSlot("value_1", 0.64, 3.15),
            ShapeSlot("desc_1", 0.64, 4.74, "Здесь может быть"),
            ShapeSlot("value_2", 6.69, 3.15),
            ShapeSlot("desc_2", 6.69, 4.74, "Здесь может быть"),
        ],
    ),
    LayoutSpec(
        role="kpi_6",
        layout_name_patterns=["6 важных цифр"],  # в шаблоне иногда с trailing space
        slots=[
            # сетка 3 колонки × 2 ряда; в каждой ячейке value сверху, desc снизу
            ShapeSlot("value_1", 0.64, 1.51),
            ShapeSlot("desc_1", 0.64, 2.57, "Здесь может быть"),
            ShapeSlot("value_2", 4.79, 1.51),
            ShapeSlot("desc_2", 4.79, 2.57, "Здесь может быть"),
            ShapeSlot("value_3", 8.91, 1.51),
            ShapeSlot("desc_3", 8.91, 2.57, "Здесь может быть"),
            ShapeSlot("value_4", 0.64, 4.17),
            ShapeSlot("desc_4", 0.64, 5.24, "Здесь может быть"),
            ShapeSlot("value_5", 4.79, 4.17),
            ShapeSlot("desc_5", 4.79, 5.24, "Здесь может быть"),
            ShapeSlot("value_6", 8.91, 4.17),
            ShapeSlot("desc_6", 8.91, 5.24, "Здесь может быть"),
        ],
    ),
    LayoutSpec(
        role="big_quote",
        layout_name_patterns=["Крупный тезис"],
        slots=[
            ShapeSlot("preface", 0.64, 0.64, "Выручка"),
            ShapeSlot("value", 3.47, 3.96, "26,4"),  # 220pt
            ShapeSlot("unit", 10.12, 6.02, "млрд"),  # 47pt
        ],
    ),
    LayoutSpec(
        role="facts_3",
        layout_name_patterns=["3 факта"],
        slots=[
            ShapeSlot("title", 0.64, 0.64, "факта"),
            ShapeSlot("num_1", 0.64, 3.81),
            ShapeSlot("desc_1", 0.64, 4.96, "Здесь может быть"),
            ShapeSlot("num_2", 4.59, 3.82),
            ShapeSlot("desc_2", 4.59, 4.95, "Здесь может быть"),
            ShapeSlot("num_3", 8.52, 3.82),
            ShapeSlot("desc_3", 8.52, 4.95, "Здесь может быть"),
        ],
    ),
    LayoutSpec(
        role="cards_3",
        layout_name_patterns=["3 карточки факта"],
        slots=[
            ShapeSlot("title", 0.64, 0.64),
            ShapeSlot("num_1", 0.96, 2.77),
            ShapeSlot("subtitle_1", 0.96, 3.73, "Подзаголовок"),
            ShapeSlot("desc_1", 0.96, 4.61, "дополнительный"),
            ShapeSlot("num_2", 5.08, 2.77),
            ShapeSlot("subtitle_2", 5.08, 3.73, "Подзаголовок"),
            ShapeSlot("desc_2", 5.08, 4.61, "дополнительный"),
            ShapeSlot("num_3", 9.19, 2.77),
            ShapeSlot("subtitle_3", 9.19, 3.73, "Подзаголовок"),
            ShapeSlot("desc_3", 9.21, 4.61, "дополнительный"),
        ],
    ),
    LayoutSpec(
        role="text_section",
        layout_name_patterns=["Заголовок в 1 строку + текст"],
        slots=[
            ShapeSlot("title", 0.64, 0.64, "Короткий"),
            ShapeSlot("body", 0.64, 2.77, "поясняющий"),
        ],
    ),
    LayoutSpec(
        role="final",
        layout_name_patterns=["Финальный слайд_1", "Финальный слайд"],
        slots=[
            ShapeSlot("title", 0.64, 0.64, "Благодарю"),
            ShapeSlot("name", 0.64, 4.56, "Имя Фамилия"),
            ShapeSlot("subtitle", 0.64, 4.98, "должность"),
        ],
    ),
    # blank — пустой слайд под нашу таблицу или график. Берём "Текстовый"
    # как «самый пустой» layout (если будет — иначе любой с минимумом контента).
    LayoutSpec(
        role="blank",
        layout_name_patterns=["Текстовый"],
        slots=[],
    ),
]


def by_role(role: str) -> LayoutSpec | None:
    for spec in LAYOUTS:
        if spec.role == role:
            return spec
    return None
