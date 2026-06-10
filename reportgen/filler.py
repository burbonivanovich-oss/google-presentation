"""Filler: набор функций, которые превращают Step + TemplateEntry в
batch-requests для Slides API.

Все функции возвращают list[dict] (requests) — композитор собирает их
в один batchUpdate.
"""
from __future__ import annotations

from typing import Iterable

from .template_index import ShapeRef, TemplateEntry


def replace_text_keep_style(shape: ShapeRef, new_text: str) -> list[dict]:
    """Заменить весь текст shape'а, сохранив стиль первого character.

    Трюк: insertText в начало → новый текст наследует style. Затем
    deleteText старого диапазона (он сместился вправо).
    """
    if shape.text == new_text:
        return []
    new_text = new_text or ""
    reqs: list[dict] = []
    if new_text:
        reqs.append({"insertText": {
            "objectId": shape.object_id,
            "insertionIndex": 0,
            "text": new_text,
        }})
    if shape.text:
        reqs.append({"deleteText": {
            "objectId": shape.object_id,
            "textRange": {
                "type": "FIXED_RANGE",
                "startIndex": len(new_text),
                "endIndex": len(new_text) + len(shape.text),
            },
        }})
    return reqs


def fill_slide_by_roles(entry: TemplateEntry, data: dict) -> list[dict]:
    """Заполняет shape'ы внутри одного слайда по карте role→value."""
    reqs: list[dict] = []
    for role, value in data.items():
        if role in ("dataframe", "categories_col", "series_cols"):
            continue
        shape = entry.shape_by_role.get(role)
        if not shape:
            continue
        reqs += replace_text_keep_style(shape, str(value))
    return reqs


def duplicate_with_ids(template_slide_id: str, new_slide_id: str,
                       shape_pairs: Iterable[tuple[str, str]]) -> dict:
    """Запрос duplicateObject с заранее заданными новыми ID.

    shape_pairs: [(old_shape_id, new_shape_id), ...]
    """
    object_ids = {template_slide_id: new_slide_id}
    for old, new in shape_pairs:
        object_ids[old] = new
    return {"duplicateObject": {
        "objectId": template_slide_id,
        "objectIds": object_ids,
    }}


def delete_slide(slide_id: str) -> dict:
    return {"deleteObject": {"objectId": slide_id}}


def move_slide_to_position(slide_id: str, position: int) -> dict:
    return {"updateSlidesPosition": {
        "slideObjectIds": [slide_id],
        "insertionIndex": position,
    }}
