"""Индекс шаблона: для каждой роли (cover, kpi_6, ...) находит первый
слайд-эталон в шаблоне и матчит его shape'ы к ролям из template_map.

Используется и для интерактивной диагностики (inspect-template), и
композитором при сборке отчёта.
"""
from __future__ import annotations

from dataclasses import dataclass

from googleapiclient.discovery import Resource

from .template_map import LAYOUTS, LayoutSpec, ShapeSlot

EMU_PER_INCH = 914400


@dataclass
class ShapeRef:
    object_id: str
    x_in: float
    y_in: float
    w_in: float
    h_in: float
    text: str

    def distance_to(self, slot: ShapeSlot) -> float:
        # манхэттенское расстояние
        return abs(self.x_in - slot.near_x) + abs(self.y_in - slot.near_y)


@dataclass
class SlideRef:
    object_id: str
    layout_name: str
    shapes: list[ShapeRef]


@dataclass
class TemplateEntry:
    """Готовый эталон для одной роли: ID слайда + role→ShapeRef."""
    role: str
    slide: SlideRef
    shape_by_role: dict[str, ShapeRef]


def fetch_slides(svc: Resource, presentation_id: str) -> tuple[dict, list[SlideRef]]:
    """Возвращает (raw pres, list[SlideRef] для каждого слайда шаблона)."""
    pres = svc.presentations().get(presentationId=presentation_id).execute()
    layouts = {lay["objectId"]: _layout_display_name(lay) for lay in pres.get("layouts", [])}

    slides: list[SlideRef] = []
    for s in pres.get("slides", []):
        layout_id = s.get("slideProperties", {}).get("layoutObjectId", "")
        layout_name = layouts.get(layout_id, "")
        shapes = []
        for el in s.get("pageElements", []):
            ref = _to_shape_ref(el)
            if ref:
                shapes.append(ref)
        slides.append(SlideRef(object_id=s["objectId"], layout_name=layout_name, shapes=shapes))
    return pres, slides


def _layout_display_name(layout: dict) -> str:
    props = layout.get("layoutProperties", {})
    return props.get("displayName") or props.get("name") or ""


def _to_shape_ref(el: dict) -> ShapeRef | None:
    """Делает ShapeRef из pageElement. Поддерживает shape, image и table —
    но возвращаем только то, что есть позиция."""
    transform = el.get("transform", {})
    size = el.get("size", {})
    if not size or not transform:
        return None
    # перевод EMU/UNIT → дюймы.
    def to_inch(v: dict | None) -> float:
        if not v:
            return 0.0
        mag = v.get("magnitude", 0)
        unit = v.get("unit", "EMU")
        if unit == "EMU":
            return mag / EMU_PER_INCH
        if unit == "PT":
            return mag / 72.0
        return float(mag)

    x_emu = transform.get("translateX", 0)
    y_emu = transform.get("translateY", 0)
    unit = transform.get("unit", "EMU")
    if unit == "EMU":
        x_in, y_in = x_emu / EMU_PER_INCH, y_emu / EMU_PER_INCH
    elif unit == "PT":
        x_in, y_in = x_emu / 72.0, y_emu / 72.0
    else:
        x_in, y_in = float(x_emu), float(y_emu)

    sx = transform.get("scaleX", 1) or 1
    sy = transform.get("scaleY", 1) or 1
    w_in = to_inch(size.get("width")) * sx
    h_in = to_inch(size.get("height")) * sy

    text = ""
    shape = el.get("shape")
    if shape:
        for te in shape.get("text", {}).get("textElements", []):
            run = te.get("textRun")
            if run:
                text += run.get("content", "")

    return ShapeRef(
        object_id=el["objectId"],
        x_in=x_in, y_in=y_in, w_in=w_in, h_in=h_in,
        text=text.strip(),
    )


def build_template_index(slides: list[SlideRef]) -> dict[str, TemplateEntry]:
    """Для каждой роли (cover, kpi_6, ...) выбираем ПЕРВЫЙ слайд с
    подходящим layout_name и сопоставляем shape'ы к ролям."""
    out: dict[str, TemplateEntry] = {}
    for spec in LAYOUTS:
        slide = _first_slide_for_layout(slides, spec)
        if not slide:
            continue
        shape_by_role = _match_shapes_to_slots(slide, spec.slots)
        out[spec.role] = TemplateEntry(
            role=spec.role, slide=slide, shape_by_role=shape_by_role
        )
    return out


def _first_slide_for_layout(slides: list[SlideRef], spec: LayoutSpec) -> SlideRef | None:
    # 1) Сначала по имени layout (если есть и оно осмысленное)
    patterns = [p.strip().lower() for p in spec.layout_name_patterns]
    for s in slides:
        name = s.layout_name.strip().lower()
        if name and any(p in name or name in p for p in patterns):
            return s
    # 2) Fallback: по сигнатурам в текстовом содержимом слайда.
    # Подходит, если все подстроки присутствуют в каком-то shape слайда.
    sigs = [s.lower() for s in spec.content_signatures]
    if not sigs:
        return None
    for s in slides:
        all_text = " | ".join(sh.text for sh in s.shapes).lower()
        if all(sig in all_text for sig in sigs):
            return s
    return None


def _match_shapes_to_slots(slide: SlideRef, slots: list[ShapeSlot]) -> dict[str, ShapeRef]:
    """Каждому slot подбираем shape с минимальным расстоянием. Уже
    привязанные shape'ы не используем повторно."""
    used: set[str] = set()
    result: dict[str, ShapeRef] = {}
    # сначала те, у кого есть text_hint и он совпал — приоритет
    pending = []
    for slot in slots:
        if slot.text_hint:
            hit = _best_by_hint(slide, slot, used)
            if hit:
                used.add(hit.object_id)
                result[slot.role] = hit
                continue
        pending.append(slot)
    # потом — по чистой близости
    for slot in pending:
        candidate = _best_by_distance(slide, slot, used)
        if candidate:
            used.add(candidate.object_id)
            result[slot.role] = candidate
    return result


def _best_by_hint(slide: SlideRef, slot: ShapeSlot, used: set[str]) -> ShapeRef | None:
    hint = (slot.text_hint or "").lower()
    candidates = [s for s in slide.shapes
                  if s.object_id not in used and hint and hint in s.text.lower()]
    if not candidates:
        return None
    return min(candidates, key=lambda s: s.distance_to(slot))


def _best_by_distance(slide: SlideRef, slot: ShapeSlot, used: set[str]) -> ShapeRef | None:
    candidates = [s for s in slide.shapes if s.object_id not in used]
    if not candidates:
        return None
    best = min(candidates, key=lambda s: s.distance_to(slot))
    # отбраковка: если расстояние > 2" — это случайно подобранный shape
    if best.distance_to(slot) > 2.0:
        return None
    return best
