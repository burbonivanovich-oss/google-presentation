"""Сборка нативной PPTX-презентации Q1 2026 на основе шаблона
«Контур Blue 2023_16x9_Montserrat». Использует layouts шаблона
(Отбивка / Заголовок+текст / Финальный) и добавляет KPI-карточки,
таблицы и графики через python-pptx — всё в брендовых цветах
(theme1.xml: accent1 #2291FF, accent2 #153177, dk2 #F1F1F1).
"""
from __future__ import annotations

import copy
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Cm, Emu, Pt

# Бренд Контур Blue (theme1.xml)
ACCENT = RGBColor(0x22, 0x91, 0xFF)
NAVY = RGBColor(0x15, 0x31, 0x77)
INK = RGBColor(0, 0, 0)
GRAY = RGBColor(0x5A, 0x65, 0x73)
LINE = RGBColor(0xE3, 0xE6, 0xEA)
PLATE = RGBColor(0xF1, 0xF1, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x26, 0xAD, 0x50)
RED = RGBColor(0x66, 0x14, 0x29)
ORANGE = RGBColor(0xFC, 0x76, 0x30)
ACCENT_SOFT = RGBColor(0xE7, 0xF2, 0xFF)

FONT = "Montserrat"

# Layouts по именам (как в шаблоне)
L_TITLE = "Титул 1"          # титул
L_DIV = "Отбивка"            # разделитель
L_HEAD_TEXT = "Заголовок в 1 строку + текст"  # универсальный холст
L_FINAL = "Финальный слайд_1"


def _layout(prs: Presentation, name: str):
    for L in prs.slide_layouts:
        if L.name == name:
            return L
    return prs.slide_layouts[0]


def _set_text(tf, text, *, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT,
              font=FONT, anchor=None):
    tf.word_wrap = True
    if anchor:
        tf.vertical_anchor = anchor
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color


def _add_text(slide, x, y, w, h, text, **kw):
    tb = slide.shapes.add_textbox(x, y, w, h)
    _set_text(tb.text_frame, text, **kw)
    return tb


def _delete_all_slides(prs: Presentation) -> None:
    """Удалить все 77 слайдов шаблона — оставить только мастер и layouts."""
    xml_slides = prs.slides._sldIdLst  # noqa: SLF001
    for sldId in list(xml_slides):
        rId = sldId.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        prs.part.drop_rel(rId)
        xml_slides.remove(sldId)


def _hide_placeholders(slide):
    """Очистить текст в плейсхолдерах layout (название/футер)."""
    for ph in list(slide.placeholders):
        try:
            ph.text_frame.clear()
        except Exception:
            pass


def _kicker(slide, x, y, text):
    _add_text(slide, x, y, Cm(20), Cm(0.7), text.upper(),
              size=10, bold=True, color=ACCENT)


def _title(slide, x, y, w, text, size=32):
    _add_text(slide, x, y, w, Cm(2.5), text,
              size=size, bold=True, color=INK)


def _foot(slide, source: str, page: str = ""):
    SLIDE_W = Cm(33.87); SLIDE_H = Cm(19.05)
    line = slide.shapes.add_connector(1, Cm(2), SLIDE_H - Cm(1.5),
                                       SLIDE_W - Cm(2), SLIDE_H - Cm(1.5))
    line.line.color.rgb = LINE
    _add_text(slide, Cm(2), SLIDE_H - Cm(1.3), Cm(20), Cm(0.8),
              source, size=9, color=GRAY)
    if page:
        tb = _add_text(slide, SLIDE_W - Cm(6), SLIDE_H - Cm(1.3),
                       Cm(4), Cm(0.8), page, size=9, color=GRAY,
                       align=PP_ALIGN.RIGHT)


def _card(slide, x, y, w, h, label, value, delta, delta_color=GRAY):
    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    bg.adjustments[0] = 0.08
    bg.fill.solid(); bg.fill.fore_color.rgb = PLATE
    bg.line.fill.background()
    bg.shadow.inherit = False
    pad = Cm(0.5)
    _add_text(slide, x + pad, y + pad, w - 2 * pad, Cm(0.7),
              label, size=10, color=GRAY, bold=True)
    _add_text(slide, x + pad, y + Cm(1.3), w - 2 * pad, Cm(1.8),
              value, size=22, color=INK, bold=True)
    _add_text(slide, x + pad, y + h - Cm(1.0), w - 2 * pad, Cm(0.7),
              delta, size=10, color=delta_color, bold=True)


def _cards_row(slide, y, items, top_y=None):
    """4 равных карточки в ряд по ширине слайда (с отступами по 2 см)."""
    SLIDE_W = Cm(33.87)
    margin = Cm(2); gap = Cm(0.4)
    n = len(items)
    total = SLIDE_W - margin * 2
    w = (total - gap * (n - 1)) / n
    h = Cm(4.0)
    yy = top_y or y
    for i, (lab, val, delta, color) in enumerate(items):
        x = margin + (w + gap) * i
        _card(slide, x, yy, w, h, lab, val, delta, color)
    return yy + h


def _add_table(slide, x, y, w, h, headers, rows):
    cols = len(headers); n = len(rows) + 1
    t = slide.shapes.add_table(n, cols, x, y, w, h).table
    for c, head in enumerate(headers):
        cell = t.cell(0, c)
        cell.fill.solid(); cell.fill.fore_color.rgb = ACCENT
        cell.margin_left = Cm(0.15); cell.margin_right = Cm(0.15)
        cell.margin_top = Cm(0.1); cell.margin_bottom = Cm(0.1)
        _set_text(cell.text_frame, str(head), size=10, bold=True,
                  color=WHITE, align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT,
                  anchor=MSO_ANCHOR.MIDDLE)
    for ri, row in enumerate(rows, 1):
        zebra = (ri % 2 == 0)
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            if zebra:
                cell.fill.solid(); cell.fill.fore_color.rgb = PLATE
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = WHITE
            cell.margin_left = Cm(0.15); cell.margin_right = Cm(0.15)
            cell.margin_top = Cm(0.06); cell.margin_bottom = Cm(0.06)
            color = INK
            txt = str(val)
            if txt.startswith("+") and "%" in txt: color = GREEN
            elif txt.startswith("−") and "%" in txt or txt.startswith("-") and "%" in txt: color = RED
            _set_text(cell.text_frame, txt, size=10,
                      align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT,
                      color=color, bold=(ci == 0),
                      anchor=MSO_ANCHOR.MIDDLE)
    return t


def _add_bar_chart(slide, x, y, w, h, cats, series, *, title=None):
    cd = CategoryChartData()
    cd.categories = cats
    for s in series:
        cd.add_series(s["name"], s["data"])
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                    x, y, w, h, cd)
    ch = gframe.chart
    ch.has_title = bool(title)
    if title:
        ch.chart_title.text_frame.text = title
        for r in ch.chart_title.text_frame.paragraphs[0].runs:
            r.font.name = FONT; r.font.size = Pt(11); r.font.bold = True
            r.font.color.rgb = GRAY
    ch.has_legend = len(series) > 1
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(9); ch.legend.font.name = FONT
    colors = [ACCENT, NAVY, ORANGE, GREEN]
    for i, ser in enumerate(ch.plots[0].series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = (RGBColor(0xC2, 0xC8, 0xD0) if series[i].get("plan") else colors[i % len(colors)])
        ser.format.line.fill.background()
    for axis in (ch.category_axis, ch.value_axis):
        try:
            axis.tick_labels.font.name = FONT
            axis.tick_labels.font.size = Pt(9)
            axis.tick_labels.font.color.rgb = GRAY
        except Exception:
            pass
    return ch


def _add_line_chart(slide, x, y, w, h, cats, series, *, title=None):
    cd = CategoryChartData()
    cd.categories = cats
    for s in series:
        cd.add_series(s["name"], s["data"])
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, w, h, cd)
    ch = gframe.chart
    ch.has_title = bool(title)
    if title:
        ch.chart_title.text_frame.text = title
        for r in ch.chart_title.text_frame.paragraphs[0].runs:
            r.font.name = FONT; r.font.size = Pt(11); r.font.bold = True
            r.font.color.rgb = GRAY
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(9); ch.legend.font.name = FONT
    colors = [ACCENT, NAVY, ORANGE, GREEN]
    for i, ser in enumerate(ch.plots[0].series):
        col = (RGBColor(0xC2, 0xC8, 0xD0) if series[i].get("plan")
               else colors[i % len(colors)])
        ser.format.line.color.rgb = col
        ser.format.line.width = Pt(2.5 if series[i].get("plan") else 3)
        try:
            if series[i].get("forecast"):
                ser.format.line.dash_style = 7  # dash
        except Exception:
            pass
        for marker in (ser.marker,):
            marker.style = 8  # circle
            marker.size = 6
            marker.format.fill.solid()
            marker.format.fill.fore_color.rgb = col
            marker.format.line.color.rgb = col
    for axis in (ch.category_axis, ch.value_axis):
        try:
            axis.tick_labels.font.name = FONT
            axis.tick_labels.font.size = Pt(9)
            axis.tick_labels.font.color.rgb = GRAY
        except Exception:
            pass
    return ch


# ── СБОРКА СЛАЙДОВ ────────────────────────────────────────────────

def build(template_path: str, out_path: str) -> None:
    prs = Presentation(template_path)
    _delete_all_slides(prs)
    SLIDE_W = prs.slide_width; SLIDE_H = prs.slide_height

    # 1. ТИТУЛ ───────────────────────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_TITLE))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(3.5), "Квартальный отчёт по продажам · Контур")
    _title(s, Cm(2), Cm(4.5), Cm(28), "Итоги Q1 2026", size=54)
    _add_text(s, Cm(2), Cm(8.5), Cm(28), Cm(2.5),
              "Розница · Общепит · Кассы · ОФД · Маркет · Контекстная реклама · прогноз 2026",
              size=14, color=GRAY)
    # 4 метки внизу
    metas = [("Период", "Q1 2026"), ("Сравнение", "Q4 2025 · Q1 2025"),
             ("Выручка факт", "68,2 млн ₽"), ("Выполнение плана", "92%")]
    for i, (k, v) in enumerate(metas):
        x = Cm(2 + i * 7.5)
        _add_text(s, x, Cm(13), Cm(7), Cm(0.6), k, size=9, color=GRAY, bold=True)
        _add_text(s, x, Cm(13.8), Cm(7), Cm(1.2), v, size=16, color=INK, bold=True)
    _foot(s, "Контур · продажи", "сформировано 13.06.2026")

    # 2. ОБЗОР ───────────────────────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(1.2), "Обзор")
    _title(s, Cm(2), Cm(2), Cm(28), "Ключевые метрики квартала")
    _cards_row(s, Cm(5.0), [
        ("Выручка факт", "68,2 млн", "92% от плана 74,6 млн", RED),
        ("Кол-во оплат", "6 859", "план 6 773 · 101%", GRAY),
        ("Средний чек", "9,9 тыс", "факт / кол-во оплат", GRAY),
        ("Недобор к плану", "−6,3 млн", "в основном март", RED),
    ])
    _add_text(s, Cm(2), Cm(9.6), Cm(15), Cm(0.7),
              "ПЛАН VS ФАКТ ПО МЕСЯЦАМ, МЛН ₽", size=9, color=GRAY, bold=True)
    _add_bar_chart(s, Cm(2), Cm(10.3), Cm(15), Cm(7), ["Январь", "Февраль", "Март"],
        [{"name": "План", "plan": True, "data": [22.33, 24.75, 27.47]},
         {"name": "Факт", "data": [21.12, 23.69, 23.42]}])
    _add_text(s, Cm(18), Cm(9.6), Cm(14), Cm(0.7),
              "ВЫРУЧКА ПО НАПРАВЛЕНИЯМ, Q1 2026", size=9, color=GRAY, bold=True)
    _add_table(s, Cm(18), Cm(10.3), Cm(14), Cm(7),
               ["Направление", "Выручка", "Доля"],
               [["Розница", "~32,5 млн", "48%"],
                ["Кассы", "~22,2 млн", "33%"],
                ["Общепит", "~12,8 млн", "19%"],
                ["в т.ч. ОФД (продукт)", "11,7 млн", "17%"]])
    _foot(s, "Источник: Царь свод (план-факт) + Царь продажи")

    # 3. ДИНАМИКА ────────────────────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(1.2), "Помесячная динамика")
    _title(s, Cm(2), Cm(2), Cm(28), "Бизнес-юнит: план и факт по месяцам")
    _add_bar_chart(s, Cm(2), Cm(4.6), Cm(16), Cm(12),
                   ["Январь", "Февраль", "Март"],
                   [{"name": "План, млн ₽", "plan": True, "data": [22.33, 24.75, 27.47]},
                    {"name": "Факт, млн ₽", "data": [21.12, 23.69, 23.42]}],
                   title="План vs факт")
    _add_table(s, Cm(19), Cm(4.6), Cm(13), Cm(5.5),
               ["Месяц", "План", "Факт", "%"],
               [["Январь", "22,33", "21,12", "95"],
                ["Февраль", "24,75", "23,69", "96"],
                ["Март", "27,47", "23,42", "85"],
                ["Итого", "74,56", "68,23", "92"]])
    notes = [
        "Январь–февраль удерживаем 95–96% плана.",
        "Март — провал до 85%: план вырос до 27,5 млн, факт — на уровне февраля.",
        "Февраль перевыполнили по оплатам (2 510 vs 2 222) — выручку тянет вниз средний чек."
    ]
    for i, n in enumerate(notes):
        _add_text(s, Cm(19), Cm(11 + i * 1.6), Cm(13), Cm(1.4),
                  "• " + n, size=11, color=INK)
    _foot(s, "Источник: Царь свод все разрезы")

    # 4. РАЗДЕЛИТЕЛЬ — НАПРАВЛЕНИЯ ──────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_DIV))
    _hide_placeholders(s)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid(); bg.fill.fore_color.rgb = NAVY
    bg.line.fill.background()
    _add_text(s, SLIDE_W - Cm(8), Cm(1.5), Cm(7), Cm(8), "01",
              size=140, color=RGBColor(0x2A, 0x4C, 0x8E), bold=True, align=PP_ALIGN.RIGHT)
    _kicker(s, Cm(2), Cm(6), "Раздел")
    _add_text(s, Cm(2), Cm(7), Cm(20), Cm(4), "Направления",
              size=60, color=WHITE, bold=True)
    _add_text(s, Cm(2), Cm(12), Cm(22), Cm(3),
              "Розница · Общепит · Кассы · ОФД — выручка, динамика QoQ/YoY и ключевые тарифы.",
              size=14, color=RGBColor(0xC9, 0xDB, 0xF5))

    # 5–8. НАПРАВЛЕНИЯ ──────────────────────────────────────────
    def direction_slide(name, cards, tariffs, source):
        s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
        _hide_placeholders(s)
        _kicker(s, Cm(2), Cm(1.2), "Направление")
        _title(s, Cm(2), Cm(2), Cm(28), name)
        _cards_row(s, Cm(5.0), cards)
        _add_text(s, Cm(2), Cm(9.6), Cm(20), Cm(0.7),
                  "ТАРИФЫ В Q1 2026, МЛН ₽", size=9, color=GRAY, bold=True)
        _add_table(s, Cm(2), Cm(10.3), Cm(30), Cm(7),
                   ["Тариф", "Q1 2025", "Q4 2025", "Q1 2026", "QoQ", "YoY"], tariffs)
        _foot(s, source)

    direction_slide("Розница",
        [("Выручка Q1 2026", "~32,5 млн", "48% БЮ", GRAY),
         ("Доля направления", "48%", "крупнейшее", GRAY),
         ("Топ-тариф", "Оптим. Розница", "6,66 млн · +9% QoQ", GREEN),
         ("Под давлением", "Все госсистемы", "3,47 млн · −13% YoY", RED)],
        [["Оптимальный Розница", "7,13", "6,13", "6,66", "+9%", "−7%"],
         ["Все госсистемы", "3,97", "3,80", "3,47", "−9%", "−13%"],
         ["Базовый Розница", "1,89", "1,59", "1,41", "−11%", "−25%"],
         ["Премиум Розница", "1,64", "1,13", "0,95", "−16%", "−42%"],
         ["Маркировка", "1,39", "3,26", "1,86", "−43%", "+34%"]],
        "Источник: Царь продажи · Бизнес-юнит = Госсистемы для розницы")

    direction_slide("Общепит",
        [("Выручка Q1 2026", "~12,8 млн", "19% БЮ", GRAY),
         ("Доля направления", "19%", "третье по объёму", GRAY),
         ("Топ-тариф", "Оптим. Общепит", "1,29 млн · +44% QoQ", GREEN),
         ("Динамика", "растёт", "+8% YoY по топ-тарифу", GREEN)],
        [["Оптимальный Общепит", "1,20", "0,89", "1,29", "+44%", "+8%"],
         ["КМ модификатор ЕГАИС", "1,17", "0,99", "1,09", "+10%", "−7%"],
         ["Базовый Общепит", "—", "—", "0,49", "·", "·"]],
        "Источник: Царь продажи · Бизнес-юнит = Госсистемы для общепита")

    direction_slide("Кассы",
        [("Выручка Q1 2026", "~22,2 млн", "33% БЮ", GRAY),
         ("Доля направления", "33%", "второе по объёму", GRAY),
         ("Топ-тариф", "Терминал", "2,75 млн · +29% QoQ", GREEN),
         ("Растущее", "Перерег. ККТ", "1,56 млн · +48% YoY", GREEN)],
        [["Терминал", "2,63", "2,14", "2,75", "+29%", "+5%"],
         ["Модификатор Маркировка", "2,46", "2,30", "2,55", "+11%", "+4%"],
         ["Лицензия Атол ИТС", "1,23", "1,39", "1,59", "+14%", "+29%"],
         ["Перерегистрация ККТ", "1,05", "1,30", "1,56", "+20%", "+48%"],
         ["Атол ИТС", "1,48", "1,00", "1,24", "+24%", "−17%"],
         ["Сканер", "1,21", "1,37", "1,16", "−15%", "−5%"]],
        "Источник: Царь продажи · Бизнес-юнит = Кассовики")

    # ОФД — отдельный с графиком
    s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(1.2), "Направление · продуктовый срез")
    _title(s, Cm(2), Cm(2), Cm(28), "ОФД")
    _cards_row(s, Cm(5.0), [
        ("Выручка Q1 2026", "11,69 млн", "+4,2% к Q4 2025", GREEN),
        ("YoY", "−4,9%", "было 12,28 млн", RED),
        ("Кол-во оплат", "2 122", "+49 к Q4", GREEN),
        ("Средний чек", "5,5 тыс", "11,69 / 2 122", GRAY),
    ])
    _add_text(s, Cm(2), Cm(9.6), Cm(15), Cm(0.7),
              "ВЫРУЧКА ПО МЕСЯЦАМ Q1, МЛН ₽", size=9, color=GRAY, bold=True)
    _add_bar_chart(s, Cm(2), Cm(10.3), Cm(15), Cm(7),
                   ["Январь", "Февраль", "Март"],
                   [{"name": "Выручка", "data": [3.94, 3.89, 3.85]}])
    _add_text(s, Cm(18), Cm(9.6), Cm(14), Cm(0.7),
              "ТОП ТАРИФОВ, МЛН ₽", size=9, color=GRAY, bold=True)
    _add_table(s, Cm(18), Cm(10.3), Cm(14), Cm(7),
               ["Тариф", "Q1 2026"],
               [["ОФД-36", "3,93"], ["ОФД-15", "3,73"], ["ОФД-13", "3,38"],
                ["ОФД-24", "0,46"], ["Прочие ОФД", "0,19"]])
    _foot(s, "Источник: Царь продажи · Тариф содержит «ОФД»")

    # 9. ТОП-20 ТАРИФОВ ─────────────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(1.2), "Тарифы")
    _title(s, Cm(2), Cm(2), Cm(28), "Топ-20 тарифов: QoQ и YoY")
    rows = [
        ["Прочие", "26,63", "21,26", "14,96", "−30%", "−44%"],
        ["Оптимальный Розница", "7,13", "6,13", "6,66", "+9%", "−7%"],
        ["ОФД-36", "4,77", "4,79", "3,93", "−18%", "−18%"],
        ["ОФД-15", "4,17", "3,00", "3,73", "+24%", "−11%"],
        ["Все госсистемы", "3,97", "3,80", "3,47", "−9%", "−13%"],
        ["ОФД-13", "2,72", "2,73", "3,38", "+24%", "+24%"],
        ["Настройка Контур.Маркет", "1,59", "1,45", "2,93", "+102%", "+85%"],
        ["Терминал", "2,63", "2,14", "2,75", "+29%", "+5%"],
        ["Модификатор Маркировка", "2,46", "2,30", "2,55", "+11%", "+4%"],
        ["Маркировка", "1,39", "3,26", "1,86", "−43%", "+34%"],
        ["Лицензия Атол ИТС", "1,23", "1,39", "1,59", "+14%", "+29%"],
        ["Перерегистрация ККТ", "1,05", "1,30", "1,56", "+20%", "+48%"],
        ["Базовый Розница", "1,89", "1,59", "1,41", "−11%", "−25%"],
        ["Оптимальный Общепит", "1,20", "0,89", "1,29", "+44%", "+8%"],
        ["Атол ИТС", "1,48", "1,00", "1,24", "+24%", "−17%"],
        ["Сертификат", "0,86", "0,66", "1,23", "+88%", "+43%"],
        ["Сканер", "1,21", "1,37", "1,16", "−15%", "−5%"],
        ["КМ модификатор ЕГАИС", "1,17", "0,99", "1,09", "+10%", "−7%"],
        ["Премиум Розница", "1,64", "1,13", "0,95", "−16%", "−42%"],
        ["Базовый Услуги", "1,04", "0,67", "0,86", "+29%", "−18%"],
    ]
    _add_table(s, Cm(2), Cm(4.0), Cm(30), Cm(13.5),
               ["Тариф", "Q1 2025", "Q4 2025", "Q1 2026", "QoQ", "YoY"], rows)
    _foot(s, "Источник: Царь продажи · группировка по тарифу, млн ₽")

    # 10. РАЗДЕЛИТЕЛЬ — ПРОГНОЗ ─────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_DIV))
    _hide_placeholders(s)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
    _add_text(s, SLIDE_W - Cm(8), Cm(1.5), Cm(7), Cm(8), "02",
              size=140, color=RGBColor(0x2A, 0x4C, 0x8E), bold=True, align=PP_ALIGN.RIGHT)
    _kicker(s, Cm(2), Cm(6), "Раздел")
    _add_text(s, Cm(2), Cm(7), Cm(22), Cm(4), "Прогноз 2026",
              size=60, color=WHITE, bold=True)
    _add_text(s, Cm(2), Cm(11.5), Cm(22), Cm(1.5), "Маркет и ОФД",
              size=24, color=RGBColor(0x8F, 0xC4, 0xFF))
    _add_text(s, Cm(2), Cm(13.5), Cm(22), Cm(3),
              "Ансамбль моделей: plan-pacing · сезонный YoY · ETS · регрессия. "
              "Факт по май, прогноз июнь–декабрь, сценарии низ/база/верх.",
              size=14, color=RGBColor(0xC9, 0xDB, 0xF5))

    # 11. МАРКЕТ — ПРОГНОЗ ──────────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(1.2), "Прогноз · выручка")
    _title(s, Cm(2), Cm(2), Cm(28), "Маркет — выполнение плана 2026")
    _cards_row(s, Cm(5.0), [
        ("Годовой план", "119,0 млн", "Продукт = Маркет", GRAY),
        ("Факт за 5 мес", "44,5 млн", "91% от плана периода", RED),
        ("YoY к 2025", "−21%", "падение спроса", RED),
        ("Прогноз года", "86%", "102 млн · 78–91%", RED),
    ])
    months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
    _add_line_chart(s, Cm(2), Cm(10), Cm(30), Cm(7.5), months,
        [{"name": "План", "plan": True, "data": [9.78,9.46,10.98,10.45,8.36,10.32,10.13,9.01,8.79,9.95,9.71,12.10]},
         {"name": "Факт", "data": [8.24,9.20,10.26,8.99,7.78,None,None,None,None,None,None,None]},
         {"name": "Прогноз", "forecast": True, "data": [None,None,None,None,7.78,9.1,8.4,7.2,7.6,7.8,7.4,9.9]}],
        title="План / факт / прогноз, млн ₽")
    _foot(s, "Источник: план-факт «Продукт = Маркет» · ансамбль моделей")

    # 12. ОФД — ПРОГНОЗ ──────────────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(1.2), "Прогноз · выручка")
    _title(s, Cm(2), Cm(2), Cm(28), "ОФД — выполнение плана 2026")
    _cards_row(s, Cm(5.0), [
        ("Годовой план", "37,5 млн", "проект п453", GRAY),
        ("Факт за 5 мес", "18,9 млн", "123% от плана периода", GREEN),
        ("YoY к 2025", "−16%", "но план консервативен", GRAY),
        ("Прогноз года", "115%", "43,3 млн · 103–124%", GREEN),
    ])
    _add_line_chart(s, Cm(2), Cm(10), Cm(30), Cm(7.5), months,
        [{"name": "План", "plan": True, "data": [2.81,3.11,3.18,3.24,2.97,3.04,3.03,2.99,2.90,3.28,3.20,3.79]},
         {"name": "Факт", "data": [4.05,4.12,4.18,3.71,2.84,None,None,None,None,None,None,None]},
         {"name": "Прогноз", "forecast": True, "data": [None,None,None,None,2.84,3.6,3.4,3.1,3.1,3.4,3.2,4.5]}],
        title="План / факт / прогноз, млн ₽")
    _foot(s, "Источник: план-факт проект п453 · ансамбль моделей")

    # 13. РАЗДЕЛИТЕЛЬ — РЕКЛАМА ────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_DIV))
    _hide_placeholders(s)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
    _add_text(s, SLIDE_W - Cm(8), Cm(1.5), Cm(7), Cm(8), "03",
              size=140, color=RGBColor(0x2A, 0x4C, 0x8E), bold=True, align=PP_ALIGN.RIGHT)
    _kicker(s, Cm(2), Cm(6), "Раздел")
    _add_text(s, Cm(2), Cm(7), Cm(22), Cm(4), "Контекстная реклама",
              size=54, color=WHITE, bold=True)
    _add_text(s, Cm(2), Cm(13), Cm(22), Cm(3),
              "Расход, CPL, CPO и привлечённые оплаты по Маркету и ОФД. "
              "Данные с дашбордов кампаний, янв 2025 – май 2026.",
              size=14, color=RGBColor(0xC9, 0xDB, 0xF5))

    # 14–15. КОНТЕКСТ ───────────────────────────────────────────
    ad_months = ["я25","ф","м","а","м","и","и","а","с","о","н","д","я26","ф","м","а","м","и"]

    s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(1.2), "Контекст · Маркет")
    _title(s, Cm(2), Cm(2), Cm(28), "Маркет — эффективность контекста")
    _cards_row(s, Cm(5.0), [
        ("Расход всего", "14,7 млн", "янв 2025 – май 2026", GRAY),
        ("Оплаты", "315", "привлечено", GRAY),
        ("CPO", "46,8 тыс", "высокая стоимость", RED),
        ("CPL средн.", "10,0 тыс", "снижается к 2026", GREEN),
    ])
    _add_bar_chart(s, Cm(2), Cm(10.3), Cm(15), Cm(7), ad_months,
        [{"name": "Расход, млн ₽",
          "data": [1.43,1.46,1.29,1.65,1.14,0.89,0.68,0.47,0.62,0.83,0.75,0.64,0.57,0.56,0.70,0.45,0.47,0.13]}],
        title="Расход по месяцам, млн ₽")
    _add_line_chart(s, Cm(18), Cm(10.3), Cm(14), Cm(7), ad_months,
        [{"name": "CPL, тыс ₽",
          "data": [16.2,14.3,14.7,13.0,9.7,8.5,6.6,8.4,8.2,11.0,11.2,9.7,7.7,9.7,10.5,6.0,4.5,4.6]}],
        title="CPL по месяцам, тыс ₽")
    _foot(s, "Источник: дашборд Контекст-Маркет · суммы сверены с «Итого»")

    s = prs.slides.add_slide(_layout(prs, L_HEAD_TEXT))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(1.2), "Контекст · ОФД")
    _title(s, Cm(2), Cm(2), Cm(28), "ОФД — эффективность контекста")
    _cards_row(s, Cm(5.0), [
        ("Расход всего", "7,1 млн", "янв 2025 – май 2026", GRAY),
        ("Оплаты", "397", "привлечено", GRAY),
        ("CPO", "17,9 тыс", "в 2,6× дешевле Маркета", GREEN),
        ("CPL средн.", "4,5 тыс", "эффективнее", GREEN),
    ])
    _add_bar_chart(s, Cm(2), Cm(10.3), Cm(15), Cm(7), ad_months,
        [{"name": "Расход, млн ₽",
          "data": [0.35,0.42,0.46,0.59,0.52,0.43,0.36,0.37,0.41,0.46,0.43,0.39,0.42,0.39,0.42,0.39,0.26,0.05]}],
        title="Расход по месяцам, млн ₽")
    _add_line_chart(s, Cm(18), Cm(10.3), Cm(14), Cm(7), ad_months,
        [{"name": "CPL, тыс ₽",
          "data": [2.6,4.1,5.0,8.6,10.1,9.3,4.7,4.7,4.6,5.2,5.0,3.9,3.5,3.5,3.6,3.2,4.0,1.9]}],
        title="CPL по месяцам, тыс ₽")
    _foot(s, "Источник: дашборд Контекст-ОФД · суммы сверены с «Итого»")

    # 16. ФИНАЛ ─────────────────────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_FINAL))
    _hide_placeholders(s)
    _kicker(s, Cm(2), Cm(2), "Спасибо")
    _add_text(s, Cm(2), Cm(3), Cm(20), Cm(4),
              "Итоги Q1 2026", size=44, color=INK, bold=True)
    _add_text(s, Cm(2), Cm(8.5), Cm(28), Cm(2),
              "Сводка собрана автоматически из 5 xlsx и дашбордов рекламы.",
              size=14, color=GRAY)
    _add_text(s, Cm(2), Cm(11), Cm(28), Cm(0.7),
              "ВЫВОДЫ", size=10, color=ACCENT, bold=True)
    bullets = [
        "Q1: 92% плана; январь–февраль 95–96%, март провален (85%).",
        "Маркет: −21% YoY, прогноз 2026 ≈ 86% годового плана — недовыполнение.",
        "ОФД: факт YTD 123% плана периода, прогноз 2026 ≈ 115% — перевыполнение.",
        "Реклама: ОФД-контекст в 2,6× эффективнее Маркета по CPO (17,9 тыс vs 46,8 тыс)."
    ]
    for i, b in enumerate(bullets):
        _add_text(s, Cm(2), Cm(12 + i * 1.3), Cm(30), Cm(1.2),
                  "• " + b, size=13, color=INK)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)


if __name__ == "__main__":
    import sys
    template = sys.argv[1] if len(sys.argv) > 1 else "Шаблон презентации Контур Blue 2023_16x9_Montserrat (2).pptx"
    out = sys.argv[2] if len(sys.argv) > 2 else "Итоги Q1 2026 — Контур.pptx"
    build(template, out)
    print(f"OK → {out}")
