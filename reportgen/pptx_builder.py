"""Сборка нативной PPTX Q1 2026 на основе шаблона «Контур Blue Montserrat».

Принципы (чтобы было «по стандарту» шаблона):
  • заголовок — это title-ПЛЕЙСХОЛДЕР layout'а (правильные шрифт/позиция),
    а не свой текстбокс;
  • все НЕиспользуемые плейсхолдеры удаляются — иначе в экспорте торчат
    подсказки «Введите заголовок / Введите текст» и картинка-заглушка;
  • контент кладётся по сетке шаблона: поля 1,63 см, заголовок сверху,
    контент-зона ниже;
  • цвета/шрифт — из темы шаблона (theme1.xml): accent1 #2291FF, Montserrat.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

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
SOFT = RGBColor(0xE7, 0xF2, 0xFF)
GRAYBAR = RGBColor(0xC2, 0xC8, 0xD0)
FONT = "Montserrat"

# сетка шаблона
MX = Cm(1.63)            # левое/правое поле
CONTENT_W = Cm(30.6)     # ширина контента
SLIDE_W = Cm(33.87)
SLIDE_H = Cm(19.05)

L_TITLE = "Титул 1"
L_DIV = "Отбивка"
L_BODY = "Заголовок в 1 строку + текст"
L_FINAL = "Финальный слайд_1"


def _layout(prs, name):
    for L in prs.slide_layouts:
        if L.name == name:
            return L
    return prs.slide_layouts[0]


def _del_slides(prs):
    lst = prs.slides._sldIdLst  # noqa: SLF001
    for sid in list(lst):
        prs.part.drop_rel(sid.attrib[qn('r:id')])
        lst.remove(sid)


def _strip_placeholders(slide, keep_idx=()):
    """Удалить все плейсхолдеры, кроме перечисленных idx (убирает подсказки)."""
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx not in keep_idx:
            ph._element.getparent().remove(ph._element)  # noqa: SLF001


def _fill_title(slide, text, *, size=30, color=INK):
    """Заполнить title-плейсхолдер layout'а (шрифт/позиция — из шаблона)."""
    title = slide.shapes.title
    tf = title.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = text
    r.font.name = FONT; r.font.size = Pt(size); r.font.bold = True
    r.font.color.rgb = color
    return title


def _txt(slide, x, y, w, h, text, *, size=12, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = ln
        r.font.name = FONT; r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = color
    return tb


def _send_back(slide, shape):
    tree = slide.shapes._spTree  # noqa: SLF001
    tree.remove(shape._element)
    tree.insert(2, shape._element)


def _foot(slide, source, page=""):
    ln = slide.shapes.add_connector(1, MX, SLIDE_H - Cm(1.4),
                                    SLIDE_W - MX, SLIDE_H - Cm(1.4))
    ln.line.color.rgb = LINE; ln.line.width = Pt(0.75)
    _txt(slide, MX, SLIDE_H - Cm(1.2), Cm(22), Cm(0.7), source, size=9, color=GRAY)
    if page:
        _txt(slide, SLIDE_W - Cm(8), SLIDE_H - Cm(1.2), Cm(6.37), Cm(0.7),
             page, size=9, color=GRAY, align=PP_ALIGN.RIGHT)


def _card(slide, x, y, w, h, label, value, delta, dcolor=GRAY, vsize=22):
    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    bg.adjustments[0] = 0.07
    bg.fill.solid(); bg.fill.fore_color.rgb = PLATE
    bg.line.fill.background(); bg.shadow.inherit = False
    pad = Cm(0.55)
    _txt(slide, x + pad, y + pad, w - 2 * pad, Cm(1.2), label, size=10, color=GRAY, bold=True)
    _txt(slide, x + pad, y + Cm(1.5), w - 2 * pad, h - Cm(2.6), value, size=vsize, color=INK, bold=True)
    _txt(slide, x + pad, y + h - pad - Cm(0.7), w - 2 * pad, Cm(0.7), delta, size=10, color=dcolor, bold=True)


def _cards4(slide, y, items, h=Cm(4.0)):
    gap = Cm(0.4)
    w = (CONTENT_W - gap * 3) / 4
    for i, it in enumerate(items):
        _card(slide, MX + (w + gap) * i, y, w, h, *it)
    return y + h


def _table(slide, x, y, w, headers, rows, *, fs=10, row_h=Cm(0.78)):
    n = len(rows) + 1; cols = len(headers)
    gf = slide.shapes.add_table(n, cols, x, y, w, row_h * n)
    t = gf.table
    for r in t.rows:
        r.height = row_h
    # выключаем встроенную полосатость стиля — зебру/шапку красим сами
    tblPr = t._tbl.tblPr  # noqa: SLF001
    tblPr.set('firstRow', '0'); tblPr.set('bandRow', '0')
    for c, head in enumerate(headers):
        cell = t.cell(0, c)
        cell.fill.solid(); cell.fill.fore_color.rgb = ACCENT
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Cm(0.2); cell.margin_right = Cm(0.2)
        cell.margin_top = Cm(0.08); cell.margin_bottom = Cm(0.08)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
        r = p.add_run(); r.text = str(head)
        r.font.name = FONT; r.font.size = Pt(fs); r.font.bold = True; r.font.color.rgb = WHITE
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            cell.fill.solid(); cell.fill.fore_color.rgb = PLATE if ri % 2 == 0 else WHITE
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Cm(0.2); cell.margin_right = Cm(0.2)
            cell.margin_top = Cm(0.05); cell.margin_bottom = Cm(0.05)
            txt = str(val); color = INK
            if "%" in txt and (txt.strip().startswith("+")): color = GREEN
            elif "%" in txt and (txt.strip().startswith("−") or txt.strip().startswith("-")): color = RED
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT
            r = p.add_run(); r.text = txt
            r.font.name = FONT; r.font.size = Pt(fs); r.font.bold = (ci == 0); r.font.color.rgb = color
    return t


def _style_chart(ch, multi):
    ch.has_title = False
    ch.has_legend = multi
    if multi:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(9); ch.legend.font.name = FONT; ch.legend.font.color.rgb = GRAY
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = PLATE
    va.major_gridlines.format.line.width = Pt(0.75)
    va.minimum_scale = 0
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(9); va.tick_labels.font.name = FONT; va.tick_labels.font.color.rgb = GRAY
    ca = ch.category_axis
    ca.has_major_gridlines = False
    ca.format.line.color.rgb = LINE
    ca.tick_labels.font.size = Pt(9); ca.tick_labels.font.name = FONT; ca.tick_labels.font.color.rgb = GRAY


def _bar(slide, x, y, w, h, cats, series, *, labels=False):
    cd = CategoryChartData(); cd.categories = cats
    for s in series:
        cd.add_series(s["name"], s["data"])
    ch = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, w, h, cd).chart
    _style_chart(ch, len(series) > 1)
    ch.plots[0].gap_width = 60
    cols = [ACCENT, NAVY, ORANGE, GREEN]
    for i, ser in enumerate(ch.series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = GRAYBAR if series[i].get("plan") else cols[i % 4]
        ser.format.line.fill.background()
    if labels:
        pl = ch.plots[0]; pl.has_data_labels = True
        pl.data_labels.number_format = '0.0'; pl.data_labels.number_format_is_linked = False
        pl.data_labels.font.size = Pt(9); pl.data_labels.font.name = FONT; pl.data_labels.font.bold = True
        pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    return ch


def _line(slide, x, y, w, h, cats, series):
    cd = CategoryChartData(); cd.categories = cats
    for s in series:
        cd.add_series(s["name"], s["data"])
    ch = slide.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, w, h, cd).chart
    _style_chart(ch, True)
    cols = [ACCENT, NAVY, ORANGE, GREEN]
    for i, ser in enumerate(ch.series):
        col = GRAYBAR if series[i].get("plan") else cols[i % 4]
        ser.format.line.color.rgb = col
        ser.format.line.width = Pt(2.25 if series[i].get("plan") else 3)
        if series[i].get("forecast"):
            ser.format.line.dash_style = 7
        m = ser.marker; m.style = 8; m.size = 6
        m.format.fill.solid(); m.format.fill.fore_color.rgb = col
        m.format.line.color.rgb = col
    return ch


def _content_slide(prs, title, *, sub=None):
    s = prs.slides.add_slide(_layout(prs, L_BODY))
    _strip_placeholders(s, keep_idx=(0,))   # оставляем только заголовок
    _fill_title(s, title)
    if sub:
        _txt(s, MX, Cm(4.5), CONTENT_W, Cm(0.8), sub, size=12, color=GRAY)
    return s


def _divider(prs, num, title, sub):
    s = prs.slides.add_slide(_layout(prs, L_DIV))
    _strip_placeholders(s, keep_idx=())      # свой дизайн
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background(); bg.shadow.inherit = False
    _send_back(s, bg)
    _txt(s, SLIDE_W - Cm(11), Cm(0.5), Cm(10), Cm(8), num,
         size=170, color=RGBColor(0x24, 0x44, 0x86), bold=True, align=PP_ALIGN.RIGHT)
    _txt(s, MX, Cm(6.2), CONTENT_W, Cm(1), "РАЗДЕЛ", size=12, color=RGBColor(0x8F, 0xC4, 0xFF), bold=True)
    _txt(s, MX, Cm(7.2), Cm(24), Cm(4.5), title, size=54, color=WHITE, bold=True)
    _txt(s, MX, Cm(13.2), Cm(24), Cm(3), sub, size=14, color=RGBColor(0xC9, 0xDB, 0xF5))
    return s


def build(template_path, out_path):
    prs = Presentation(template_path)
    _del_slides(prs)
    months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
    adm = ["я25","ф","м","а","м","и","и","а","с","о","н","д","я26","ф","м","а","м","и"]

    # ── ТИТУЛ (чистая кастомная раскладка, без плейсхолдеров) ──
    s = prs.slides.add_slide(_layout(prs, L_TITLE))
    _strip_placeholders(s, keep_idx=())       # убрать все подсказки и картинку-заглушку
    # акцентная плашка-логотип
    dot = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, MX, Cm(2.4), Cm(1.0), Cm(1.0))
    dot.adjustments[0] = 0.3
    dot.fill.solid(); dot.fill.fore_color.rgb = ACCENT
    dot.line.fill.background(); dot.shadow.inherit = False
    _txt(s, MX + Cm(1.3), Cm(2.45), Cm(8), Cm(1.0), "Контур", size=20, bold=True, color=INK)
    _txt(s, MX, Cm(4.6), CONTENT_W, Cm(0.8),
         "КВАРТАЛЬНЫЙ ОТЧЁТ ПО ПРОДАЖАМ", size=12, color=ACCENT, bold=True)
    _txt(s, MX, Cm(5.5), CONTENT_W, Cm(3.5), "Итоги Q1 2026", size=60, bold=True, color=INK)
    _txt(s, MX, Cm(9.6), CONTENT_W, Cm(1.6),
         "Розница · Общепит · Кассы · ОФД · Маркет · Контекстная реклама · прогноз 2026",
         size=14, color=GRAY)
    metas = [("Период", "Q1 2026"), ("Сравнение", "Q4 2025 · Q1 2025"),
             ("Выручка факт", "68,2 млн ₽"), ("Выполнение плана", "92%")]
    for i, (k, v) in enumerate(metas):
        x = MX + Cm(i * 7.6)
        _txt(s, x, Cm(13.0), Cm(7.2), Cm(0.6), k, size=9, color=GRAY, bold=True)
        _txt(s, x, Cm(13.7), Cm(7.2), Cm(1.2), v, size=18, color=INK, bold=True)
    _foot(s, "Контур · продажи", "сформировано 13.06.2026")

    # ── ОБЗОР ──────────────────────────────────────────────────
    s = _content_slide(prs, "Ключевые метрики квартала")
    _cards4(s, Cm(5.6), [
        ("Выручка факт", "68,2 млн", "92% от плана 74,6 млн", RED),
        ("Кол-во оплат", "6 859", "план 6 773 · 101%", GRAY),
        ("Средний чек", "9,9 тыс", "факт / кол-во оплат", GRAY),
        ("Недобор к плану", "−6,3 млн", "в основном март", RED),
    ])
    _txt(s, MX, Cm(10.2), Cm(15), Cm(0.6), "ПЛАН VS ФАКТ ПО МЕСЯЦАМ, МЛН ₽", size=9, color=GRAY, bold=True)
    _bar(s, MX, Cm(10.9), Cm(15), Cm(6.3), ["Январь", "Февраль", "Март"],
         [{"name": "План", "plan": True, "data": [22.33, 24.75, 27.47]},
          {"name": "Факт", "data": [21.12, 23.69, 23.42]}], labels=True)
    _txt(s, Cm(17.8), Cm(10.2), Cm(14.4), Cm(0.6), "ВЫРУЧКА ПО НАПРАВЛЕНИЯМ", size=9, color=GRAY, bold=True)
    _table(s, Cm(17.8), Cm(10.9), Cm(14.4),
           ["Направление", "Выручка", "Доля"],
           [["Розница", "~32,5 млн", "48%"], ["Кассы", "~22,2 млн", "33%"],
            ["Общепит", "~12,8 млн", "19%"], ["в т.ч. ОФД", "11,7 млн", "17%"]])
    _foot(s, "Источник: Царь свод (план-факт) + Царь продажи")

    # ── ДИНАМИКА ───────────────────────────────────────────────
    s = _content_slide(prs, "Бизнес-юнит: план и факт по месяцам")
    _txt(s, MX, Cm(5.2), Cm(16), Cm(0.6), "ПЛАН VS ФАКТ, МЛН ₽", size=9, color=GRAY, bold=True)
    _bar(s, MX, Cm(5.9), Cm(16), Cm(10.5), ["Январь", "Февраль", "Март"],
         [{"name": "План", "plan": True, "data": [22.33, 24.75, 27.47]},
          {"name": "Факт", "data": [21.12, 23.69, 23.42]}], labels=True)
    _table(s, Cm(18.5), Cm(5.9), Cm(13.7),
           ["Месяц", "План", "Факт", "%"],
           [["Январь", "22,33", "21,12", "95"], ["Февраль", "24,75", "23,69", "96"],
            ["Март", "27,47", "23,42", "85"], ["Итого", "74,56", "68,23", "92"]])
    for i, n in enumerate([
        "Январь–февраль удерживаем 95–96% плана.",
        "Март — провал до 85%: план вырос до 27,5 млн, факт остался на уровне февраля.",
        "Февраль перевыполнен по оплатам — выручку тянет вниз средний чек."]):
        _txt(s, Cm(18.5), Cm(10.4 + i * 1.7), Cm(13.7), Cm(1.6), "•  " + n, size=11, color=INK)
    _foot(s, "Источник: Царь свод все разрезы")

    # ── РАЗДЕЛ 01 ──────────────────────────────────────────────
    _divider(prs, "01", "Направления",
             "Розница · Общепит · Кассы · ОФД — выручка, динамика QoQ/YoY и ключевые тарифы.")

    def direction(name, cards, tariffs, src):
        s = _content_slide(prs, name)
        _cards4(s, Cm(5.6), cards)
        _txt(s, MX, Cm(10.2), Cm(20), Cm(0.6), "ТАРИФЫ В Q1 2026, МЛН ₽", size=9, color=GRAY, bold=True)
        _table(s, MX, Cm(10.9), CONTENT_W,
               ["Тариф", "Q1 2025", "Q4 2025", "Q1 2026", "QoQ", "YoY"], tariffs)
        _foot(s, src)

    direction("Розница",
        [("Выручка Q1 2026", "~32,5 млн", "48% БЮ", GRAY),
         ("Доля направления", "48%", "крупнейшее", GRAY),
         ("Топ-тариф", "Оптим. Розница", "6,66 млн · +9% QoQ", GREEN, 18),
         ("Под давлением", "Все госсистемы", "3,47 млн · −13% YoY", RED, 18)],
        [["Оптимальный Розница", "7,13", "6,13", "6,66", "+9%", "−7%"],
         ["Все госсистемы", "3,97", "3,80", "3,47", "−9%", "−13%"],
         ["Базовый Розница", "1,89", "1,59", "1,41", "−11%", "−25%"],
         ["Премиум Розница", "1,64", "1,13", "0,95", "−16%", "−42%"],
         ["Маркировка", "1,39", "3,26", "1,86", "−43%", "+34%"]],
        "Источник: Царь продажи · Бизнес-юнит = Госсистемы для розницы")

    direction("Общепит",
        [("Выручка Q1 2026", "~12,8 млн", "19% БЮ", GRAY),
         ("Доля направления", "19%", "третье по объёму", GRAY),
         ("Топ-тариф", "Оптим. Общепит", "1,29 млн · +44% QoQ", GREEN, 18),
         ("Динамика", "растёт", "+8% YoY по топ-тарифу", GREEN)],
        [["Оптимальный Общепит", "1,20", "0,89", "1,29", "+44%", "+8%"],
         ["КМ модификатор ЕГАИС", "1,17", "0,99", "1,09", "+10%", "−7%"],
         ["Базовый Общепит", "—", "—", "0,49", "·", "·"]],
        "Источник: Царь продажи · Бизнес-юнит = Госсистемы для общепита")

    direction("Кассы",
        [("Выручка Q1 2026", "~22,2 млн", "33% БЮ", GRAY),
         ("Доля направления", "33%", "второе по объёму", GRAY),
         ("Топ-тариф", "Терминал", "2,75 млн · +29% QoQ", GREEN),
         ("Растущее", "Перерег. ККТ", "1,56 млн · +48% YoY", GREEN, 18)],
        [["Терминал", "2,63", "2,14", "2,75", "+29%", "+5%"],
         ["Модификатор Маркировка", "2,46", "2,30", "2,55", "+11%", "+4%"],
         ["Лицензия Атол ИТС", "1,23", "1,39", "1,59", "+14%", "+29%"],
         ["Перерегистрация ККТ", "1,05", "1,30", "1,56", "+20%", "+48%"],
         ["Атол ИТС", "1,48", "1,00", "1,24", "+24%", "−17%"],
         ["Сканер", "1,21", "1,37", "1,16", "−15%", "−5%"]],
        "Источник: Царь продажи · Бизнес-юнит = Кассовики")

    # ОФД
    s = _content_slide(prs, "ОФД")
    _cards4(s, Cm(5.6), [
        ("Выручка Q1 2026", "11,69 млн", "+4,2% к Q4 2025", GREEN),
        ("YoY", "−4,9%", "было 12,28 млн", RED),
        ("Кол-во оплат", "2 122", "+49 к Q4", GREEN),
        ("Средний чек", "5,5 тыс", "11,69 / 2 122", GRAY)])
    _txt(s, MX, Cm(10.2), Cm(15), Cm(0.6), "ВЫРУЧКА ПО МЕСЯЦАМ Q1, МЛН ₽", size=9, color=GRAY, bold=True)
    _bar(s, MX, Cm(10.9), Cm(15), Cm(6.3), ["Январь", "Февраль", "Март"],
         [{"name": "Выручка", "data": [3.94, 3.89, 3.85]}], labels=True)
    _txt(s, Cm(17.8), Cm(10.2), Cm(14.4), Cm(0.6), "ТОП ТАРИФОВ, МЛН ₽", size=9, color=GRAY, bold=True)
    _table(s, Cm(17.8), Cm(10.9), Cm(14.4), ["Тариф", "Q1 2026"],
           [["ОФД-36", "3,93"], ["ОФД-15", "3,73"], ["ОФД-13", "3,38"],
            ["ОФД-24", "0,46"], ["Прочие ОФД", "0,19"]])
    _foot(s, "Источник: Царь продажи · Тариф содержит «ОФД»")

    # ТОП-20
    s = _content_slide(prs, "Топ-20 тарифов: QoQ и YoY")
    _table(s, MX, Cm(5.2), CONTENT_W,
           ["Тариф", "Q1 2025", "Q4 2025", "Q1 2026", "QoQ", "YoY"],
           [["Прочие","26,63","21,26","14,96","−30%","−44%"],
            ["Оптимальный Розница","7,13","6,13","6,66","+9%","−7%"],
            ["ОФД-36","4,77","4,79","3,93","−18%","−18%"],
            ["ОФД-15","4,17","3,00","3,73","+24%","−11%"],
            ["Все госсистемы","3,97","3,80","3,47","−9%","−13%"],
            ["ОФД-13","2,72","2,73","3,38","+24%","+24%"],
            ["Настройка Контур.Маркет","1,59","1,45","2,93","+102%","+85%"],
            ["Терминал","2,63","2,14","2,75","+29%","+5%"],
            ["Модификатор Маркировка","2,46","2,30","2,55","+11%","+4%"],
            ["Маркировка","1,39","3,26","1,86","−43%","+34%"],
            ["Лицензия Атол ИТС","1,23","1,39","1,59","+14%","+29%"],
            ["Перерегистрация ККТ","1,05","1,30","1,56","+20%","+48%"],
            ["Базовый Розница","1,89","1,59","1,41","−11%","−25%"],
            ["Оптимальный Общепит","1,20","0,89","1,29","+44%","+8%"],
            ["Атол ИТС","1,48","1,00","1,24","+24%","−17%"],
            ["Сертификат","0,86","0,66","1,23","+88%","+43%"],
            ["Сканер","1,21","1,37","1,16","−15%","−5%"],
            ["КМ модификатор ЕГАИС","1,17","0,99","1,09","+10%","−7%"],
            ["Премиум Розница","1,64","1,13","0,95","−16%","−42%"],
            ["Базовый Услуги","1,04","0,67","0,86","+29%","−18%"]], fs=9, row_h=Cm(0.56))
    _foot(s, "Источник: Царь продажи · группировка по тарифу, млн ₽")

    # ── РАЗДЕЛ 02 ──────────────────────────────────────────────
    _divider(prs, "02", "Прогноз 2026 — Маркет и ОФД",
             "Ансамбль моделей: plan-pacing · сезонный YoY · ETS · регрессия. "
             "Факт по май, прогноз июнь–декабрь, сценарии низ/база/верх.")

    def fc_slide(title, cards, plan, fact, fc, src):
        s = _content_slide(prs, title)
        _cards4(s, Cm(5.6), cards)
        _txt(s, MX, Cm(10.2), Cm(20), Cm(0.6), "ПЛАН / ФАКТ / ПРОГНОЗ, МЛН ₽", size=9, color=GRAY, bold=True)
        _line(s, MX, Cm(10.9), CONTENT_W, Cm(6.3), months,
              [{"name": "План", "plan": True, "data": plan},
               {"name": "Факт", "data": fact},
               {"name": "Прогноз", "forecast": True, "data": fc}])
        _foot(s, src)

    fc_slide("Маркет — выполнение плана 2026",
        [("Годовой план", "119,0 млн", "Продукт = Маркет", GRAY),
         ("Факт за 5 мес", "44,5 млн", "91% от плана периода", RED),
         ("YoY к 2025", "−21%", "падение спроса", RED),
         ("Прогноз года", "86%", "102 млн · 78–91%", RED)],
        [9.78,9.46,10.98,10.45,8.36,10.32,10.13,9.01,8.79,9.95,9.71,12.10],
        [8.24,9.20,10.26,8.99,7.78,None,None,None,None,None,None,None],
        [None,None,None,None,7.78,9.1,8.4,7.2,7.6,7.8,7.4,9.9],
        "Источник: план-факт «Продукт = Маркет» · ансамбль моделей")

    fc_slide("ОФД — выполнение плана 2026",
        [("Годовой план", "37,5 млн", "проект п453", GRAY),
         ("Факт за 5 мес", "18,9 млн", "123% от плана периода", GREEN),
         ("YoY к 2025", "−16%", "план консервативен", GRAY),
         ("Прогноз года", "115%", "43,3 млн · 103–124%", GREEN)],
        [2.81,3.11,3.18,3.24,2.97,3.04,3.03,2.99,2.90,3.28,3.20,3.79],
        [4.05,4.12,4.18,3.71,2.84,None,None,None,None,None,None,None],
        [None,None,None,None,2.84,3.6,3.4,3.1,3.1,3.4,3.2,4.5],
        "Источник: план-факт проект п453 · ансамбль моделей")

    # ── РАЗДЕЛ 03 ──────────────────────────────────────────────
    _divider(prs, "03", "Контекстная реклама",
             "Расход, CPL, CPO и привлечённые оплаты по Маркету и ОФД. Дашборды кампаний, янв 2025 – май 2026.")

    def ad_slide(title, cards, spend, cpl, src):
        s = _content_slide(prs, title)
        _cards4(s, Cm(5.6), cards)
        _txt(s, MX, Cm(10.2), Cm(15), Cm(0.6), "РАСХОД ПО МЕСЯЦАМ, МЛН ₽", size=9, color=GRAY, bold=True)
        _bar(s, MX, Cm(10.9), Cm(15), Cm(6.3), adm, [{"name": "Расход", "data": spend}])
        _txt(s, Cm(17.8), Cm(10.2), Cm(14.4), Cm(0.6), "CPL ПО МЕСЯЦАМ, ТЫС ₽", size=9, color=GRAY, bold=True)
        _line(s, Cm(17.8), Cm(10.9), Cm(14.4), Cm(6.3), adm, [{"name": "CPL", "data": cpl}])
        _foot(s, src)

    ad_slide("Контекст · Маркет",
        [("Расход всего", "14,7 млн", "янв 2025 – май 2026", GRAY),
         ("Оплаты", "315", "привлечено", GRAY),
         ("CPO", "46,8 тыс", "высокая стоимость", RED),
         ("CPL средний", "10,0 тыс", "снижается к 2026", GREEN)],
        [1.43,1.46,1.29,1.65,1.14,0.89,0.68,0.47,0.62,0.83,0.75,0.64,0.57,0.56,0.70,0.45,0.47,0.13],
        [16.2,14.3,14.7,13.0,9.7,8.5,6.6,8.4,8.2,11.0,11.2,9.7,7.7,9.7,10.5,6.0,4.5,4.6],
        "Источник: дашборд Контекст-Маркет · суммы сверены с «Итого»")

    ad_slide("Контекст · ОФД",
        [("Расход всего", "7,1 млн", "янв 2025 – май 2026", GRAY),
         ("Оплаты", "397", "привлечено", GRAY),
         ("CPO", "17,9 тыс", "в 2,6× дешевле Маркета", GREEN),
         ("CPL средний", "4,5 тыс", "эффективнее", GREEN)],
        [0.35,0.42,0.46,0.59,0.52,0.43,0.36,0.37,0.41,0.46,0.43,0.39,0.42,0.39,0.42,0.39,0.26,0.05],
        [2.6,4.1,5.0,8.6,10.1,9.3,4.7,4.7,4.6,5.2,5.0,3.9,3.5,3.5,3.6,3.2,4.0,1.9],
        "Источник: дашборд Контекст-ОФД · суммы сверены с «Итого»")

    # ── ФИНАЛ ──────────────────────────────────────────────────
    s = prs.slides.add_slide(_layout(prs, L_BODY))
    _strip_placeholders(s, keep_idx=(0,))
    _fill_title(s, "Итоги Q1 2026: ключевые выводы")
    for i, b in enumerate([
        "Q1: 92% плана. Январь–февраль 95–96%, март провален (85%).",
        "Маркет: −21% YoY, прогноз 2026 ≈ 86% годового плана — недовыполнение.",
        "ОФД: факт YTD 123% плана периода, прогноз 2026 ≈ 115% — перевыполнение.",
        "Реклама: ОФД-контекст в 2,6× эффективнее Маркета по CPO (17,9 vs 46,8 тыс ₽).",
    ]):
        _txt(s, MX, Cm(6.0 + i * 2.2), CONTENT_W, Cm(2.0), "•  " + b, size=16, color=INK)
    _foot(s, "Контур · сводка и прогноз собраны автоматически")

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)


if __name__ == "__main__":
    import sys
    t = sys.argv[1] if len(sys.argv) > 1 else "Шаблон презентации Контур Blue 2023_16x9_Montserrat (2).pptx"
    o = sys.argv[2] if len(sys.argv) > 2 else "Итоги Q1 2026 — Контур.pptx"
    build(t, o)
    print(f"OK → {o}")
