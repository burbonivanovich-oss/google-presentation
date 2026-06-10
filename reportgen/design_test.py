"""Design-fidelity test deck.

Создаёт презентацию с нуля и наполняет её эталонными элементами, чтобы
посмотреть, насколько точно Google Slides API воспроизводит то, что мы
просим: типографику, списки, таблицы, фигуры, картинки, графики из
Sheets и поведение replaceAllText.

Каждый слайд в начале содержит блок «Что просили» — для визуального
сравнения с тем, что получилось.
"""
from __future__ import annotations

from googleapiclient.discovery import Resource

BLACK = {"red": 0.0, "green": 0.0, "blue": 0.0}
WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
GRAY = {"red": 0.37, "green": 0.40, "blue": 0.44}
LIGHT_GRAY = {"red": 0.94, "green": 0.94, "blue": 0.94}
BLUE = {"red": 0.10, "green": 0.40, "blue": 0.85}
RED = {"red": 0.85, "green": 0.18, "blue": 0.18}
GREEN = {"red": 0.20, "green": 0.55, "blue": 0.28}
ORANGE = {"red": 0.95, "green": 0.55, "blue": 0.05}

# Стандартный размер слайда 10" × 5.625" = 720 × 405 pt
SLIDE_W = 720
SLIDE_H = 405


def pt(v: float) -> dict:
    return {"magnitude": v, "unit": "PT"}


def color_field(rgb: dict) -> dict:
    return {"opaqueColor": {"rgbColor": rgb}}


def transform(x: float, y: float) -> dict:
    return {"scaleX": 1, "scaleY": 1, "translateX": x, "translateY": y, "unit": "PT"}


def element_props(slide_id: str, x: float, y: float, w: float, h: float) -> dict:
    return {
        "pageObjectId": slide_id,
        "size": {"width": pt(w), "height": pt(h)},
        "transform": transform(x, y),
    }


def build_design_test_deck(slides_svc: Resource, sheets_svc: Resource | None, drive_svc: Resource, parent_folder_id: str | None = None) -> str:
    """Создаёт презентацию-стенд. Возвращает её ID."""
    pres = slides_svc.presentations().create(body={"title": "Design fidelity test"}).execute()
    pres_id = pres["presentationId"]
    default_slide_id = pres["slides"][0]["objectId"]

    if parent_folder_id:
        file = drive_svc.files().get(fileId=pres_id, fields="parents", supportsAllDrives=True).execute()
        prev_parents = ",".join(file.get("parents", []))
        drive_svc.files().update(
            fileId=pres_id,
            addParents=parent_folder_id,
            removeParents=prev_parents,
            supportsAllDrives=True,
        ).execute()

    _run(slides_svc, pres_id, _slide_typography())
    _run(slides_svc, pres_id, _slide_lists())
    _run(slides_svc, pres_id, _slide_tables())
    _run(slides_svc, pres_id, _slide_shapes())
    _run(slides_svc, pres_id, _slide_image())
    _run(slides_svc, pres_id, _slide_replace_setup())
    _run(slides_svc, pres_id, _slide_replace_apply())

    if sheets_svc is not None:
        try:
            ssid, chart_id = _build_chart_source(sheets_svc, parent_folder_id, drive_svc)
            _run(slides_svc, pres_id, _slide_chart(ssid, chart_id))
        except Exception as e:  # noqa: BLE001
            _run(slides_svc, pres_id, _slide_chart_error(str(e)))

    _run(slides_svc, pres_id, [{"deleteObject": {"objectId": default_slide_id}}])
    return pres_id


def _run(slides_svc: Resource, pres_id: str, requests: list[dict]) -> dict:
    if not requests:
        return {}
    return slides_svc.presentations().batchUpdate(
        presentationId=pres_id, body={"requests": requests}
    ).execute()


def _title_and_brief(slide_id: str, title: str, brief: str, suffix: str) -> list[dict]:
    """Стандартная шапка слайда: жирный заголовок + серая подсказка."""
    title_id = f"title_{suffix}"
    brief_id = f"brief_{suffix}"
    return [
        {"createShape": {"objectId": title_id, "shapeType": "TEXT_BOX",
                         "elementProperties": element_props(slide_id, 30, 20, 660, 30)}},
        {"insertText": {"objectId": title_id, "text": title}},
        {"updateTextStyle": {"objectId": title_id, "textRange": {"type": "ALL"},
                             "style": {"fontSize": pt(22), "bold": True,
                                       "foregroundColor": color_field(BLACK)},
                             "fields": "fontSize,bold,foregroundColor"}},
        {"createShape": {"objectId": brief_id, "shapeType": "TEXT_BOX",
                         "elementProperties": element_props(slide_id, 30, 52, 660, 24)}},
        {"insertText": {"objectId": brief_id, "text": "Что просили: " + brief}},
        {"updateTextStyle": {"objectId": brief_id, "textRange": {"type": "ALL"},
                             "style": {"fontSize": pt(10), "italic": True,
                                       "foregroundColor": color_field(GRAY)},
                             "fields": "fontSize,italic,foregroundColor"}},
    ]


def _slide_typography() -> list[dict]:
    slide_id = "slide_typo"
    reqs: list[dict] = [
        {"createSlide": {"objectId": slide_id,
                         "slideLayoutReference": {"predefinedLayout": "BLANK"}}}
    ]
    reqs += _title_and_brief(
        slide_id, "1. Типографика",
        "28pt bold чёрный, 18pt italic серый, 14pt regular, 10pt caption; "
        "выравнивания LEFT/CENTER/RIGHT/JUSTIFY; line-height 1.0/1.5/2.0",
        "typo",
    )
    blocks = [
        ("typo_h", "Heading 28pt Bold", 30, 90, 660, 40,
         {"fontSize": pt(28), "bold": True, "foregroundColor": color_field(BLACK)},
         "fontSize,bold,foregroundColor", None),
        ("typo_sub", "Subheading 18pt Italic Gray", 30, 130, 660, 32,
         {"fontSize": pt(18), "italic": True, "foregroundColor": color_field(GRAY)},
         "fontSize,italic,foregroundColor", None),
        ("typo_body", "Body 14pt regular — обычный текст для проверки кириллицы, "
                     "пунктуации «ёлочек» и тире —", 30, 165, 660, 26,
         {"fontSize": pt(14), "foregroundColor": color_field(BLACK)},
         "fontSize,foregroundColor", None),
        ("typo_cap", "Caption 10pt regular gray — мелкий служебный текст", 30, 195, 660, 18,
         {"fontSize": pt(10), "foregroundColor": color_field(GRAY)},
         "fontSize,foregroundColor", None),
        ("typo_left", "LEFT: lorem ipsum dolor sit amet — выравнивание по левому краю",
         30, 230, 660, 22, {"fontSize": pt(11)}, "fontSize", "START"),
        ("typo_center", "CENTER: lorem ipsum dolor sit amet — выравнивание по центру",
         30, 252, 660, 22, {"fontSize": pt(11)}, "fontSize", "CENTER"),
        ("typo_right", "RIGHT: lorem ipsum dolor sit amet — выравнивание по правому краю",
         30, 274, 660, 22, {"fontSize": pt(11)}, "fontSize", "END"),
        ("typo_just", "JUSTIFIED: " + ("lorem ipsum dolor sit amet consectetur " * 4),
         30, 296, 660, 50, {"fontSize": pt(11)}, "fontSize", "JUSTIFIED"),
    ]
    for oid, text, x, y, w, h, style, fields, align in blocks:
        reqs += [
            {"createShape": {"objectId": oid, "shapeType": "TEXT_BOX",
                             "elementProperties": element_props(slide_id, x, y, w, h)}},
            {"insertText": {"objectId": oid, "text": text}},
            {"updateTextStyle": {"objectId": oid, "textRange": {"type": "ALL"},
                                 "style": style, "fields": fields}},
        ]
        if align:
            reqs.append({"updateParagraphStyle": {
                "objectId": oid, "textRange": {"type": "ALL"},
                "style": {"alignment": align}, "fields": "alignment"
            }})

    # line-height тест
    lh_blocks = [
        ("typo_lh10", "line-height 1.0\nвторая строка\nтретья", 30, 350, 200, 50, 1.0),
        ("typo_lh15", "line-height 1.5\nвторая строка\nтретья", 260, 350, 200, 50, 1.5),
        ("typo_lh20", "line-height 2.0\nвторая строка\nтретья", 490, 350, 200, 50, 2.0),
    ]
    for oid, text, x, y, w, h, lh in lh_blocks:
        reqs += [
            {"createShape": {"objectId": oid, "shapeType": "TEXT_BOX",
                             "elementProperties": element_props(slide_id, x, y, w, h)}},
            {"insertText": {"objectId": oid, "text": text}},
            {"updateTextStyle": {"objectId": oid, "textRange": {"type": "ALL"},
                                 "style": {"fontSize": pt(10)}, "fields": "fontSize"}},
            {"updateParagraphStyle": {"objectId": oid, "textRange": {"type": "ALL"},
                                      "style": {"lineSpacing": lh * 100},
                                      "fields": "lineSpacing"}},
        ]
    return reqs


def _slide_lists() -> list[dict]:
    slide_id = "slide_lists"
    reqs: list[dict] = [
        {"createSlide": {"objectId": slide_id,
                         "slideLayoutReference": {"predefinedLayout": "BLANK"}}}
    ]
    reqs += _title_and_brief(
        slide_id, "2. Списки",
        "bulleted • numbered 1. 2. 3. • nested 3 уровня с разными маркерами",
        "lists",
    )

    bulleted_text = "Первый пункт\nВторой пункт\nТретий пункт\nЧетвёртый пункт"
    numbered_text = "Шаг 1\nШаг 2\nШаг 3\nШаг 4"
    # Для вложенного списка \t в начале строки = уровень вложенности
    nested_text = (
        "Уровень 1 — рынок\n"
        "\tУровень 2 — канал\n"
        "\t\tУровень 3 — кампания\n"
        "\t\tУровень 3 — другая кампания\n"
        "\tУровень 2 — другой канал\n"
        "Уровень 1 — другой рынок"
    )

    cols = [
        ("list_bullet", "Bulleted", bulleted_text, 30, 90, 220, 200,
         "BULLET_DISC_CIRCLE_SQUARE"),
        ("list_num", "Numbered", numbered_text, 260, 90, 200, 200,
         "NUMBERED_DIGIT_ALPHA_ROMAN"),
        ("list_nest", "Nested 3 уровня", nested_text, 470, 90, 220, 220,
         "BULLET_DISC_CIRCLE_SQUARE"),
    ]
    for oid, title, text, x, y, w, h, preset in cols:
        title_id = oid + "_t"
        reqs += [
            {"createShape": {"objectId": title_id, "shapeType": "TEXT_BOX",
                             "elementProperties": element_props(slide_id, x, y, w, 20)}},
            {"insertText": {"objectId": title_id, "text": title}},
            {"updateTextStyle": {"objectId": title_id, "textRange": {"type": "ALL"},
                                 "style": {"fontSize": pt(11), "bold": True,
                                           "foregroundColor": color_field(GRAY)},
                                 "fields": "fontSize,bold,foregroundColor"}},
            {"createShape": {"objectId": oid, "shapeType": "TEXT_BOX",
                             "elementProperties": element_props(slide_id, x, y + 24, w, h)}},
            {"insertText": {"objectId": oid, "text": text}},
            {"updateTextStyle": {"objectId": oid, "textRange": {"type": "ALL"},
                                 "style": {"fontSize": pt(12)}, "fields": "fontSize"}},
            {"createParagraphBullets": {"objectId": oid, "textRange": {"type": "ALL"},
                                        "bulletPreset": preset}},
        ]
    return reqs


def _slide_tables() -> list[dict]:
    slide_id = "slide_table"
    table_id = "tbl_main"
    rows, cols = 5, 5
    reqs: list[dict] = [
        {"createSlide": {"objectId": slide_id,
                         "slideLayoutReference": {"predefinedLayout": "BLANK"}}},
    ]
    reqs += _title_and_brief(
        slide_id, "3. Таблицы",
        "5×5; header — синий фон, белый bold; колонка 1 ширина 200pt, остальные 100pt; "
        "ячейка [2,2] жёлтая; merge верхней строки [0,1..4]; зебра строк 2/4",
        "tbl",
    )
    reqs.append({"createTable": {
        "objectId": table_id,
        "elementProperties": element_props(slide_id, 30, 95, 600, 260),
        "rows": rows, "columns": cols,
    }})
    # column widths
    reqs.append({"updateTableColumnProperties": {
        "objectId": table_id, "columnIndices": [0],
        "tableColumnProperties": {"columnWidth": pt(200)},
        "fields": "columnWidth"}})
    for i in range(1, cols):
        reqs.append({"updateTableColumnProperties": {
            "objectId": table_id, "columnIndices": [i],
            "tableColumnProperties": {"columnWidth": pt(100)},
            "fields": "columnWidth"}})

    # header: текст + merge + заливка + стиль
    headers = ["Канал", "Q4-2025", "Q1-2026", "Δ %", "Комментарий"]
    for c, h in enumerate(headers):
        reqs.append({"insertText": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 0, "columnIndex": c},
            "text": h}})
    reqs.append({"updateTableCellProperties": {
        "objectId": table_id,
        "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0},
                       "rowSpan": 1, "columnSpan": cols},
        "tableCellProperties": {
            "tableCellBackgroundFill": {"solidFill": {"color": color_field(BLUE)}}},
        "fields": "tableCellBackgroundFill.solidFill.color"}})
    for c in range(cols):
        reqs.append({"updateTextStyle": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 0, "columnIndex": c},
            "textRange": {"type": "ALL"},
            "style": {"bold": True, "foregroundColor": color_field(WHITE),
                      "fontSize": pt(11)},
            "fields": "bold,foregroundColor,fontSize"}})

    # data rows
    data = [
        ["Google Ads", "1 200 000", "1 350 000", "+12.5", "↑ ROAS стабилен"],
        ["Yandex Direct", "800 000", "740 000", "−7.5", "↓ снижение трафика"],
        ["VK Ads", "450 000", "520 000", "+15.6", "↑ новый креатив"],
        ["Telegram Ads", "200 000", "180 000", "−10.0", "↓ слабый CTR"],
    ]
    for r, row in enumerate(data, start=1):
        for c, val in enumerate(row):
            reqs.append({"insertText": {
                "objectId": table_id,
                "cellLocation": {"rowIndex": r, "columnIndex": c},
                "text": val}})
            reqs.append({"updateTextStyle": {
                "objectId": table_id,
                "cellLocation": {"rowIndex": r, "columnIndex": c},
                "textRange": {"type": "ALL"},
                "style": {"fontSize": pt(10)}, "fields": "fontSize"}})

    # зебра: строки 2 и 4 — серый фон
    for r in (2, 4):
        reqs.append({"updateTableCellProperties": {
            "objectId": table_id,
            "tableRange": {"location": {"rowIndex": r, "columnIndex": 0},
                           "rowSpan": 1, "columnSpan": cols},
            "tableCellProperties": {
                "tableCellBackgroundFill": {"solidFill": {"color": color_field(LIGHT_GRAY)}}},
            "fields": "tableCellBackgroundFill.solidFill.color"}})

    # одиночная ячейка [2,2] — оранжевая (поверх зебры)
    reqs.append({"updateTableCellProperties": {
        "objectId": table_id,
        "tableRange": {"location": {"rowIndex": 2, "columnIndex": 2},
                       "rowSpan": 1, "columnSpan": 1},
        "tableCellProperties": {
            "tableCellBackgroundFill": {"solidFill": {"color": color_field(ORANGE)}}},
        "fields": "tableCellBackgroundFill.solidFill.color"}})

    # цвет дельт: + зелёным, − красным
    for r, val in enumerate(data, start=1):
        delta = val[3]
        col = GREEN if delta.startswith("+") else RED
        reqs.append({"updateTextStyle": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": r, "columnIndex": 3},
            "textRange": {"type": "ALL"},
            "style": {"bold": True, "foregroundColor": color_field(col)},
            "fields": "bold,foregroundColor"}})
    return reqs


def _slide_shapes() -> list[dict]:
    slide_id = "slide_shapes"
    reqs: list[dict] = [
        {"createSlide": {"objectId": slide_id,
                         "slideLayoutReference": {"predefinedLayout": "BLANK"}}},
    ]
    reqs += _title_and_brief(
        slide_id, "4. Фигуры",
        "rectangle (синяя заливка) • ellipse (без заливки, красный stroke 3pt) • "
        "line с dash DASH • стрелка RIGHT_ARROW • rounded rectangle с тенью",
        "shp",
    )
    # 1. синий прямоугольник
    reqs += [
        {"createShape": {"objectId": "shp_rect", "shapeType": "RECTANGLE",
                         "elementProperties": element_props(slide_id, 30, 100, 150, 100)}},
        {"updateShapeProperties": {
            "objectId": "shp_rect",
            "shapeProperties": {
                "shapeBackgroundFill": {"solidFill": {"color": color_field(BLUE)}},
                "outline": {"propertyState": "NOT_RENDERED"}},
            "fields": "shapeBackgroundFill.solidFill.color,outline.propertyState"}},
    ]
    # 2. красный овал без заливки
    reqs += [
        {"createShape": {"objectId": "shp_ellipse", "shapeType": "ELLIPSE",
                         "elementProperties": element_props(slide_id, 200, 100, 150, 100)}},
        {"updateShapeProperties": {
            "objectId": "shp_ellipse",
            "shapeProperties": {
                "shapeBackgroundFill": {"solidFill": {"alpha": 0.0,
                                                       "color": color_field(WHITE)}},
                "outline": {"outlineFill": {"solidFill": {"color": color_field(RED)}},
                            "weight": pt(3), "dashStyle": "SOLID"}},
            "fields": "shapeBackgroundFill,outline"}},
    ]
    # 3. line dashed
    reqs += [
        {"createLine": {"objectId": "shp_line", "lineCategory": "STRAIGHT",
                         "elementProperties": element_props(slide_id, 370, 100, 300, 1)}},
        {"updateLineProperties": {
            "objectId": "shp_line",
            "lineProperties": {
                "lineFill": {"solidFill": {"color": color_field(GRAY)}},
                "weight": pt(2), "dashStyle": "DASH"},
            "fields": "lineFill,weight,dashStyle"}},
    ]
    # 4. правая стрелка
    reqs += [
        {"createShape": {"objectId": "shp_arrow", "shapeType": "RIGHT_ARROW",
                         "elementProperties": element_props(slide_id, 370, 130, 300, 70)}},
        {"updateShapeProperties": {
            "objectId": "shp_arrow",
            "shapeProperties": {
                "shapeBackgroundFill": {"solidFill": {"color": color_field(GREEN)}},
                "outline": {"propertyState": "NOT_RENDERED"}},
            "fields": "shapeBackgroundFill,outline"}},
    ]
    # 5. скруглённый прямоугольник + текст внутри
    reqs += [
        {"createShape": {"objectId": "shp_round", "shapeType": "ROUND_RECTANGLE",
                         "elementProperties": element_props(slide_id, 30, 220, 660, 140)}},
        {"updateShapeProperties": {
            "objectId": "shp_round",
            "shapeProperties": {
                "shapeBackgroundFill": {"solidFill": {"color": color_field(LIGHT_GRAY)}},
                "outline": {"outlineFill": {"solidFill": {"color": color_field(GRAY)}},
                            "weight": pt(1)}},
            "fields": "shapeBackgroundFill,outline"}},
        {"insertText": {"objectId": "shp_round",
                        "text": "Текст внутри ROUND_RECTANGLE с заливкой и обводкой.\n"
                                "Проверяем, что текст центруется и не вылезает."}},
        {"updateTextStyle": {"objectId": "shp_round", "textRange": {"type": "ALL"},
                             "style": {"fontSize": pt(14), "foregroundColor": color_field(BLACK)},
                             "fields": "fontSize,foregroundColor"}},
        {"updateParagraphStyle": {"objectId": "shp_round", "textRange": {"type": "ALL"},
                                  "style": {"alignment": "CENTER"},
                                  "fields": "alignment"}},
    ]
    return reqs


def _slide_image() -> list[dict]:
    slide_id = "slide_image"
    reqs: list[dict] = [
        {"createSlide": {"objectId": slide_id,
                         "slideLayoutReference": {"predefinedLayout": "BLANK"}}},
    ]
    reqs += _title_and_brief(
        slide_id, "5. Картинки",
        "три раза одна и та же картинка из picsum.photos: оригинал, "
        "масштабированная вниз, и третья растянутая под прямоугольник",
        "img",
    )
    url = "https://picsum.photos/id/1015/800/600"
    reqs += [
        {"createImage": {"objectId": "img_a", "url": url,
                         "elementProperties": element_props(slide_id, 30, 100, 200, 150)}},
        {"createImage": {"objectId": "img_b", "url": url,
                         "elementProperties": element_props(slide_id, 250, 100, 100, 75)}},
        {"createImage": {"objectId": "img_c", "url": url,
                         "elementProperties": element_props(slide_id, 370, 100, 320, 180)}},
    ]
    # подписи
    captions = [
        ("cap_a", "200×150 (4:3 — нативный)", 30, 255, 200),
        ("cap_b", "100×75 (4:3, уменьшенная)", 250, 180, 100),
        ("cap_c", "320×180 (16:9 — должна обрезаться?)", 370, 285, 320),
    ]
    for oid, text, x, y, w in captions:
        reqs += [
            {"createShape": {"objectId": oid, "shapeType": "TEXT_BOX",
                             "elementProperties": element_props(slide_id, x, y, w, 20)}},
            {"insertText": {"objectId": oid, "text": text}},
            {"updateTextStyle": {"objectId": oid, "textRange": {"type": "ALL"},
                                 "style": {"fontSize": pt(9),
                                           "foregroundColor": color_field(GRAY),
                                           "italic": True},
                                 "fields": "fontSize,foregroundColor,italic"}},
        ]
    return reqs


def _slide_replace_setup() -> list[dict]:
    slide_id = "slide_replace"
    reqs: list[dict] = [
        {"createSlide": {"objectId": slide_id,
                         "slideLayoutReference": {"predefinedLayout": "BLANK"}}},
    ]
    reqs += _title_and_brief(
        slide_id, "6. replaceAllText — сохранение стиля",
        "{{name}} оформлен красным bold, {{plan}} синим italic. После замены "
        "стили должны остаться на новых значениях.",
        "rep",
    )
    # before
    reqs += [
        {"createShape": {"objectId": "rep_before", "shapeType": "TEXT_BOX",
                         "elementProperties": element_props(slide_id, 30, 100, 660, 40)}},
        {"insertText": {"objectId": "rep_before",
                        "text": "ДО: клиент {{name}}, тариф {{plan}}"}},
        {"updateTextStyle": {"objectId": "rep_before", "textRange": {"type": "ALL"},
                             "style": {"fontSize": pt(16)}, "fields": "fontSize"}},
        # стилизуем {{name}} — позиции в строке "ДО: клиент {{name}}, тариф {{plan}}"
        # {{name}} начинается с индекса 11, длина 8
        {"updateTextStyle": {"objectId": "rep_before",
                             "textRange": {"type": "FIXED_RANGE",
                                           "startIndex": 11, "endIndex": 19},
                             "style": {"bold": True, "foregroundColor": color_field(RED)},
                             "fields": "bold,foregroundColor"}},
        # {{plan}} начинается с индекса 27, длина 8
        {"updateTextStyle": {"objectId": "rep_before",
                             "textRange": {"type": "FIXED_RANGE",
                                           "startIndex": 27, "endIndex": 35},
                             "style": {"italic": True, "foregroundColor": color_field(BLUE)},
                             "fields": "italic,foregroundColor"}},
    ]
    # вторая копия — будет ниже, оставим как "после" — для неё применим replace
    reqs += [
        {"createShape": {"objectId": "rep_after", "shapeType": "TEXT_BOX",
                         "elementProperties": element_props(slide_id, 30, 160, 660, 40)}},
        {"insertText": {"objectId": "rep_after",
                        "text": "ПОСЛЕ: клиент {{name}}, тариф {{plan}}"}},
        {"updateTextStyle": {"objectId": "rep_after", "textRange": {"type": "ALL"},
                             "style": {"fontSize": pt(16)}, "fields": "fontSize"}},
        {"updateTextStyle": {"objectId": "rep_after",
                             "textRange": {"type": "FIXED_RANGE",
                                           "startIndex": 15, "endIndex": 23},
                             "style": {"bold": True, "foregroundColor": color_field(RED)},
                             "fields": "bold,foregroundColor"}},
        {"updateTextStyle": {"objectId": "rep_after",
                             "textRange": {"type": "FIXED_RANGE",
                                           "startIndex": 31, "endIndex": 39},
                             "style": {"italic": True, "foregroundColor": color_field(BLUE)},
                             "fields": "italic,foregroundColor"}},
    ]
    # пояснение
    reqs += [
        {"createShape": {"objectId": "rep_note", "shapeType": "TEXT_BOX",
                         "elementProperties": element_props(slide_id, 30, 220, 660, 100)}},
        {"insertText": {"objectId": "rep_note",
                        "text": "В строке «ДО» — placeholder как есть.\n"
                                "В строке «ПОСЛЕ» — после batchUpdate с replaceAllText "
                                "{{name}} → «Иванов И.И.», {{plan}} → «Premium».\n"
                                "Сравните: красный bold и синий italic должны "
                                "перенестись на новые значения."}},
        {"updateTextStyle": {"objectId": "rep_note", "textRange": {"type": "ALL"},
                             "style": {"fontSize": pt(10), "foregroundColor": color_field(GRAY)},
                             "fields": "fontSize,foregroundColor"}},
    ]
    return reqs


def _slide_replace_apply() -> list[dict]:
    # Тут только сама замена — в той же презентации, но только на rep_after.
    # pageObjectIds ограничивает scope замены конкретным слайдом.
    return [
        {"replaceAllText": {
            "containsText": {"text": "{{name}}", "matchCase": True},
            "replaceText": "Иванов И.И.",
            "pageObjectIds": ["slide_replace"]}},
        {"replaceAllText": {
            "containsText": {"text": "{{plan}}", "matchCase": True},
            "replaceText": "Premium",
            "pageObjectIds": ["slide_replace"]}},
    ]


def _build_chart_source(sheets_svc: Resource, parent_folder_id: str | None, drive_svc: Resource) -> tuple[str, int]:
    """Создаёт Sheet с данными и chart внутри, возвращает (spreadsheetId, chartId)."""
    ss = sheets_svc.spreadsheets().create(
        body={"properties": {"title": "Design test data"}}
    ).execute()
    ssid = ss["spreadsheetId"]
    if parent_folder_id:
        file = drive_svc.files().get(fileId=ssid, fields="parents", supportsAllDrives=True).execute()
        drive_svc.files().update(
            fileId=ssid,
            addParents=parent_folder_id,
            removeParents=",".join(file.get("parents", [])),
            supportsAllDrives=True,
        ).execute()

    sheet_id = ss["sheets"][0]["properties"]["sheetId"]

    values = [
        ["Канал", "Q4-2025", "Q1-2026"],
        ["Google Ads", 1200000, 1350000],
        ["Yandex Direct", 800000, 740000],
        ["VK Ads", 450000, 520000],
        ["Telegram Ads", 200000, 180000],
    ]
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=ssid, range="A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    resp = sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=ssid,
        body={"requests": [{
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Расходы по каналам",
                        "basicChart": {
                            "chartType": "COLUMN",
                            "legendPosition": "BOTTOM_LEGEND",
                            "axis": [
                                {"position": "BOTTOM_AXIS", "title": "Канал"},
                                {"position": "LEFT_AXIS", "title": "Спенд, ₽"},
                            ],
                            "domains": [{
                                "domain": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0, "endRowIndex": 5,
                                    "startColumnIndex": 0, "endColumnIndex": 1,
                                }]}}
                            }],
                            "series": [
                                {"series": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0, "endRowIndex": 5,
                                    "startColumnIndex": 1, "endColumnIndex": 2,
                                }]}}, "targetAxis": "LEFT_AXIS"},
                                {"series": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0, "endRowIndex": 5,
                                    "startColumnIndex": 2, "endColumnIndex": 3,
                                }]}}, "targetAxis": "LEFT_AXIS"},
                            ],
                            "headerCount": 1,
                        },
                    },
                    "position": {"newSheet": True},
                }
            }
        }]},
    ).execute()
    chart_id = resp["replies"][0]["addChart"]["chart"]["chartId"]
    return ssid, chart_id


def _slide_chart(ssid: str, chart_id: int) -> list[dict]:
    slide_id = "slide_chart"
    reqs: list[dict] = [
        {"createSlide": {"objectId": slide_id,
                         "slideLayoutReference": {"predefinedLayout": "BLANK"}}},
    ]
    reqs += _title_and_brief(
        slide_id, "7. График из Sheets (LINKED)",
        "column chart 2 серии × 4 категории; вставлен в прямоугольник 600×280pt; "
        "режим LINKED — Slides рендерит превью, обновляется по кнопке",
        "chart",
    )
    reqs.append({"createSheetsChart": {
        "objectId": "chart_main",
        "spreadsheetId": ssid,
        "chartId": chart_id,
        "linkingMode": "LINKED",
        "elementProperties": element_props(slide_id, 60, 90, 600, 280),
    }})
    return reqs


def _slide_chart_error(msg: str) -> list[dict]:
    slide_id = "slide_chart_err"
    reqs: list[dict] = [
        {"createSlide": {"objectId": slide_id,
                         "slideLayoutReference": {"predefinedLayout": "BLANK"}}},
    ]
    reqs += _title_and_brief(
        slide_id, "7. График из Sheets — ОШИБКА",
        "не удалось собрать chart-источник, см. текст ниже",
        "cherr",
    )
    reqs += [
        {"createShape": {"objectId": "cherr_msg", "shapeType": "TEXT_BOX",
                         "elementProperties": element_props(slide_id, 30, 100, 660, 200)}},
        {"insertText": {"objectId": "cherr_msg", "text": msg[:1500]}},
        {"updateTextStyle": {"objectId": "cherr_msg", "textRange": {"type": "ALL"},
                             "style": {"fontSize": pt(10),
                                       "foregroundColor": color_field(RED)},
                             "fields": "fontSize,foregroundColor"}},
    ]
    return reqs
