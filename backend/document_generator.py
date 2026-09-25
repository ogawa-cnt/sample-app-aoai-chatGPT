import io
import json
import re

_ACCENT = "1F3A5F"  # 紺色(アクセントカラー)
_ACCENT_LIGHT = "E8EDF4"  # 薄い紺(表のヘッダー背景など)
_TEXT_DARK = "333333"

_INDENT_UNIT = 2  # 半角スペース何個で1階層とみなすか
_MAX_LIST_LEVEL = 3


def _leading_indent_level(line: str) -> int:
    stripped = line.lstrip(" ")
    indent = len(line) - len(stripped)
    return min(indent // _INDENT_UNIT, _MAX_LIST_LEVEL)


def _parse_markdown_blocks(markdown_text: str) -> list:
    # 見出し(#/##/###)・箇条書き(-、インデントで階層化)・番号付きリスト(1.)・
    # 表(| | |)・段落、という限定的なMarkdownの書式だけを対象にした簡易パーサー
    lines = markdown_text.splitlines()
    blocks = []
    i = 0
    n = len(lines)

    while i < n:
        raw_line = lines[i]
        stripped = raw_line.strip()
        if not stripped:
            i += 1
            continue

        heading_match = re.match(r"^(#{1,3})\s+(.*)$", stripped)
        if heading_match:
            blocks.append({
                "type": "heading",
                "level": len(heading_match.group(1)),
                "text": heading_match.group(2).strip(),
            })
            i += 1
            continue

        if re.match(r"^[-*]\s+", stripped):
            items = []
            while i < n and re.match(r"^[-*]\s+", lines[i].strip()):
                items.append({
                    "text": re.sub(r"^[-*]\s+", "", lines[i].strip()),
                    "level": _leading_indent_level(lines[i]),
                })
                i += 1
            blocks.append({"type": "bullet_list", "items": items})
            continue

        if re.match(r"^\d+\.\s+", stripped):
            items = []
            while i < n and re.match(r"^\d+\.\s+", lines[i].strip()):
                items.append({
                    "text": re.sub(r"^\d+\.\s+", "", lines[i].strip()),
                    "level": _leading_indent_level(lines[i]),
                })
                i += 1
            blocks.append({"type": "numbered_list", "items": items})
            continue

        if stripped.startswith("|"):
            table_lines = []
            while i < n and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            rows = [
                [cell.strip() for cell in line.strip("|").split("|")]
                for line in table_lines
            ]
            if len(rows) >= 2 and all(re.match(r"^:?-+:?$", c) for c in rows[1]):
                header, data_rows = rows[0], rows[2:]
            else:
                header, data_rows = rows[0], rows[1:]
            blocks.append({"type": "table", "header": header, "rows": data_rows})
            continue

        para_lines = [stripped]
        i += 1
        while i < n and lines[i].strip() and not re.match(r"^(#{1,3}\s+|[-*]\s+|\d+\.\s+|\|)", lines[i].strip()):
            para_lines.append(lines[i].strip())
            i += 1
        blocks.append({"type": "paragraph", "text": " ".join(para_lines)})

    return blocks


_DOCX_BULLET_STYLES = ["List Bullet", "List Bullet 2", "List Bullet 3", "List Bullet 3"]
_DOCX_NUMBER_STYLES = ["List Number", "List Number 2", "List Number 3", "List Number 3"]


def _generate_docx(markdown_text: str) -> bytes:
    from docx import Document

    document = Document()
    for block in _parse_markdown_blocks(markdown_text):
        if block["type"] == "heading":
            document.add_heading(block["text"], level=block["level"])
        elif block["type"] == "paragraph":
            document.add_paragraph(block["text"])
        elif block["type"] == "bullet_list":
            for item in block["items"]:
                style = _DOCX_BULLET_STYLES[min(item["level"], len(_DOCX_BULLET_STYLES) - 1)]
                document.add_paragraph(item["text"], style=style)
        elif block["type"] == "numbered_list":
            for item in block["items"]:
                style = _DOCX_NUMBER_STYLES[min(item["level"], len(_DOCX_NUMBER_STYLES) - 1)]
                document.add_paragraph(item["text"], style=style)
        elif block["type"] == "table":
            header = block["header"]
            table = document.add_table(rows=1, cols=len(header))
            table.style = "Light Grid Accent 1"
            for idx, cell_text in enumerate(header):
                table.rows[0].cells[idx].text = cell_text
            for row in block["rows"]:
                cells = table.add_row().cells
                for idx, cell_text in enumerate(row):
                    if idx < len(cells):
                        cells[idx].text = cell_text

    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue()


def _rgb(hex_color: str):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(hex_color)


def _add_tf_paragraph(text_frame, text: str, level: int = 0, bold: bool = False, bullet_char: str = None):
    if len(text_frame.paragraphs) == 1 and not text_frame.paragraphs[0].runs and not text_frame.paragraphs[0].text:
        paragraph = text_frame.paragraphs[0]
    else:
        paragraph = text_frame.add_paragraph()
    prefix = f"{bullet_char} " if bullet_char else ""
    paragraph.text = prefix + text
    paragraph.level = min(level, 4)
    if bold:
        for run in paragraph.runs:
            run.font.bold = True
    return paragraph


_BULLET_CHARS = ["●", "○", "・", "‐"]


def _add_pptx_table(slide, left, top, width, height, header, rows):
    from pptx.util import Pt

    num_rows = len(rows) + 1
    num_cols = len(header)
    graphic_frame = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
    table = graphic_frame.table

    for col_idx, text in enumerate(header):
        cell = table.cell(0, col_idx)
        cell.text = text
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(_ACCENT)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(14)
            paragraph.font.bold = True
            paragraph.font.color.rgb = _rgb("FFFFFF")

    for row_idx, row in enumerate(rows, start=1):
        for col_idx in range(num_cols):
            cell = table.cell(row_idx, col_idx)
            cell.text = row[col_idx] if col_idx < len(row) else ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(_ACCENT_LIGHT) if row_idx % 2 == 0 else _rgb("FFFFFF")
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(13)
                paragraph.font.color.rgb = _rgb(_TEXT_DARK)

    return graphic_frame


def _style_title_slide(prs, slide, title_text):
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN

    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(_ACCENT)

    title_shape = slide.shapes.title
    title_shape.left = Inches(0.8)
    title_shape.top = Inches(2.7)
    title_shape.width = prs.slide_width - Inches(1.6)
    title_shape.height = Inches(2.0)
    tf = title_shape.text_frame
    tf.word_wrap = True
    tf.text = title_text
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = _rgb("FFFFFF")

    # サブタイトルのプレースホルダーが存在する場合は非表示にする(使わないため)
    for shape in list(slide.placeholders):
        if shape.placeholder_format.idx != title_shape.placeholder_format.idx:
            shape._element.getparent().remove(shape._element)


_BODY_TOP = 1.35  # インチ。本文プレースホルダー・表の開始位置(タイトル+アクセントバーの下)


def _style_content_slide(prs, slide, title_text):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt

    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb("FFFFFF")

    title_shape = slide.shapes.title
    title_shape.left = Inches(0.5)
    title_shape.top = Inches(0.3)
    title_shape.width = prs.slide_width - Inches(1.0)
    title_shape.height = Inches(0.9)
    tf = title_shape.text_frame
    tf.word_wrap = True
    tf.text = title_text
    p = tf.paragraphs[0]
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = _rgb(_ACCENT)

    accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.05), Inches(1.4), Pt(3))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = _rgb(_ACCENT)
    accent_bar.line.fill.background()
    accent_bar.shadow.inherit = False

    # 本文プレースホルダーを、タイトル+アクセントバーの下に収まるよう位置調整する
    if len(slide.placeholders) > 1:
        body_placeholder = slide.placeholders[1]
        body_placeholder.left = Inches(0.5)
        body_placeholder.top = Inches(_BODY_TOP)
        body_placeholder.width = prs.slide_width - Inches(1.0)
        body_placeholder.height = prs.slide_height - Inches(_BODY_TOP) - Inches(0.4)


def _generate_pptx(markdown_text: str) -> bytes:
    from pptx import Presentation
    from pptx.util import Inches

    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)

    slide = None
    body_tf = None
    is_first_heading = True

    def ensure_slide():
        nonlocal slide, body_tf
        if slide is None:
            layout = presentation.slide_layouts[1]
            slide = presentation.slides.add_slide(layout)
            slide.shapes.title.text = ""
            _style_content_slide(presentation, slide, "")
            body_tf = slide.placeholders[1].text_frame
            body_tf.clear()

    for block in _parse_markdown_blocks(markdown_text):
        if block["type"] == "heading" and block["level"] <= 2:
            if is_first_heading:
                # 最初の見出しは表紙スライドとして扱う
                layout = presentation.slide_layouts[0]
                slide = presentation.slides.add_slide(layout)
                _style_title_slide(presentation, slide, block["text"])
                body_tf = None
                is_first_heading = False
                continue

            layout = presentation.slide_layouts[1]
            slide = presentation.slides.add_slide(layout)
            _style_content_slide(presentation, slide, block["text"])
            body_tf = slide.placeholders[1].text_frame
            body_tf.clear()
            continue

        if block["type"] == "table":
            ensure_slide()
            top = Inches(_BODY_TOP)
            left = Inches(0.6)
            width = presentation.slide_width - Inches(1.2)
            height = presentation.slide_height - top - Inches(0.5)
            _add_pptx_table(slide, left, top, width, height, block["header"], block["rows"])
            # 表を配置した後、同じスライドに追加のテキストが来ても表と重ならないようにする
            body_tf = None
            continue

        ensure_slide()
        if body_tf is None:
            continue

        if block["type"] == "heading":
            _add_tf_paragraph(body_tf, block["text"], bold=True)
        elif block["type"] == "paragraph":
            _add_tf_paragraph(body_tf, block["text"])
        elif block["type"] == "bullet_list":
            for item in block["items"]:
                bullet = _BULLET_CHARS[min(item["level"], len(_BULLET_CHARS) - 1)]
                _add_tf_paragraph(body_tf, item["text"], level=item["level"], bullet_char=bullet)
        elif block["type"] == "numbered_list":
            counters = {}
            for item in block["items"]:
                lvl = item["level"]
                counters[lvl] = counters.get(lvl, 0) + 1
                _add_tf_paragraph(body_tf, item["text"], level=lvl, bullet_char=f"{counters[lvl]}.")

    if slide is None:
        presentation.slides.add_slide(presentation.slide_layouts[1])

    buf = io.BytesIO()
    presentation.save(buf)
    return buf.getvalue()


def _sanitize_sheet_title(title: str) -> str:
    title = re.sub(r'[:\\/?*\[\]]', "_", title).strip()
    return title[:31] or "Sheet1"


def _generate_xlsx(markdown_text: str) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    workbook = Workbook()
    workbook.remove(workbook.active)
    current_sheet = None
    header_fill = PatternFill(start_color=_ACCENT, end_color=_ACCENT, fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for block in _parse_markdown_blocks(markdown_text):
        if block["type"] == "heading":
            current_sheet = workbook.create_sheet(title=_sanitize_sheet_title(block["text"]))
            continue

        if current_sheet is None:
            current_sheet = workbook.create_sheet(title="Sheet1")

        if block["type"] == "table":
            current_sheet.append(block["header"])
            header_row = current_sheet.max_row
            for col_idx in range(1, len(block["header"]) + 1):
                cell = current_sheet.cell(row=header_row, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
            for row in block["rows"]:
                current_sheet.append(row)
        elif block["type"] in ("bullet_list", "numbered_list"):
            for item in block["items"]:
                indent = "　" * item["level"]
                current_sheet.append([f"{indent}{item['text']}"])
        elif block["type"] == "paragraph":
            current_sheet.append([block["text"]])

    if not workbook.sheetnames:
        workbook.create_sheet(title="Sheet1")

    buf = io.BytesIO()
    workbook.save(buf)
    return buf.getvalue()


def generate_file(ext: str, content: str) -> bytes:
    ext = ext.lower()

    if ext in ("txt", "md"):
        return content.encode("utf-8")

    if ext == "json":
        parsed = json.loads(content)
        return json.dumps(parsed, ensure_ascii=False, indent=2).encode("utf-8")

    if ext == "html":
        return content.encode("utf-8")

    if ext == "docx":
        return _generate_docx(content)

    if ext == "xlsx":
        return _generate_xlsx(content)

    if ext == "pptx":
        return _generate_pptx(content)

    raise ValueError(f"Unsupported generation format: .{ext}")
