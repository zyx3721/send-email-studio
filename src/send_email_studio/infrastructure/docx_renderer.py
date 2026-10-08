"""将 Word 文档中的段落、表格和图片转换为邮件可用的 HTML。"""

from dataclasses import dataclass
from html import escape
from pathlib import Path
import re
from uuid import uuid4

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from ..domain.models import InlineImage


@dataclass(frozen=True)
class DocxRenderResult:
    html: str
    inline_images: tuple[InlineImage, ...] = ()


def render_docx(path: Path) -> str:
    """兼容旧调用方，仅返回渲染后的 HTML。"""
    return render_docx_with_images(path).html


def render_docx_with_images(path: Path) -> DocxRenderResult:
    document = Document(path)
    inline_images: list[InlineImage] = []
    blocks = [_render_block(block, inline_images) for block in _iter_block_items(document)]
    html = '<div style="font-family:Microsoft YaHei;font-size:12pt;color:#000;line-height:normal;">' + "".join(blocks) + "</div>"
    return DocxRenderResult(html, tuple(inline_images))


def _iter_block_items(parent: DocumentType | _Cell):
    parent_element = parent.element.body if isinstance(parent, DocumentType) else parent._tc
    for child in parent_element.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


def _render_block(block: Paragraph | Table, inline_images: list[InlineImage]) -> str:
    if isinstance(block, Table):
        return _render_table(block, inline_images)
    style = ["margin:0"]
    style.extend(_paragraph_border_styles(block))
    if block.alignment is not None:
        style.append(f"text-align:{_alignment(block.alignment)}")
    fmt = block.paragraph_format
    if fmt.left_indent:
        style.append(f"margin-left:{fmt.left_indent.pt:.2f}pt")
    if fmt.first_line_indent:
        style.append(f"text-indent:{fmt.first_line_indent.pt:.2f}pt")
    if fmt.space_before:
        style.append(f"margin-top:{fmt.space_before.pt:.2f}pt")
    if fmt.space_after:
        style.append(f"margin-bottom:{fmt.space_after.pt:.2f}pt")
    if fmt.line_spacing:
        if isinstance(fmt.line_spacing, float):
            style.append(f"line-height:{fmt.line_spacing:.2f}")
        elif hasattr(fmt.line_spacing, "pt"):
            style.append(f"line-height:{fmt.line_spacing.pt:.2f}pt")
    content = "".join(_render_run(group, block, inline_images) for group in _run_groups(block)) or "&nbsp;"
    style_text = " ;".join(style)
    return f'<p style="{style_text}">{content}</p>'


def _run_groups(paragraph: Paragraph) -> list[list]:
    """按“{”是否闭合把 run 分组，避免占位符被切进多个 span。"""
    groups: list[list] = []
    pending: list = []
    for run in paragraph.runs:
        pending.append(run)
        if _has_unclosed_brace("".join(item.text for item in pending)):
            continue
        groups.append(pending)
        pending = []
    if pending:
        groups.append(pending)
    return groups


def _has_unclosed_brace(text: str) -> bool:
    return text.rfind("{") > text.rfind("}")


def _render_run(runs: list, paragraph: Paragraph, inline_images: list[InlineImage]) -> str:
    run = runs[0]
    styles = []
    font_name = _effective_font_name(run, paragraph)
    if font_name:
        styles.append(f"font-family:{escape(font_name, quote=True)}")
    font_size = _effective_font_value(run, paragraph, "size")
    if font_size:
        styles.append(f"font-size:{font_size.pt:.2f}pt")
    if run.font.color and run.font.color.rgb:
        styles.append(f"color:#{run.font.color.rgb}")
    if _effective_font_value(run, paragraph, "bold"):
        styles.append("font-weight:700")
    if _effective_font_value(run, paragraph, "italic"):
        styles.append("font-style:italic")
    decorations = []
    if _effective_font_value(run, paragraph, "underline"):
        decorations.append("underline")
    if _effective_font_value(run, paragraph, "strike"):
        decorations.append("line-through")
    if decorations:
        styles.append(f"text-decoration:{' '.join(decorations)}")
    text = escape("".join(item.text for item in runs)).replace(" ", "&nbsp;").replace("\n", "<br>")
    objects_html = "".join(html for item in runs for html in _render_image(item, inline_images))
    objects_html += "".join(html for item in runs for html in _render_horizontal_shapes(item))
    style_text = " ;".join(styles)
    return f'<span style="{style_text}">{text}{objects_html}</span>'


def _effective_font_name(run, paragraph: Paragraph) -> str | None:
    for name in (
        run.font.name,
        _style_font_name(run.style),
        _style_font_name(paragraph.style),
        _style_font_name(_normal_style(run)),
    ):
        if name:
            return name
    return None


def _effective_font_value(run, paragraph: Paragraph, attribute: str):
    direct_value = getattr(run.font, attribute)
    if direct_value is not None:
        return direct_value
    for style in (run.style, paragraph.style, _normal_style(run)):
        value = _style_font_value(style, attribute)
        if value is not None:
            return value
    return None


def _normal_style(run):
    styles = getattr(run.part, "styles", None)
    if styles is None:
        return None
    try:
        return styles["Normal"]
    except KeyError:
        return None


def _style_font_value(style, attribute: str):
    current = style
    while current is not None:
        value = getattr(current.font, attribute)
        if value is not None:
            return value
        current = current.base_style
    return None


def _style_font_name(style) -> str | None:
    if style is None:
        return None
    name = _style_font_value(style, "name")
    if name:
        return name
    current = style
    while current is not None:
        rpr = current.element.find(qn("w:rPr"))
        if rpr is not None:
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is not None:
                for attribute in ("eastAsia", "ascii", "hAnsi", "cs"):
                    value = rfonts.get(qn(f"w:{attribute}"))
                    if value:
                        return value
        current = current.base_style
    return None



def _paragraph_border_styles(paragraph: Paragraph) -> list[str]:
    """将 Word 段落边框转换为邮件客户端普遍支持的 CSS 边框。"""
    paragraph_properties = paragraph._p.pPr
    if paragraph_properties is None:
        return []
    borders = paragraph_properties.find(qn("w:pBdr"))
    if borders is None:
        return []

    styles: list[str] = []
    for edge in ("top", "right", "bottom", "left"):
        border = borders.find(qn(f"w:{edge}"))
        if border is None:
            continue
        border_type = border.get(qn("w:val"), "single")
        if border_type in {"nil", "none"}:
            continue
        try:
            width_pt = max(int(border.get(qn("w:sz"), "4")) / 8, 0.5)
        except ValueError:
            width_pt = 0.5
        css_type = {
            "double": "double",
            "dashed": "dashed",
            "dashSmallGap": "dashed",
            "dotted": "dotted",
        }.get(border_type, "solid")
        color = border.get(qn("w:color"), "000000")
        if color == "auto":
            color = "000000"
        styles.append(f"border-{edge}:{width_pt:.2f}pt {css_type} #{color}")
        space = border.get(qn("w:space"))
        if space and edge in {"top", "bottom"}:
            styles.append(f"padding-{edge}:{space}pt")
    return styles


def _render_horizontal_shapes(run) -> list[str]:
    """将 Word/WPS 中常见的 VML 或 DrawingML 横线形状转换为 HTML。"""
    html: list[str] = []
    for shape in run._element.iter("{urn:schemas-microsoft-com:vml}shape"):
        shape_style = shape.get("style", "")
        width = _shape_css_dimension(shape_style, "width") or "100%"
        height = _shape_numeric_dimension(shape_style, "height")
        shape_type = shape.get("type", "")
        if "_x0000_t32" not in shape_type and height is not None and height > 8:
            continue
        thickness = shape.get("strokeweight", "1pt")
        color = _css_color(shape.get("strokecolor", "#000000"))
        html.append(_horizontal_line_html(width, thickness, color))

    office_namespace = "{urn:schemas-microsoft-com:office:office}"
    for rectangle in run._element.iter("{urn:schemas-microsoft-com:vml}rect"):
        rectangle_style = rectangle.get("style", "")
        is_horizontal_rule = rectangle.get(office_namespace + "hr", "").lower() in {"t", "true", "1"}
        height = _shape_numeric_dimension(rectangle_style, "height")
        if not is_horizontal_rule and (height is None or height > 8):
            continue
        width = _shape_css_dimension(rectangle_style, "width")
        if width is None:
            width = _horizontal_rule_percent(rectangle.get(office_namespace + "hrpct"))
        thickness = _shape_css_dimension(rectangle_style, "height")
        if thickness is None or (height is not None and height <= 0):
            thickness = rectangle.get("strokeweight", "1pt")
        color = _css_color(rectangle.get("fillcolor") or rectangle.get("strokecolor") or "#000000")
        alignment = rectangle.get(office_namespace + "hralign", "left")
        html.append(_horizontal_line_html(width or "100%", thickness, color, alignment))

    for drawing in run._element.iter(qn("w:drawing")):
        if next(drawing.iter(qn("a:blip")), None) is not None:
            continue
        line = next(drawing.iter(qn("a:ln")), None)
        if line is None:
            continue
        width_px, height_px = _image_size_px(drawing)
        if not width_px or (height_px is not None and height_px > 12):
            continue
        try:
            thickness = f"{max(int(line.get('w', '12700')) / 12700, 0.5):.2f}pt"
        except ValueError:
            thickness = "1pt"
        color_element = next(line.iter(qn("a:srgbClr")), None)
        color = _css_color(color_element.get("val") if color_element is not None else "000000")
        html.append(_horizontal_line_html(f"{width_px}px", thickness, color))
    return html


def _shape_css_dimension(style: str, name: str) -> str | None:
    match = re.search(rf"(?:^|;)\s*{name}:\s*([0-9.]+(?:pt|px|cm|mm|in))", style, re.IGNORECASE)
    return match.group(1) if match else None


def _shape_numeric_dimension(style: str, name: str) -> float | None:
    value = _shape_css_dimension(style, name)
    if value is None:
        return None
    match = re.match(r"[0-9.]+", value)
    return float(match.group(0)) if match else None


def _css_color(value: str | None) -> str:
    color = (value or "000000").strip()
    if color.lower() in {"auto", "windowtext"}:
        return "#000000"
    return color if color.startswith("#") else f"#{color}"


def _horizontal_rule_percent(value: str | None) -> str | None:
    if not value:
        return None
    try:
        number = float(value)
    except ValueError:
        return None
    percentage = number / 10 if number > 100 else number
    return f"{min(max(percentage, 0), 100):g}%"


def _horizontal_line_html(width: str, thickness: str, color: str, alignment: str = "left") -> str:
    margins = {
        "center": "margin-left:auto;margin-right:auto",
        "right": "margin-left:auto;margin-right:0",
    }.get(alignment.lower(), "margin-left:0;margin-right:auto")
    return (
        f'<span style="display:block;width:{escape(width, quote=True)};max-width:100%;{margins};'
        f'height:0;border-top:{escape(thickness, quote=True)} solid {escape(color, quote=True)};"></span>'
    )

def _render_table(table: Table, inline_images: list[InlineImage]) -> str:
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            content = "".join(_render_block(block, inline_images) for block in _iter_block_items(cell))
            cells.append(f'<td style="border:1px solid #000;padding:4pt;vertical-align:middle;">{content or "&nbsp;"}</td>')
        rows.append(f"<tr>{''.join(cells)}</tr>")
    return f'<table style="border-collapse:collapse;margin:8pt 0;">{''.join(rows)}</table>'


def _render_image(run, inline_images: list[InlineImage]) -> list[str]:
    """提取 run 中的图片，并生成对应的 CID HTML 引用。"""
    html: list[str] = []
    for blip in run._element.iter(qn("a:blip")):
        relationship_id = blip.get(qn("r:embed"))
        if not relationship_id:
            continue
        image_part = run.part.related_parts.get(relationship_id)
        if image_part is None:
            raise ValueError(f"Word 图片关系不存在：{relationship_id}")

        content_id = f"word-image-{uuid4().hex}"
        width_px, height_px = _image_size_px(run._element)
        inline_images.append(
            InlineImage(
                content_id=content_id,
                data=image_part.blob,
                mime_type=image_part.content_type,
                filename=Path(image_part.partname).name,
                width_px=width_px,
                height_px=height_px,
            )
        )
        size_style = []
        if width_px:
            size_style.append(f"width:{width_px}px")
        if height_px:
            size_style.append(f"height:{height_px}px")
        size_style.append("max-width:100%")
        size_text = " ;".join(size_style)
        html.append(
            f'<img src="cid:{escape(content_id, quote=True)}" '
            f'style="{size_text};vertical-align:middle;" alt="">'
        )
    return html


def _image_size_px(element) -> tuple[int | None, int | None]:
    """读取嵌入型或浮动图片的 EMU 尺寸，转换为 CSS 像素。"""
    extent = next(element.iter(qn("wp:extent")), None)
    if extent is None:
        extent = next(element.iter(qn("a:ext")), None)
    if extent is None:
        return None, None
    try:
        width = round(int(extent.get("cx", "0")) / 9525)
        height = round(int(extent.get("cy", "0")) / 9525)
    except (TypeError, ValueError):
        return None, None
    return (width or None), (height or None)


def _alignment(value) -> str:
    mapping = {0: "left", 1: "center", 2: "right", 3: "justify"}
    return mapping.get(getattr(value, "value", value), "left")
