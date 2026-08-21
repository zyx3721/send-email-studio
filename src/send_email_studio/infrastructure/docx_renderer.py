"""将 Word 文档中的段落和表格转换为邮件可用的 HTML。"""

from html import escape
from pathlib import Path

from docx.document import Document as DocumentType
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx import Document


def render_docx(path: Path) -> str:
    document = Document(path)
    blocks = [_render_block(block) for block in _iter_block_items(document)]
    return "<div style=\"font-family:宋体;font-size:12pt;color:#000;\">" + "".join(blocks) + "</div>"


def _iter_block_items(parent: DocumentType | _Cell):
    parent_element = parent.element.body if isinstance(parent, DocumentType) else parent._tc
    for child in parent_element.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


def _render_block(block: Paragraph | Table) -> str:
    if isinstance(block, Table):
        return _render_table(block)
    style = []
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
    if fmt.line_spacing and isinstance(fmt.line_spacing, float):
        style.append(f"line-height:{fmt.line_spacing:.2f}")
    content = "".join(_render_run(run) for run in block.runs) or "&nbsp;"
    return f"<p style=\"{' ;'.join(style)}\">{content}</p>"


def _render_run(run) -> str:
    styles = []
    if run.font.name:
        styles.append(f"font-family:{escape(run.font.name)}")
    if run.font.size:
        styles.append(f"font-size:{run.font.size.pt:.2f}pt")
    if run.font.color and run.font.color.rgb:
        styles.append(f"color:#{run.font.color.rgb}")
    if run.bold:
        styles.append("font-weight:700")
    if run.italic:
        styles.append("font-style:italic")
    decorations = []
    if run.underline:
        decorations.append("underline")
    if run.font.strike:
        decorations.append("line-through")
    if decorations:
        styles.append(f"text-decoration:{' '.join(decorations)}")
    text = escape(run.text).replace(" ", "&nbsp;").replace("\n", "<br>")
    return f"<span style=\"{' ;'.join(styles)}\">{text}</span>"


def _render_table(table: Table) -> str:
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            content = "".join(_render_block(block) for block in _iter_block_items(cell))
            cells.append(f"<td style=\"border:1px solid #000;padding:4pt;vertical-align:middle;\">{content or '&nbsp;'}</td>")
        rows.append(f"<tr>{''.join(cells)}</tr>")
    return f"<table style=\"border-collapse:collapse;margin:8pt 0;\">{''.join(rows)}</table>"


def _alignment(value) -> str:
    mapping = {0: "left", 1: "center", 2: "right", 3: "justify"}
    return mapping.get(getattr(value, "value", value), "left")

