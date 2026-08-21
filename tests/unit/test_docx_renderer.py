import base64
from io import BytesIO

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

from send_email_studio.infrastructure.docx_renderer import render_docx, render_docx_with_images


_PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def test_docx_renderer_keeps_table_and_basic_formatting(tmp_path):
    path = tmp_path / "template.docx"
    document = Document()
    paragraph = document.add_paragraph()
    run = paragraph.add_run("标题")
    run.bold = True
    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "姓名"
    table.cell(0, 1).text = "{姓名}"
    document.save(path)
    html = render_docx(path)
    assert "<table" in html and "姓名" in html and "font-weight:700" in html


def test_docx_renderer_extracts_pasted_image_as_inline_cid(tmp_path):
    path = tmp_path / "image-template.docx"
    document = Document()
    document.add_paragraph("签名：")
    document.add_picture(BytesIO(_PNG_1X1), width=914400)
    document.save(path)

    result = render_docx_with_images(path)

    assert len(result.inline_images) == 1
    assert result.inline_images[0].data == _PNG_1X1
    assert result.inline_images[0].mime_type == "image/png"
    assert '<img src="cid:word-image-' in result.html


def test_docx_renderer_keeps_inherited_font_bold_and_spacing(tmp_path):
    path = tmp_path / "styled-template.docx"
    document = Document()
    document.styles["Normal"].font.name = "Microsoft YaHei"
    strong = document.styles.add_style("测试加粗", WD_STYLE_TYPE.CHARACTER)
    strong.font.bold = True
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = Pt(18)
    paragraph.add_run("普通文字")
    paragraph.add_run("加粗文字").style = strong
    document.save(path)

    html = render_docx(path)

    assert "font-family:Microsoft YaHei" in html
    assert "font-weight:700" in html
    assert "margin:0" in html
    assert "line-height:18.00pt" in html



def test_docx_renderer_keeps_paragraph_bottom_border(tmp_path):
    path = tmp_path / "paragraph-line.docx"
    document = Document()
    paragraph = document.add_paragraph()
    paragraph_properties = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:color"), "000000")
    borders.append(bottom)
    paragraph_properties.append(borders)
    document.save(path)

    html = render_docx(path)

    assert "border-bottom:1.50pt solid #000000" in html


def test_docx_renderer_converts_vml_line_shape(tmp_path):
    path = tmp_path / "shape-line.docx"
    document = Document()
    run = document.add_paragraph().add_run()
    run._r.append(
        parse_xml(
            '<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:v="urn:schemas-microsoft-com:vml">'
            '<v:shape type="#_x0000_t32" style="width:160pt;height:0pt" '
            'strokecolor="#000000" strokeweight="1.5pt"/>'
            '</w:pict>'
        )
    )
    document.save(path)

    html = render_docx(path)

    assert "width:160pt" in html
    assert "border-top:1.5pt solid #000000" in html



def test_docx_renderer_converts_word_horizontal_rule_object(tmp_path):
    path = tmp_path / "word-horizontal-rule.docx"
    document = Document()
    run = document.add_paragraph().add_run()
    run._r.append(
        parse_xml(
            '<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office">'
            '<v:rect style="width:104.9pt;height:0pt" o:hr="t" o:hralign="left" '
            'fillcolor="#000000" stroked="f"/>'
            '</w:pict>'
        )
    )
    document.save(path)

    html = render_docx(path)

    assert "width:104.9pt" in html
    assert "border-top:1pt solid #000000" in html
    assert "margin-left:0;margin-right:auto" in html
