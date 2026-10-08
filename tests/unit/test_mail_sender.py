from email.mime.multipart import MIMEMultipart

import pytest

from send_email_studio.domain.models import InlineImage, MailSettings, RecipientBatch
from send_email_studio.infrastructure import mail_sender
from send_email_studio.infrastructure.mail_sender import render_template


DOTTED_COLUMNS = (
    "区域-A.1平台CPU",
    "区域-A.1平台内存",
    "区域-A.1平台硬盘",
    "区域-A.1平台许可",
    "区域-B.20平台CPU",
    "分支-C.5平台CPU",
    "分支-D.16平台CPU",
    "分支-D.70平台CPU",
)


@pytest.mark.parametrize("column", DOTTED_COLUMNS)
def test_render_template_accepts_column_names_with_dots(column):
    """含点的列名必须整段匹配，不能被 format 拆成属性访问。"""
    assert render_template("{" + column + "}", {column: "27%"}) == "27%"


def test_render_template_replaces_row_values_like_real_body():
    values = {
        "业务主体": "某中心",
        "区域-A.1平台CPU": "27%",
        "区域-A.1平台内存": "73%",
        "区域-B.20平台CPU": "21%",
        "分支-D.70平台许可": "无",
    }
    body = (
        "【{业务主体}】资源使用率：A.1 CPU {区域-A.1平台CPU}、"
        "内存 {区域-A.1平台内存}；B.20 CPU {区域-B.20平台CPU}；"
        "D.70 许可 {分支-D.70平台许可}。"
    )

    assert render_template(body, values) == (
        "【某中心】资源使用率：A.1 CPU 27%、内存 73%；B.20 CPU 21%；D.70 许可 无。"
    )


def test_render_template_keeps_unknown_placeholder():
    assert render_template("值：{不存在的列}", {"姓名": "张三"}) == "值：{不存在的列}"


def test_render_template_still_supports_format_spec_and_conversion():
    values = {"数量": 3, "单价": 7.95, "名称": "张三"}
    assert render_template("{数量} × {单价:.2f}", values) == "3 × 7.95"
    assert render_template("{名称!r}", values) == "'张三'"


def test_render_template_handles_braces_and_plain_text():
    assert render_template("没有占位符", {}) == "没有占位符"


class _FakeSmtp:
    sent_messages = []

    def __init__(self, *_args, **_kwargs):
        self.messages = []

    def ehlo(self):
        pass

    def starttls(self):
        pass

    def login(self, *_args):
        pass

    def send_message(self, message, **_kwargs):
        self.messages.append(message)
        self.sent_messages.append(message)

    def quit(self):
        pass


def test_sender_renders_dotted_column_in_body(monkeypatch):
    """端到端：正文模板用含点列名时也要能正常发出。"""
    _FakeSmtp.sent_messages = []
    monkeypatch.setattr(mail_sender.smtplib, "SMTP", _FakeSmtp)
    settings = MailSettings(smtp_user="sender@example.com", smtp_password="secret")
    batch = RecipientBatch(
        2,
        ("recipient@example.com",),
        values={"区域-A.1平台CPU": "27%"},
    )

    mail_sender.SmtpMailSender(settings).send([batch], "<p>CPU：{区域-A.1平台CPU}</p>")

    body = _FakeSmtp.sent_messages[0].get_payload(0).get_payload(decode=True).decode("utf-8")
    assert "CPU：27%" in body


def test_sender_builds_related_mime_for_inline_images(monkeypatch):
    _FakeSmtp.sent_messages = []
    monkeypatch.setattr(mail_sender.smtplib, "SMTP", _FakeSmtp)
    settings = MailSettings(smtp_user="sender@example.com", smtp_password="secret")
    batch = RecipientBatch(2, ("recipient@example.com",), values={"姓名": "张三"})
    image = InlineImage("word-image-test", b"png-data", "image/png", "image.png")

    mail_sender.SmtpMailSender(settings).send(
        [batch],
        '<p>您好，{姓名}</p><img src="cid:word-image-test">',
        inline_images=(image,),
    )

    message = _FakeSmtp.sent_messages[0]
    related = next(part for part in message.get_payload() if isinstance(part, MIMEMultipart))
    assert related.get_content_type() == "multipart/related"
    assert "张三" in related.get_payload(0).get_payload(decode=True).decode("utf-8")
    image_part = related.get_payload(1)
    assert image_part["Content-ID"] == "<word-image-test>"
    assert image_part.get_content_type() == "image/png"
