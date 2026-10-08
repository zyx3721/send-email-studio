import mimetypes
import smtplib
import string
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Callable

from ..domain.models import InlineImage, MailSettings, RecipientBatch


def render_template(template: str, values: dict[str, object]) -> str:
    """按整段列名替换 {列名}，避免列名里的点被 format 当成属性访问。"""
    return _ColumnFormatter(values).vformat(template, (), {})


class _ColumnFormatter(string.Formatter):
    """整段匹配字段名；未提供的列名原样保留。"""

    def __init__(self, values: dict[str, object]) -> None:
        super().__init__()
        self._values = values

    def get_field(self, field_name, args, kwargs):
        if field_name in self._values:
            return self._values[field_name], field_name
        return "{" + field_name + "}", field_name


def subject_for_batch(batch: RecipientBatch, fallback: str) -> str:
    """Excel 列名包含“主题”时，按当前行覆盖界面主题。"""
    subject_columns = [key for key in batch.values if "主题" in str(key)]
    if subject_columns:
        value = batch.values.get(subject_columns[0], "")
        return "" if value is None else str(value).strip()
    return fallback


def _attachment(path: Path) -> MIMEBase:
    if not path.is_file():
        raise FileNotFoundError(f"附件不存在：{path}")
    content_type, _ = mimetypes.guess_type(path.name)
    maintype, subtype = (content_type or "application/octet-stream").split("/", 1)
    part = MIMEBase(maintype, subtype)
    part.set_payload(path.read_bytes())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=Header(path.name, "utf-8").encode())
    return part


def _inline_image(image: InlineImage) -> MIMEBase:
    maintype, subtype = (image.mime_type or "application/octet-stream").split("/", 1)
    if maintype == "image":
        part: MIMEBase = MIMEImage(image.data, _subtype=subtype)
    else:
        part = MIMEBase(maintype, subtype)
        part.set_payload(image.data)
        encoders.encode_base64(part)
    part.add_header("Content-ID", f"<{image.content_id}>")
    part.add_header("Content-Disposition", "inline", filename=Header(image.filename, "utf-8").encode())
    return part


class SmtpMailSender:
    def __init__(self, settings: MailSettings, log: Callable[[str], None] | None = None):
        self.settings = settings
        self.log = log or (lambda _: None)

    def send(
        self,
        batches: list[RecipientBatch],
        template: str | None = None,
        progress: Callable[[int, int, str], None] | None = None,
        inline_images: tuple[InlineImage, ...] = (),
    ) -> None:
        if not batches:
            raise ValueError("Excel 中没有可发送的收件人")
        total = len(batches)
        server = smtplib.SMTP(self.settings.smtp_host, self.settings.smtp_port, timeout=20)
        try:
            server.ehlo()
            if self.settings.smtp_port != 25:
                server.starttls()
                server.ehlo()
            server.login(self.settings.smtp_user, self.settings.smtp_password)
            for index, batch in enumerate(batches, 1):
                body = render_template(template, dict(batch.values)) if template is not None else self.settings.content_html
                message = MIMEMultipart()
                message["From"] = self.settings.smtp_user
                message["To"] = ", ".join(batch.to)
                if batch.cc:
                    message["Cc"] = ", ".join(batch.cc)
                message["Subject"] = subject_for_batch(batch, self.settings.subject)
                if inline_images:
                    related = MIMEMultipart("related")
                    related.attach(MIMEText(body, "html", "utf-8"))
                    for image in inline_images:
                        related.attach(_inline_image(image))
                    message.attach(related)
                else:
                    message.attach(MIMEText(body, "html", "utf-8"))
                for path in batch.attachments:
                    message.attach(_attachment(path))
                server.send_message(message, from_addr=self.settings.smtp_user, to_addrs=list(batch.to) + list(batch.cc))
                text = f"第 {batch.row_number} 行邮件发送成功：{', '.join(batch.to)}"
                self.log(text)
                if progress:
                    progress(index, total, text)
        finally:
            try:
                server.quit()
            except smtplib.SMTPException:
                server.close()
