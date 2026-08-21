import mimetypes
import smtplib
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from typing import Callable

from ..domain.models import MailSettings, RecipientBatch


def render_template(template: str, values: dict[str, object]) -> str:
    return template.format_map(_SafeValues(values))


def subject_for_batch(batch: RecipientBatch, fallback: str) -> str:
    """Excel 列名包含“主题”时，按当前行覆盖界面主题。"""
    subject_columns = [key for key in batch.values if "主题" in str(key)]
    if subject_columns:
        value = batch.values.get(subject_columns[0], "")
        return "" if value is None else str(value).strip()
    return fallback


class _SafeValues(dict):
    def __missing__(self, key: str) -> str:
        return "{" + key + "}"


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


class SmtpMailSender:
    def __init__(self, settings: MailSettings, log: Callable[[str], None] | None = None):
        self.settings = settings
        self.log = log or (lambda _: None)

    def send(self, batches: list[RecipientBatch], template: str | None = None, progress: Callable[[int, int, str], None] | None = None) -> None:
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
