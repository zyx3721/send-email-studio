from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping


@dataclass(frozen=True)
class InlineImage:
    """邮件正文中通过 Content-ID 引用的内嵌图片。"""

    content_id: str
    data: bytes
    mime_type: str
    filename: str = "image"
    width_px: int | None = None
    height_px: int | None = None


@dataclass(frozen=True)
class RecipientBatch:
    """一行 Excel 数据展开后的邮件收件人集合。"""

    row_number: int
    to: tuple[str, ...]
    cc: tuple[str, ...] = ()
    values: Mapping[str, Any] = field(default_factory=dict)
    attachments: tuple[Path, ...] = ()


@dataclass(frozen=True)
class MailSettings:
    smtp_host: str = "smtp.exmail.qq.com"
    smtp_port: int = 587
    smtp_user: str = ""
    smtp_password: str = ""
    subject: str = ""
    content_html: str = ""
    template_path: Path | None = None
    recipient_mode: str = "individual"  # individual 或 grouped
