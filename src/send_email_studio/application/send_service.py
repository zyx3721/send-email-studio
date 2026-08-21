from pathlib import Path
from typing import Callable

from ..domain.excel import build_batches
from ..domain.models import MailSettings
from ..infrastructure.excel_reader import read_excel_rows
from ..infrastructure.mail_sender import SmtpMailSender
from ..infrastructure.docx_renderer import render_docx


def _read_template(path: Path) -> str:
    if path.suffix.lower() == ".docx":
        return render_docx(path)
    return path.read_text(encoding="utf-8")


def send_from_excel(settings: MailSettings, excel_path: str | Path, log: Callable[[str], None] | None = None, progress: Callable[[int, int, str], None] | None = None) -> int:
    columns, rows = read_excel_rows(excel_path)
    workbook_path = Path(excel_path).resolve()
    batches = build_batches(rows, columns, settings.recipient_mode, workbook_path.parent)
    template = _read_template(settings.template_path) if settings.template_path else None
    SmtpMailSender(settings, log).send(batches, template, progress)
    return len(batches)
