import re
from pathlib import Path
from typing import Any, Iterable, Mapping

from .models import RecipientBatch

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
SPLIT_RE = re.compile(r"[,;、]+")
RECIPIENT_KEYWORDS = ("收件人", "收件", "收件名", "收件员工", "收件邮箱", "人员邮箱", "员工邮箱")


def split_emails(value: Any) -> list[str]:
    if value is None:
        return []
    result: list[str] = []
    for item in SPLIT_RE.split(str(value)):
        email = item.strip()
        if email and EMAIL_RE.fullmatch(email) and email not in result:
            result.append(email)
    return result


def matching_columns(columns: Iterable[Any], keywords: tuple[str, ...]) -> list[str]:
    return [str(column) for column in columns if any(keyword in str(column) for keyword in keywords)]


def _values(row: Mapping[str, Any], columns: Iterable[str]) -> list[str]:
    values: list[str] = []
    for column in columns:
        values.extend(split_emails(row.get(column)))
    return list(dict.fromkeys(values))


def build_batches(
    rows: Iterable[Mapping[str, Any]],
    columns: Iterable[Any],
    recipient_mode: str = "individual",
    attachment_base_dir: Path | None = None,
) -> list[RecipientBatch]:
    columns = [str(column) for column in columns]
    to_columns = matching_columns(columns, RECIPIENT_KEYWORDS)
    cc_columns = matching_columns(columns, ("抄送",))
    attachment_columns = matching_columns(columns, ("附件",))
    if not to_columns:
        raise ValueError("Excel 中未找到列名包含“收件”的收件人列")
    batches: list[RecipientBatch] = []
    for row_number, raw_row in enumerate(rows, 2):
        row = {str(k): v for k, v in raw_row.items()}
        recipients = _values(row, to_columns)
        if not recipients:
            continue
        cc = tuple(_values(row, cc_columns))
        attachments = tuple(
            _attachment_path(str(row[column]).strip(), attachment_base_dir)
            for column in attachment_columns
            if str(row.get(column, "")).strip()
        )
        targets = recipients if recipient_mode == "individual" else [", ".join(recipients)]
        for target in targets:
            batches.append(RecipientBatch(row_number, tuple(split_emails(target)), cc, row, attachments))
    return batches


def _attachment_path(value: str, base_dir: Path | None) -> Path:
    path = Path(value)
    if path.is_absolute() or base_dir is None:
        return path
    return base_dir / path
