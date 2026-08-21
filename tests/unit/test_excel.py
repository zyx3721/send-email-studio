import pytest
from openpyxl import Workbook
from send_email_studio.domain.excel import build_batches, split_emails
from send_email_studio.infrastructure.excel_reader import read_excel_rows
from send_email_studio.infrastructure.mail_sender import subject_for_batch


def test_split_emails_supports_delimiters_and_deduplicates():
    assert split_emails(" a@example.com,b@example.com;a@example.com、bad ") == ["a@example.com", "b@example.com"]


def test_build_batches_individual_and_cc():
    rows = [{"姓名": "甲", "主要收件邮箱": "a@example.com,b@example.com", "抄送人": "c@example.com;d@example.com"}]
    batches = build_batches(rows, rows[0].keys(), "individual")
    assert [batch.to for batch in batches] == [("a@example.com",), ("b@example.com",)]
    assert batches[0].cc == ("c@example.com", "d@example.com")


def test_build_batches_grouped():
    rows = [{"收件邮箱": "a@example.com,b@example.com"}]
    assert build_batches(rows, rows[0].keys(), "grouped")[0].to == ("a@example.com", "b@example.com")


def test_missing_column():
    with pytest.raises(ValueError): build_batches([{"姓名": "甲"}], ["姓名"])


def test_excel_subject_column_overrides_default_subject():
    batches = build_batches([{"收件邮箱": "a@example.com", "邮件主题": "本行主题"}], ["收件邮箱", "邮件主题"])
    assert subject_for_batch(batches[0], "工具主题") == "本行主题"


def test_relative_attachment_is_based_on_workbook_directory(tmp_path):
    rows = [{"收件邮箱": "a@example.com", "附件文件": "attachments/report.xlsx"}]
    batches = build_batches(rows, rows[0].keys(), attachment_base_dir=tmp_path)
    assert batches[0].attachments[0] == tmp_path / "attachments/report.xlsx"


def test_excel_percentage_keeps_display_format(tmp_path):
    path = tmp_path / "percent.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["收件邮箱", "利用率"])
    sheet.append(["a@example.com", 0.74])
    sheet["B2"].number_format = "0%"
    workbook.save(path)

    columns, rows = read_excel_rows(path)

    assert columns == ["收件邮箱", "利用率"]
    assert rows[0]["利用率"] == "74%"
