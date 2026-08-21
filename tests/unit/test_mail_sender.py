from email.mime.multipart import MIMEMultipart

from send_email_studio.domain.models import InlineImage, MailSettings, RecipientBatch
from send_email_studio.infrastructure import mail_sender


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
