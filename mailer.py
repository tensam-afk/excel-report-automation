import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

REQUIRED = ["SMTP_USER", "SMTP_PASSWORD"]


def is_configured():
    return all(os.getenv(k) for k in REQUIRED)


def send_report(path, to, subject, body):
    path = Path(path)
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = os.getenv("SMTP_FROM") or os.environ["SMTP_USER"]
    msg["To"] = to
    msg.set_content(body)
    msg.add_attachment(
        path.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name,
    )

    with smtplib.SMTP(os.getenv("SMTP_HOST", "smtp.gmail.com"), int(os.getenv("SMTP_PORT", 587)), timeout=30) as s:
        s.starttls()
        s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASSWORD"])
        s.send_message(msg)
