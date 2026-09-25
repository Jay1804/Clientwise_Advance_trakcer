"""SMTP email sending for split report files.

Replaces the old Outlook COM automation (outlook_email.py) - send_email is
the only function that actually talks to SMTP, so callers in app.py and
scheduled_tracker_runner.py don't need to know how the email is sent.
Credentials come from .env (SMTP_HOST, SMTP_PORT, EMAIL_SENDER,
EMAIL_PASSWORD) - load_dotenv() is called here too so this module works
whether or not the caller already loaded .env (e.g. scheduled_tracker_runner.py
runs standalone via Task Scheduler and doesn't call it itself).
"""

import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv

load_dotenv()


def _split_addresses(value: str) -> list[str]:
    """Splits a To/CC field on ';' or ',' - the app's placeholder text asks
    for semicolon-separated addresses, but comma-separated input is accepted
    too since that's the more common convention elsewhere."""
    if not value:
        return []
    normalized = value.replace(",", ";")
    return [addr.strip() for addr in normalized.split(";") if addr.strip()]


def email_available() -> bool:
    """True if SMTP_HOST/EMAIL_SENDER/EMAIL_PASSWORD are all set in the
    environment. Checked lazily so the UI can disable emailing instead of
    crashing when they're missing."""
    return bool(os.getenv("SMTP_HOST") and os.getenv("EMAIL_SENDER") and os.getenv("EMAIL_PASSWORD"))


def send_email(
    to_address: str,
    cc_address: str,
    subject: str,
    body: str,
    attachment_bytes: bytes,
    attachment_filename: str,
) -> None:
    """Compose and send one email over SMTP with a single attachment."""
    host = os.getenv("SMTP_HOST")
    port = int(os.getenv("SMTP_PORT", "587"))
    sender = os.getenv("EMAIL_SENDER")
    password = os.getenv("EMAIL_PASSWORD")

    to_list = _split_addresses(to_address)
    cc_list = _split_addresses(cc_address)
    if not to_list:
        raise ValueError("No valid To_address to send to.")

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.attach(MIMEText(body))

    attachment = MIMEApplication(attachment_bytes, Name=attachment_filename)
    attachment["Content-Disposition"] = f'attachment; filename="{attachment_filename}"'
    msg.attach(attachment)

    with smtplib.SMTP(host, port, timeout=30) as server:
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, to_list + cc_list, msg.as_string())
