"""Email delivery of completed audit workbooks.

The Streamlit app holds the finished workbook only in session memory, which is
lost when Streamlit Community Cloud winds the container down before the user
downloads it.  Emailing the workbook as an attachment the moment the run
finishes decouples delivery from the live session, so the artifact survives
even if the app is later evicted.

This module is intentionally UI-free and dependency-free (standard-library
``smtplib`` only) so it can be unit-tested without Streamlit or a real SMTP
server.  The app wraps it with config from ``st.secrets`` and error handling.
"""
from __future__ import annotations

import re
import smtplib
import ssl
from email.message import EmailMessage

# openpyxl workbook MIME type, split into the main/sub parts add_attachment wants
_XLSX_MAINTYPE = "application"
_XLSX_SUBTYPE = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Pragmatic address check — not RFC-complete, just enough to catch typos before
# we bother the SMTP server.
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class EmailConfigError(ValueError):
    """Raised when required SMTP settings are missing."""


def is_valid_email(address: str) -> bool:
    """Return True if *address* looks like a plausible single email address."""
    return bool(address) and bool(_EMAIL_RE.match(address.strip()))


def send_audit_email(
    *,
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    sender: str,
    recipient: str,
    subject: str,
    body: str,
    attachment_bytes: bytes,
    attachment_filename: str,
    timeout: int = 60,
) -> None:
    """Send *attachment_bytes* as an .xlsx attachment via SMTP over STARTTLS.

    Raises ``EmailConfigError`` if any required field is missing and lets
    ``smtplib``/socket errors propagate so the caller can report them.
    """
    missing = [
        name
        for name, value in (
            ("SMTP_HOST", smtp_host),
            ("SMTP_USER", smtp_user),
            ("SMTP_PASSWORD", smtp_password),
            ("EMAIL_FROM", sender),
        )
        if not value
    ]
    if missing:
        raise EmailConfigError(
            "Missing SMTP settings: " + ", ".join(missing)
        )
    if not is_valid_email(recipient):
        raise EmailConfigError(f"Invalid recipient address: {recipient!r}")

    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    if attachment_bytes:
        message.add_attachment(
            attachment_bytes,
            maintype=_XLSX_MAINTYPE,
            subtype=_XLSX_SUBTYPE,
            filename=attachment_filename,
        )

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_host, int(smtp_port), timeout=timeout) as server:
        server.starttls(context=context)
        server.login(smtp_user, smtp_password)
        server.send_message(message)
