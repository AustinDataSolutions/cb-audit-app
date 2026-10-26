"""Email delivery of completed audit workbooks.

The Streamlit app holds the finished workbook only in session memory, which is
lost when Streamlit Community Cloud winds the container down before the user
downloads it.  Emailing the workbook as an attachment the moment the run
finishes decouples delivery from the live session, so the artifact survives
even if the app is later evicted.

This module is intentionally UI-free and dependency-free (standard-library
``smtplib`` only) so it can be unit-tested without Streamlit or a real SMTP
server.  The app wraps it with the settings from the environment and error handling.
"""
from __future__ import annotations

import re
import smtplib
import socket
import ssl
import time
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


def check_smtp_login(
    *,
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    timeout: int = 15,
) -> str | None:
    """Connect to the SMTP server, start encryption, sign in, and disconnect without sending.

    Returns ``None`` when the server accepts the sign-in, otherwise a short plain-English
    description of the step that failed, suitable for showing to a user or writing to a log.
    The description never contains the password; the server's own reply is included because it
    says what is wrong (a wrong or deactivated key, for example).
    """
    missing = [
        name
        for name, value in (
            ("SMTP_HOST", smtp_host),
            ("SMTP_USER", smtp_user),
            ("SMTP_PASSWORD", smtp_password),
        )
        if not value
    ]
    if missing:
        return "the email settings on the server are incomplete (missing " + ", ".join(missing) + ")"
    where = f"{smtp_host}:{int(smtp_port)}"
    try:
        server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=timeout)
    except (OSError, smtplib.SMTPException) as exc:  # socket.timeout and gaierror are OSErrors
        return f"could not reach the email server {where} ({_describe(exc)})"
    try:
        with server:
            try:
                server.starttls(context=ssl.create_default_context())
            except (OSError, smtplib.SMTPException) as exc:
                return f"the email server {where} did not accept an encrypted connection ({_describe(exc)})"
            try:
                server.login(smtp_user, smtp_password)
            except smtplib.SMTPAuthenticationError as exc:
                return (
                    "the email server rejected the sign-in; the SMTP key may be wrong, expired, "
                    f"or deactivated ({_describe(exc)})"
                )
            except (OSError, smtplib.SMTPException) as exc:
                return f"signing in to the email server failed ({_describe(exc)})"
    except (OSError, smtplib.SMTPException) as exc:  # disconnect problems after a good sign-in
        return f"the email server {where} dropped the connection ({_describe(exc)})"
    return None


def _describe(exc: BaseException) -> str:
    """A compact rendering of an exception: its type and the first line of its message."""
    if isinstance(exc, socket.timeout):
        return "timed out"
    text = str(exc).strip().splitlines()[0] if str(exc).strip() else ""
    return f"{type(exc).__name__}: {text[:160]}" if text else type(exc).__name__


# Result of the last check per server and login, so that confirming an address a second time
# (or several people confirming within a few minutes) does not sign in to the server again.
_CHECK_MEMO: dict[tuple[str, int, str], tuple[float, str | None]] = {}


def cached_smtp_check(
    *,
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    max_age: float = 300.0,
    timeout: int = 15,
    now=time.monotonic,
) -> str | None:
    """``check_smtp_login`` with its result remembered for ``max_age`` seconds per server and login.

    A failure is remembered too, so a server that is down is not hammered by retries; a user who
    wants to try again sooner can still do so after the memo expires.
    """
    key = (smtp_host, int(smtp_port), smtp_user)
    current = now()
    hit = _CHECK_MEMO.get(key)
    if hit is not None and current - hit[0] < max_age:
        return hit[1]
    result = check_smtp_login(
        smtp_host=smtp_host, smtp_port=smtp_port, smtp_user=smtp_user,
        smtp_password=smtp_password, timeout=timeout,
    )
    _CHECK_MEMO[key] = (current, result)
    return result


def forget_smtp_checks() -> None:
    """Drop remembered check results (tests, or after the settings change)."""
    _CHECK_MEMO.clear()
