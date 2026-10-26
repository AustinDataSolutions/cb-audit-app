#!/usr/bin/env python
"""Send one short email through the app's SMTP settings so the provider's key stays active.

Brevo deactivates an SMTP key that has sent nothing for 90 days. The Render cron job
`aarp-automatic-audit-email-keepalive` (render.yaml) runs this once a month; it also works by hand:

    python scripts/email_keepalive.py

Reads the same SMTP_* and EMAIL_FROM settings as the app (`core/config.py`; `.env` locally, the
service's environment on Render) plus KEEPALIVE_TO, the address the check is sent to. Sends no
user data: the message is one fixed sentence. Prints one line and exits 0 when the message was
accepted, 1 when sending failed, 2 when a setting is missing. Never prints the key.
"""
from __future__ import annotations

import os
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import audit_email  # noqa: E402
from core.config import load_settings  # noqa: E402

SUBJECT = "Automatic Audit: monthly email check"
BODY = (
    "This is the Automatic Audit app's monthly check that results email still works.\n"
    "It keeps the email provider's sending key active. Nothing to do.\n"
    "If this message stops arriving, see \"Email delivery\" in docs/OPERATIONS.md.\n"
)


def run(environ=None, *, out=print, send=audit_email.send_audit_email, today: date | None = None) -> int:
    environ = os.environ if environ is None else environ
    dotenv_path = None if environ.get("RENDER", "").lower() == "true" else ".env"
    try:
        settings = load_settings(environ, dotenv_path=dotenv_path)
    except ValueError as exc:
        out(f"keepalive: {exc}")
        return 2
    recipient = environ.get("KEEPALIVE_TO", "").strip()
    if not audit_email.is_valid_email(recipient):
        out("keepalive: KEEPALIVE_TO is not set to an email address; nothing sent.")
        return 2
    if not settings.email_configured:
        out("keepalive: the SMTP settings are not all set (SMTP_HOST, SMTP_USER, SMTP_PASSWORD, EMAIL_FROM); nothing sent.")
        return 2
    smtp = settings.smtp
    stamp = (today or date.today()).isoformat()
    try:
        send(
            smtp_host=smtp["host"], smtp_port=smtp["port"], smtp_user=smtp["user"], smtp_password=smtp["password"],
            sender=smtp["sender"], recipient=recipient,
            subject=f"{SUBJECT} ({stamp})", body=BODY,
            attachment_bytes=b"", attachment_filename="",
        )
    except Exception as exc:  # noqa: BLE001 — report every failure the same way, then exit non-zero
        text = str(exc).strip().splitlines()[0] if str(exc).strip() else ""
        detail = f"{type(exc).__name__}: {text[:160]}" if text else type(exc).__name__
        out(f"keepalive: FAILED sending from {smtp['sender']} to {recipient} via {smtp['host']}:{smtp['port']} ({detail})")
        return 1
    out(f"keepalive: sent from {smtp['sender']} to {recipient} via {smtp['host']}:{smtp['port']} as {smtp['user']}")
    return 0


if __name__ == "__main__":
    sys.exit(run())
