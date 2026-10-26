#!/usr/bin/env python
"""Startup checks run before Streamlit (see scripts/start.sh). Exit non-zero to refuse to serve.

The checks live in `ads_auth.preflight`; this wrapper supplies the app's settings. After they pass
it also signs in to the email server behind "Email results to me" (without sending anything) and
reports the result as a `preflight: email ...` line, so a broken key shows up in the service log at
every restart. That check never refuses to serve: the audit itself works without email.
Never prints secrets. Safe to run repeatedly.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import audit_email  # noqa: E402
from ads_auth import preflight  # noqa: E402
from core.config import load_settings  # noqa: E402


def run(environ=None, *, out=print, check_email=audit_email.check_smtp_login) -> int:
    environ = os.environ if environ is None else environ
    dotenv_path = None if environ.get("RENDER", "").lower() == "true" else ".env"
    settings_box = []

    def _load():
        settings_box.append(load_settings(environ, dotenv_path=dotenv_path))
        return settings_box[-1]

    code = preflight.run(_load, out=out)
    if code == 0:
        out(email_status_line(settings_box[-1], check_email=check_email))
    return code


def email_status_line(settings, *, check_email=audit_email.check_smtp_login) -> str:
    """One log line describing whether "Email results to me" can sign in to the email server."""
    if not settings.email_configured:
        return "preflight: email delivery is off (SMTP settings not set); \"Email results to me\" is unavailable."
    smtp = settings.smtp
    problem = check_email(
        smtp_host=smtp["host"], smtp_port=smtp["port"], smtp_user=smtp["user"], smtp_password=smtp["password"],
    )
    if problem:
        return (
            f"preflight: WARNING email delivery check failed: {problem}. The app will start, but "
            "\"Email results to me\" will not work until the SMTP settings on the service are fixed."
        )
    return f"preflight: email delivery ok ({smtp['host']}:{smtp['port']} accepted the sign-in as {smtp['user']})"


if __name__ == "__main__":
    sys.exit(run())
