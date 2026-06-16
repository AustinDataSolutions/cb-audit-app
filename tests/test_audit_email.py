"""Tests for audit_email — the pure SMTP delivery helper."""
from __future__ import annotations

import sys

import pytest

# audit_email.py is a normal module name (no hyphen), so a plain import works
# once the project root is importable.
sys.path.insert(0, "..")
import audit_email  # noqa: E402


# ---------------------------------------------------------------------------
# Fake SMTP server captured via monkeypatch
# ---------------------------------------------------------------------------

class _FakeSMTP:
    """Records calls so tests can assert what would have been sent."""

    instances = []

    def __init__(self, host, port, timeout=None):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.started_tls = False
        self.login_args = None
        self.sent_message = None
        _FakeSMTP.instances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def starttls(self, context=None):
        self.started_tls = True

    def login(self, user, password):
        self.login_args = (user, password)

    def send_message(self, message):
        self.sent_message = message


@pytest.fixture
def fake_smtp(monkeypatch):
    _FakeSMTP.instances = []
    monkeypatch.setattr(audit_email.smtplib, "SMTP", _FakeSMTP)
    return _FakeSMTP


def _valid_kwargs(**overrides):
    kwargs = dict(
        smtp_host="smtp.example.com",
        smtp_port=587,
        smtp_user="apikey",
        smtp_password="secret",
        sender="from@example.com",
        recipient="to@example.com",
        subject="Audit results",
        body="Here is your audit.",
        attachment_bytes=b"PK\x03\x04fake-xlsx",
        attachment_filename="audit_completed.xlsx",
    )
    kwargs.update(overrides)
    return kwargs


# ---------------------------------------------------------------------------
# is_valid_email
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("addr", ["a@b.co", "first.last@sub.domain.org"])
def test_is_valid_email_accepts_plausible(addr):
    assert audit_email.is_valid_email(addr)


@pytest.mark.parametrize("addr", ["", "  ", "nope", "a@b", "a@@b.com", "a b@c.com"])
def test_is_valid_email_rejects_bad(addr):
    assert not audit_email.is_valid_email(addr)


# ---------------------------------------------------------------------------
# send_audit_email — happy path
# ---------------------------------------------------------------------------

def test_send_audit_email_sends_with_attachment(fake_smtp):
    audit_email.send_audit_email(**_valid_kwargs())

    assert len(fake_smtp.instances) == 1
    server = fake_smtp.instances[0]
    assert server.host == "smtp.example.com"
    assert server.port == 587
    assert server.started_tls is True
    assert server.login_args == ("apikey", "secret")

    msg = server.sent_message
    assert msg["From"] == "from@example.com"
    assert msg["To"] == "to@example.com"
    assert msg["Subject"] == "Audit results"

    attachments = list(msg.iter_attachments())
    assert len(attachments) == 1
    assert attachments[0].get_filename() == "audit_completed.xlsx"
    assert attachments[0].get_payload(decode=True) == b"PK\x03\x04fake-xlsx"


def test_send_audit_email_coerces_string_port(fake_smtp):
    audit_email.send_audit_email(**_valid_kwargs(smtp_port="2525"))
    assert fake_smtp.instances[0].port == 2525


# ---------------------------------------------------------------------------
# send_audit_email — validation failures (never opens a connection)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("field", ["smtp_host", "smtp_user", "smtp_password", "sender"])
def test_missing_config_raises_before_connecting(fake_smtp, field):
    with pytest.raises(audit_email.EmailConfigError):
        audit_email.send_audit_email(**_valid_kwargs(**{field: ""}))
    assert fake_smtp.instances == []


def test_invalid_recipient_raises_before_connecting(fake_smtp):
    with pytest.raises(audit_email.EmailConfigError):
        audit_email.send_audit_email(**_valid_kwargs(recipient="not-an-email"))
    assert fake_smtp.instances == []
