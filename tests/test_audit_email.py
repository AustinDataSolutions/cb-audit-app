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
        type(self).instances.append(self)

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


# ---------------------------------------------------------------------------
# check_smtp_login / cached_smtp_check — the sign-in check behind "Confirm email address"
# ---------------------------------------------------------------------------

class _ScriptedSMTP(_FakeSMTP):
    """A fake server that fails at a chosen step: 'connect', 'starttls', or 'login'."""

    fail_at = None

    def __init__(self, host, port, timeout=None):
        if _ScriptedSMTP.fail_at == "connect":
            raise OSError(111, "Connection refused")
        super().__init__(host, port, timeout)

    def starttls(self, context=None):
        if _ScriptedSMTP.fail_at == "starttls":
            raise audit_email.smtplib.SMTPNotSupportedError("STARTTLS extension not supported by server.")
        super().starttls(context)

    def login(self, user, password):
        if _ScriptedSMTP.fail_at == "login":
            raise audit_email.smtplib.SMTPAuthenticationError(535, b"5.7.8 Authentication failed")
        super().login(user, password)


@pytest.fixture
def scripted_smtp(monkeypatch):
    _ScriptedSMTP.instances = []
    _ScriptedSMTP.fail_at = None
    monkeypatch.setattr(audit_email.smtplib, "SMTP", _ScriptedSMTP)
    audit_email.forget_smtp_checks()
    yield _ScriptedSMTP
    audit_email.forget_smtp_checks()


_CHECK = dict(smtp_host="smtp.example.com", smtp_port=587, smtp_user="login@smtp.example.com", smtp_password="s3cret")


def test_check_signs_in_without_sending(scripted_smtp):
    assert audit_email.check_smtp_login(**_CHECK) is None
    (server,) = scripted_smtp.instances
    assert server.started_tls and server.login_args == ("login@smtp.example.com", "s3cret")
    assert server.sent_message is None


@pytest.mark.parametrize("step, phrase", [
    ("connect", "could not reach the email server smtp.example.com:587"),
    ("starttls", "did not accept an encrypted connection"),
    ("login", "rejected the sign-in; the SMTP key may be wrong, expired, or deactivated"),
])
def test_check_names_the_failing_step_without_the_password(scripted_smtp, step, phrase):
    scripted_smtp.fail_at = step
    problem = audit_email.check_smtp_login(**_CHECK)
    assert phrase in problem
    assert "s3cret" not in problem


def test_check_reports_missing_settings_before_connecting(scripted_smtp):
    problem = audit_email.check_smtp_login(**{**_CHECK, "smtp_password": "", "smtp_host": ""})
    assert problem == "the email settings on the server are incomplete (missing SMTP_HOST, SMTP_PASSWORD)"
    assert scripted_smtp.instances == []


def test_cached_check_remembers_the_result_until_it_ages_out(scripted_smtp):
    clock = [1000.0]
    kwargs = {**_CHECK, "now": lambda: clock[0], "max_age": 300}
    scripted_smtp.fail_at = "login"
    first = audit_email.cached_smtp_check(**kwargs)
    assert "rejected the sign-in" in first
    scripted_smtp.fail_at = None
    assert audit_email.cached_smtp_check(**kwargs) == first      # remembered: the failure is not retried
    assert len(scripted_smtp.instances) == 1
    clock[0] += 301
    assert audit_email.cached_smtp_check(**kwargs) is None        # aged out: checked again, now fine
    assert len(scripted_smtp.instances) == 2
    # Another server or login is its own entry.
    assert audit_email.cached_smtp_check(**{**kwargs, "smtp_user": "other@smtp.example.com"}) is None
    assert len(scripted_smtp.instances) == 3
