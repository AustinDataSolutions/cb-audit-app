"""scripts/email_keepalive.py: the monthly message that keeps the SMTP key active."""
import importlib.util
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

_spec = importlib.util.spec_from_file_location("email_keepalive", ROOT / "scripts" / "email_keepalive.py")
keepalive = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(keepalive)

ENV = {"RENDER": "true", "SMTP_HOST": "smtp.example.test", "SMTP_PORT": "587", "SMTP_USER": "login@smtp.example.test",
       "SMTP_PASSWORD": "not-a-real-key", "EMAIL_FROM": "no-reply@example.test", "KEEPALIVE_TO": "owner@example.test"}


def test_sends_one_fixed_message_without_attachment_and_reports_it():
    sent, lines = [], []
    code = keepalive.run(ENV, out=lines.append, send=lambda **kw: sent.append(kw), today=date(2026, 11, 1))
    assert code == 0
    (msg,) = sent
    assert msg["smtp_host"] == "smtp.example.test" and msg["smtp_port"] == 587
    assert msg["smtp_user"] == "login@smtp.example.test" and msg["smtp_password"] == "not-a-real-key"
    assert msg["sender"] == "no-reply@example.test" and msg["recipient"] == "owner@example.test"
    assert msg["subject"] == "Automatic Audit: monthly email check (2026-11-01)"
    assert "keeps the email provider's sending key active" in msg["body"]
    assert msg["attachment_bytes"] == b""
    assert lines == ["keepalive: sent from no-reply@example.test to owner@example.test via smtp.example.test:587 "
                     "as login@smtp.example.test"]


def test_a_failed_send_exits_non_zero_with_the_reason_and_no_key():
    def boom(**kw):
        raise keepalive.audit_email.smtplib.SMTPAuthenticationError(535, b"5.7.8 Authentication failed")
    lines = []
    assert keepalive.run(ENV, out=lines.append, send=boom) == 1
    assert lines[0].startswith("keepalive: FAILED sending from no-reply@example.test to owner@example.test")
    assert "SMTPAuthenticationError" in lines[0] and "Authentication failed" in lines[0]
    assert "not-a-real-key" not in lines[0]


def test_missing_recipient_or_settings_sends_nothing():
    lines = []
    never = lambda **kw: (_ for _ in ()).throw(AssertionError("must not send"))  # noqa: E731
    assert keepalive.run({**ENV, "KEEPALIVE_TO": ""}, out=lines.append, send=never) == 2
    assert keepalive.run({**ENV, "KEEPALIVE_TO": "not-an-address"}, out=lines.append, send=never) == 2
    assert keepalive.run({**ENV, "SMTP_PASSWORD": ""}, out=lines.append, send=never) == 2
    assert [l.split(":")[1].strip()[:12] for l in lines] == ["KEEPALIVE_TO", "KEEPALIVE_TO", "the SMTP set"]


def test_the_real_mailer_builds_a_message_with_no_attachment(monkeypatch):
    from tests.test_audit_email import _FakeSMTP
    _FakeSMTP.instances = []
    monkeypatch.setattr(keepalive.audit_email.smtplib, "SMTP", _FakeSMTP)
    assert keepalive.run(ENV, out=lambda *a: None) == 0
    (server,) = _FakeSMTP.instances
    assert server.started_tls and server.login_args == ("login@smtp.example.test", "not-a-real-key")
    assert server.sent_message["To"] == "owner@example.test"
    assert not server.sent_message.is_multipart()
