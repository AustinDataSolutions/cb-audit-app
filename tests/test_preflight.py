"""The app's preflight wrapper: the package's checks, then the email sign-in check as a log line."""
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

_spec = importlib.util.spec_from_file_location("app_preflight", ROOT / "scripts" / "preflight.py")
preflight = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(preflight)


def _environ(tmp_path, **extra):
    env = {"AUTH_MODE": "required", "DATA_DIR": str(tmp_path / "data"), "ORGANIZATION": "TestOrg"}
    env.update(extra)
    return env


SMTP = {"SMTP_HOST": "smtp.example.test", "SMTP_PORT": "587", "SMTP_USER": "login@smtp.example.test",
        "SMTP_PASSWORD": "not-a-real-key", "EMAIL_FROM": "audit@example.test"}


def test_email_check_passes_and_names_the_server_and_login(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    seen, lines = [], []
    code = preflight.run(_environ(tmp_path, **SMTP), out=lines.append, check_email=lambda **kw: seen.append(kw) or None)
    assert code == 0
    assert lines[0].startswith("preflight: ok")
    assert lines[1] == "preflight: email delivery ok (smtp.example.test:587 accepted the sign-in as login@smtp.example.test)"
    assert seen == [{"smtp_host": "smtp.example.test", "smtp_port": 587, "smtp_user": "login@smtp.example.test",
                     "smtp_password": "not-a-real-key"}]


def test_email_check_failure_is_a_warning_and_never_refuses_to_serve(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    lines = []
    code = preflight.run(_environ(tmp_path, **SMTP), out=lines.append,
                         check_email=lambda **kw: "the email server rejected the sign-in (535)")
    assert code == 0
    assert lines[1].startswith("preflight: WARNING email delivery check failed: the email server rejected the sign-in (535).")
    assert "not-a-real-key" not in " ".join(lines)


def test_unconfigured_email_is_reported_as_off_without_connecting(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    lines = []
    code = preflight.run(_environ(tmp_path), out=lines.append, check_email=lambda **kw: (_ for _ in ()).throw(AssertionError))
    assert code == 0
    assert lines[1].startswith("preflight: email delivery is off")


def test_a_failed_package_check_skips_the_email_line(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    lines = []
    code = preflight.run(_environ(tmp_path, RENDER="true", AUTH_MODE="disabled", **SMTP), out=lines.append,
                         check_email=lambda **kw: (_ for _ in ()).throw(AssertionError))
    assert code != 0
    assert len(lines) == 1 and "Refusing to start" in lines[0]
