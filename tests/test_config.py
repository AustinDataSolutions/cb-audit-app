"""This app's settings object: the ads_auth contract filled in for the Automatic Audit app, the values it
reads from the environment, and the usage event vocabulary it declares (every name must be accepted by the
package and listed in the disclosure)."""
from pathlib import Path

import pytest

from ads_auth import AppSettings, db
from ads_auth.usage.events import CORE_EVENTS, log_event
from core.config import (APP_SLUG, APP_TITLE, DEFAULT_AUDIENCE, DEFAULT_ORGANIZATION, PAGES, REPO_ROOT,
                         USAGE_DEFINITIONS, USAGE_EVENTS, Settings, load_settings)

EXPECTED_EVENTS = ("upload_audit_file", "upload_model_tree", "run_audit", "email_results", "download")


def test_defaults_without_env():
    s = load_settings({}, dotenv_path=None)
    assert isinstance(s, AppSettings) and isinstance(s, Settings)
    assert s.app_name == APP_SLUG == "automatic-audit"
    assert s.title == APP_TITLE == "Automatic Audit"
    assert s.data_dir == REPO_ROOT / ".data"
    assert s.db_path == REPO_ROOT / ".data" / "app.db"
    assert s.disclosure_path == REPO_ROOT / "DATA_HANDLING.md" and s.disclosure_path.exists()
    assert s.auth_mode == "required" and not s.auth_disabled and not s.is_render
    assert s.usage_exclude_users == ("ads-dev",)
    assert s.definitions == USAGE_DEFINITIONS
    assert s.organization == DEFAULT_ORGANIZATION and s.audience == DEFAULT_AUDIENCE
    assert s.api_key("anthropic") == "" and s.api_key("OPENAI") == ""
    assert not s.email_configured
    assert s.smtp == {"host": "", "port": 587, "user": "", "password": "", "sender": ""}


def test_environment_values(tmp_path):
    env = {
        "DATA_DIR": str(tmp_path / "d"), "RENDER": "true", "AUTH_MODE": "disabled", "USAGE_EXCLUDE_USERS": "Ops, qa-bot",
        "ORGANIZATION": " AARP ", "AUDIENCE": "members", "ANTHROPIC_API_KEY": "a-key", "OPENAI_API_KEY": "o-key",
        "SMTP_HOST": "smtp.example.test", "SMTP_PORT": "2525", "SMTP_USER": "apikey", "SMTP_PASSWORD": "pw",
        "EMAIL_FROM": "audit@example.test",
    }
    s = load_settings(env, dotenv_path=None)
    assert s.data_dir == tmp_path / "d" and not s.data_dir.exists()    # load_settings performs no I/O
    assert s.is_render and s.auth_disabled
    assert s.usage_exclude_users == ("ops", "qa-bot")
    assert s.organization == "AARP" and s.audience == "members"
    assert s.title == "AARP Automatic Audit"
    assert s.api_key("anthropic") == "a-key" and s.api_key("openai") == "o-key" and s.api_key("other") == ""
    assert s.email_configured
    assert s.smtp == {"host": "smtp.example.test", "port": 2525, "user": "apikey", "password": "pw", "sender": "audit@example.test"}


def test_blank_values_fall_back_to_defaults():
    s = load_settings({"ORGANIZATION": "  ", "AUDIENCE": "", "SMTP_PORT": " "}, dotenv_path=None)
    assert s.organization == DEFAULT_ORGANIZATION and s.audience == DEFAULT_AUDIENCE and s.smtp_port == 587


def test_secrets_never_appear_in_repr():
    s = load_settings({"ANTHROPIC_API_KEY": "sk-secret-a", "OPENAI_API_KEY": "sk-secret-o", "SMTP_USER": "mail-user",
                       "SMTP_PASSWORD": "mail-secret"}, dotenv_path=None)
    text = repr(s)
    for secret in ("sk-secret-a", "sk-secret-o", "mail-user", "mail-secret"):
        assert secret not in text


def test_relative_data_dir_is_under_the_repository():
    s = load_settings({"DATA_DIR": "elsewhere"}, dotenv_path=None)
    assert s.data_dir == REPO_ROOT / "elsewhere" and not s.data_dir.exists()


@pytest.mark.parametrize("env", [{"AUTH_MODE": "shared_password"}, {"SMTP_PORT": "mail"}, {"SMTP_PORT": "0"}])
def test_invalid_values_are_refused(env):
    with pytest.raises(ValueError):
        load_settings(env, dotenv_path=None)


def test_local_dotenv_is_loaded_without_overriding_the_environment(tmp_path, monkeypatch):
    dotenv = tmp_path / ".env"
    dotenv.write_text("AUTH_MODE=disabled\nDATA_DIR=from_file\nORGANIZATION=FileCo\n")
    for key in ("RENDER", "AUTH_MODE", "DATA_DIR", "USAGE_EXCLUDE_USERS", "ORGANIZATION"):
        monkeypatch.delenv(key, raising=False)
    monkeypatch.setenv("DATA_DIR", str(tmp_path / "from_env"))
    s = load_settings(dotenv_path=dotenv)
    assert s.auth_disabled and s.data_dir == tmp_path / "from_env" and s.organization == "FileCo"


def test_render_never_loads_dotenv(tmp_path, monkeypatch):
    dotenv = tmp_path / ".env"
    dotenv.write_text("AUTH_MODE=disabled\n")
    for key in ("AUTH_MODE", "DATA_DIR", "USAGE_EXCLUDE_USERS"):
        monkeypatch.delenv(key, raising=False)
    monkeypatch.setenv("RENDER", "true")
    assert load_settings(dotenv_path=dotenv).auth_mode == "required"


def test_page_registry_is_shared_by_every_settings_instance():
    """The router fills the registry; the audit page builds its own Settings and must see the same pages."""
    a = load_settings({}, dotenv_path=None)
    b = load_settings({"AUTH_MODE": "disabled"}, dotenv_path=None)
    assert a.pages is b.pages is PAGES


def test_declared_events_are_this_apps_five_and_all_count_as_activity():
    assert USAGE_EVENTS == EXPECTED_EVENTS
    s = load_settings({}, dotenv_path=None)
    assert s.events == EXPECTED_EVENTS
    assert not set(s.events) & CORE_EVENTS
    assert s.activity_events == EXPECTED_EVENTS
    assert s.all_events == CORE_EVENTS | set(EXPECTED_EVENTS)
    assert s.all_activity_events == {"login", *EXPECTED_EVENTS}


def test_every_declared_event_is_accepted_and_undeclared_names_are_rejected(tmp_path):
    s = load_settings({"DATA_DIR": str(tmp_path)}, dotenv_path=None)
    db.migrate(s.db_path)
    with db.connection(s.db_path) as conn:
        for name in (*CORE_EVENTS, *s.events):
            log_event(conn, name, allowed=s.events)
        for name in ("upload", "search", "process", "summarize", "audit", "email"):
            with pytest.raises(ValueError):
                log_event(conn, name, allowed=s.events)
        assert conn.execute("SELECT COUNT(*) FROM usage_events").fetchone()[0] == len(CORE_EVENTS) + len(s.events)


def test_definitions_explain_every_event_and_promise_no_content():
    for phrase in ("Active user", "Upload audit file", "Upload model tree", "Run audit", "Email results", "Download"):
        assert f"**{phrase}**" in USAGE_DEFINITIONS, phrase
    assert "never recorded" in USAGE_DEFINITIONS
    assert "resumed checkpoint" in USAGE_DEFINITIONS and "summary-only" in USAGE_DEFINITIONS
    assert "actually" in USAGE_DEFINITIONS          # email: counted only when a message is sent


def test_disclosure_lists_every_recorded_action_in_plain_words():
    text = (REPO_ROOT / "DATA_HANDLING.md").read_text(encoding="utf-8")
    row = next(line for line in text.splitlines() if line.startswith("| Usage events"))
    for words in ("sign-in", "failed sign-in", "sign-out", "audit workbook", "model tree", "audit run", "results email", "download"):
        assert words in row, words
    assert "13 months" in row
    for token in ("upload_audit_file", "upload_model_tree", "run_audit", "email_results", "log_event", ".py", "DATA_DIR"):
        assert token not in text, token
    for provider in ("Anthropic", "OpenAI"):
        assert provider in text
    assert "shared password" in text                 # tells former users it is gone
