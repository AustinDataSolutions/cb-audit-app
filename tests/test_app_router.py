"""The router (app.py) and the usage events the audit screen records.

Router tests drive app.py with AppTest on a seeded database: sign-in screen, the audit screen as the
default page, About without signing in, Usage statistics for managers only, logout. Event tests run the
audit screen through the smoke wrapper with the real gate and a signed-in session, so every declared
name is exercised at the place in the script that logs it (the package rejects undeclared names).
"""
import time
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest
from streamlit.util import calc_md5

from ads_auth import db
from ads_auth.store import UserStore
from ads_auth.ui.gate import SESSION_AUTHENTICATED_UTC, SESSION_USER_ID, SESSION_VERSION

ROOT = Path(__file__).resolve().parents[1]
WRAPPER = ROOT / "tests" / "smoke" / "audit_page_wrapper.py"
PW = "a-perfectly-adequate-password"
TITLE = "TestOrg Automatic Audit"


@pytest.fixture
def data(monkeypatch, tmp_path):
    """A required-auth environment with two accounts. Runs from an empty directory so the repository's
    local .env is never read. The provider keys are placeholders; the wrapper replaces the client, and
    the router tests never reach the sidebar's model-list request without signing in first."""
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv("RENDER", raising=False)
    monkeypatch.delenv("TEST_EMAIL_FAILS", raising=False)
    monkeypatch.delenv("TEST_EMAIL_CHECK_FAILS", raising=False)
    monkeypatch.setenv("AUTH_MODE", "required")
    monkeypatch.setenv("TEST_FAKE_IO", "1")              # the smoke wrapper's fakes apply only while this is set
    data_dir = tmp_path / "data"
    monkeypatch.setenv("DATA_DIR", str(data_dir))
    monkeypatch.setenv("ORGANIZATION", "TestOrg")
    monkeypatch.setenv("AUDIENCE", "diners")
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-placeholder-not-a-key")
    monkeypatch.setenv("OPENAI_API_KEY", "")
    for key, value in (("SMTP_HOST", "smtp.example.test"), ("SMTP_PORT", "587"), ("SMTP_USER", "apikey"),
                       ("SMTP_PASSWORD", "not-a-real-password"), ("EMAIL_FROM", "audit@example.test")):
        monkeypatch.setenv(key, value)
    db.migrate(data_dir / "app.db")
    store = UserStore(data_dir / "app.db")
    store.create_user("alice", password=PW)
    store.create_user("maria", password=PW, role="manager")
    return data_dir


def _events(data_dir):
    with db.connection(data_dir / "app.db") as conn:
        return [tuple(row) for row in conn.execute("SELECT username, event FROM usage_events ORDER BY id")]


def _router():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120)
    at.run()
    assert not at.exception, at.exception
    return at


def _sign_in(at, username):
    at.text_input(key="login_username").input(username)
    at.text_input(key="login_password").input(PW)
    at.button[0].click().run()
    assert not at.exception, at.exception
    return at


def _page(at, url_path):
    at._page_hash = calc_md5(url_path) if url_path else ""      # function pages hash their url_path; "" is the default page
    at.run()
    assert not at.exception, at.exception
    return at


def test_signed_out_visitor_sees_the_sign_in_screen_and_no_upload(data):
    at = _router()
    assert [t.value for t in at.title] == [TITLE]
    assert [ti.key for ti in at.text_input] == ["login_username", "login_password"]
    assert not at.get("file_uploader")
    assert not at.sidebar.selectbox                      # the LLM settings are behind the gate too
    assert _events(data) == []


def test_sign_in_shows_the_audit_screen_with_the_account_sidebar(data, monkeypatch):
    import anthropic
    monkeypatch.setattr(anthropic, "Anthropic", lambda *a, **k: _FakeClient())
    at = _sign_in(_router(), "alice")
    assert [t.value for t in at.title] == [TITLE]
    assert len(at.get("file_uploader")) == 2                       # audit workbook and model tree
    assert "Signed in as **alice**" in [m.value for m in at.sidebar.markdown]
    assert any(b.key == "logout_button" for b in at.button)
    assert len(at.get("page_link")) >= 1                          # sidebar About link
    assert [s.label for s in at.sidebar.selectbox] == ["Provider", "Model"]
    assert _events(data) == [("alice", "login")]


class _FakeClient:
    class models:
        @staticmethod
        def list():
            return []


def test_wrong_password_is_refused_and_recorded_without_revealing_the_app(data):
    at = _router()
    at.text_input(key="login_username").input("alice")
    at.text_input(key="login_password").input("not the password")
    at.button[0].click().run()
    assert not at.exception, at.exception
    assert any("not accepted" in e.value for e in at.error)
    assert not at.get("file_uploader")
    assert _events(data) == [("alice", "login_failed")]


def test_about_page_is_readable_without_signing_in(data):
    at = _page(_router(), "about")
    assert at.markdown[0].value.startswith("# How this app handles your data")
    assert not at.get("file_uploader")
    assert _events(data) == []


def test_usage_statistics_is_for_managers_only(data):
    at = _page(_sign_in(_router(), "alice"), "reports")
    assert [t.value for t in at.title] == ["Usage statistics"]
    assert any("managers and administrators" in e.value for e in at.error)
    assert not at.dataframe

    at = _page(_sign_in(_router(), "maria"), "reports")
    assert [t.value for t in at.title] == ["Usage statistics"]
    assert not at.error
    assert len(at.dataframe) == 3
    assert "**Run audit**" in at.caption[0].value and "**Email results**" in at.caption[0].value
    columns = list(at.dataframe[2].value.columns)
    assert columns == ["User", "Logins", "Upload audit files", "Upload model trees", "Run audits", "Email results", "Downloads"]


def test_logout_returns_to_the_sign_in_screen_and_clears_the_session(data, monkeypatch):
    import anthropic
    monkeypatch.setattr(anthropic, "Anthropic", lambda *a, **k: _FakeClient())
    at = _sign_in(_router(), "alice")
    at.session_state["audit_output_bytes"] = b"private session data"
    at.button(key="logout_button").click().run()
    assert not at.exception, at.exception
    assert [ti.key for ti in at.text_input] == ["login_username", "login_password"]
    assert not at.get("file_uploader")
    assert "audit_output_bytes" not in at.session_state and SESSION_USER_ID not in at.session_state
    assert _events(data) == [("alice", "login"), ("alice", "logout")]


def test_disabled_mode_is_refused_on_render(data, monkeypatch):
    monkeypatch.setenv("RENDER", "true")
    monkeypatch.setenv("AUTH_MODE", "disabled")
    at = _router()
    assert any("Refusing to serve" in e.value for e in at.error)
    assert not at.get("file_uploader")


# --- usage events, at the place each one is logged ------------------------------------------------

def _signed_in_script(data_dir, username, **state):
    """The audit screen through the smoke wrapper, with the real gate and `username` already signed in."""
    user = UserStore(data_dir / "app.db").get_by_username(username)
    app = AppTest.from_file(str(WRAPPER), default_timeout=120)
    app.session_state["test_auth_required"] = True
    app.session_state[SESSION_USER_ID] = user.id
    app.session_state[SESSION_VERSION] = user.session_version
    app.session_state[SESSION_AUTHENTICATED_UTC] = db.to_iso(db.utc_now())
    for key, value in state.items():
        app.session_state[key] = value
    app.run()
    assert not app.exception, app.exception
    return app


def _button(app, **match):
    (key, value), = match.items()
    return next(button for button in app.button if getattr(button, key) == value)


def _wait_for_run_to_finish():
    import audit_worker
    registry = audit_worker.get_registry()
    deadline = time.time() + 30
    while registry.active() is not None and time.time() < deadline:
        time.sleep(0.05)
    assert registry.active() is None, "the fake audit run did not finish"


def _confirm_email(app, address):
    app.text_input(key="email_results_address").input(address)
    _button(app, label="Confirm email address").click().run()
    assert not app.exception, app.exception
    assert any("Email address set" in c.value for c in app.caption)


def _run_audit(app):
    """Click Run audit and return once the outcome is on screen. The fake run is instant, so the launch's
    own rerun may already have finalized the job; otherwise one more run after the thread ends does."""
    _button(app, label="Run audit").click().run()
    assert not app.exception, app.exception
    _wait_for_run_to_finish()
    if "active_run_id" in app.session_state:
        app.run()
        assert not app.exception, app.exception
    return app


def test_every_declared_event_is_logged_where_the_action_happens(data):
    import audit_email
    app = _signed_in_script(data, "alice", test_model_tree=True)
    assert "Signed in as **alice**" in [m.value for m in app.sidebar.markdown]
    assert [t.value for t in app.title] == [TITLE]
    assert _events(data) == [("alice", "upload_audit_file"), ("alice", "upload_model_tree")]

    app.run()                                                    # a rerun with the same files is not a new upload
    assert _events(data) == [("alice", "upload_audit_file"), ("alice", "upload_model_tree")]

    _confirm_email(app, "alice@example.test")
    del audit_email.test_sent[:]
    _run_audit(app)
    assert any("Audit complete" in s.value for s in app.success)
    assert any("Results emailed" in s.value for s in app.success)
    assert [m["recipient"] for m in audit_email.test_sent] == ["alice@example.test"]
    # run_audit is logged by the UI after the job starts; the instant fake run may log email_results first.
    assert sorted(e for _, e in _events(data)) == sorted(["upload_audit_file", "upload_model_tree", "run_audit", "email_results"])

    # AppTest cannot click a download button; the wrapper kept each button's click callback.
    callbacks = app.session_state["test_download_callbacks"]
    assert "final_output_download" in callbacks
    on_click, args = callbacks["final_output_download"]
    on_click(*args)
    assert _events(data)[-1] == ("alice", "download")
    assert sorted(e for _, e in _events(data)) == sorted(["upload_audit_file", "upload_model_tree", "run_audit", "email_results", "download"])
    assert {u for u, _ in _events(data)} == {"alice"}


def test_a_run_without_email_and_a_failed_send_record_no_email_event(data, monkeypatch):
    app = _signed_in_script(data, "alice")
    assert _events(data) == [("alice", "upload_audit_file")]       # no model tree this time
    app.checkbox(key="email_results_enabled").uncheck().run()
    _run_audit(app)
    assert [e for _, e in _events(data)] == ["upload_audit_file", "run_audit"]

    monkeypatch.setenv("TEST_EMAIL_FAILS", "1")
    app = _signed_in_script(data, "maria")
    _confirm_email(app, "maria@example.test")
    _run_audit(app)
    assert any("emailing the results failed" in w.value for w in app.warning)
    assert [e for u, e in _events(data) if u == "maria"] == ["upload_audit_file", "run_audit"]


def test_an_address_is_not_confirmed_while_the_email_server_rejects_the_sign_in(data, monkeypatch):
    monkeypatch.setenv("TEST_EMAIL_CHECK_FAILS", "1")
    app = _signed_in_script(data, "alice")
    app.text_input(key="email_results_address").input("alice@example.test")
    _button(app, label="Confirm email address").click().run()
    assert not app.exception, app.exception
    assert any("Email sending isn't working right now: the email server rejected the sign-in (test)" in e.value
               for e in app.error)
    assert not any("Email address set" in c.value for c in app.caption)
    assert "email_confirmed_address" not in app.session_state
    # Running without email is still open to the user.
    app.checkbox(key="email_results_enabled").uncheck().run()
    _run_audit(app)
    assert [e for _, e in _events(data)] == ["upload_audit_file", "run_audit"]


def test_disabled_mode_records_nothing(data, monkeypatch):
    monkeypatch.setenv("AUTH_MODE", "disabled")
    app = AppTest.from_file(str(WRAPPER), default_timeout=120)
    app.session_state["test_model_tree"] = True
    app.run()
    assert not app.exception, app.exception
    assert any("authentication is disabled" in w.value for w in app.warning)
    app.checkbox(key="email_results_enabled").uncheck().run()
    _run_audit(app)
    assert _events(data) == []
