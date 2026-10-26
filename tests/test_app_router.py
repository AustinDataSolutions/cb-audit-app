"""The router (app.py) and the usage events the audit screen records.

Router tests drive app.py with AppTest on a seeded database: sign-in screen, the audit screen as the
default page, About without signing in, Usage statistics for managers only, logout, the "leave this
page?" dialog, and the last-run status line across a page visit. Event tests run the audit screen
through the smoke wrapper with the real gate and a signed-in session, so every declared name is
exercised at the place in the script that logs it (the package rejects undeclared names); the same
runs check what the status line says after a finished run.
"""
import io
import time
from pathlib import Path

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest
from streamlit.util import calc_md5

from ads_auth import db
from ads_auth.store import UserStore
from ads_auth.ui.gate import PENDING_PAGE, SESSION_AUTHENTICATED_UTC, SESSION_USER_ID, SESSION_VERSION
from tests.conftest import _make_workbook_bytes

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


# --- leaving the audit screen with inputs to lose -------------------------------------------------

LEAVE_MESSAGE = ("Leaving this page will clear all inputs, including any uploaded files and custom prompts. "
                 "In-progress audits will continue, and completed audits will continue to be available for download.")
LEAVE_QUESTION = "**Do you still wish to proceed?**"
MODEL_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<model name="Dining Model" description="Feedback about dining">
  <tree><node id="root" name="Dining Model"><node id="n1" name="Food" description="About food"/></node></tree>
</model>
"""
AUDIT_XLSX = _make_workbook_bytes({"Sentences": [["#", "Sentences", "Category"], [1, "The food was great", "Food"]]})


class _Upload(io.BytesIO):
    def __init__(self, data, name):
        super().__init__(data)
        self.name = name


def _fake_uploads(monkeypatch, audit=None, model_tree=None):
    """Replace the two file uploaders for this test: each returns the given upload (or nothing) and,
    like the real widget, keeps it under its key in session state."""
    def uploader(label, *args, **kwargs):
        upload = audit if "audit file" in label else model_tree if "model tree" in label else None
        if kwargs.get("key"):
            st.session_state[kwargs["key"]] = upload
        return upload
    monkeypatch.setattr(st, "file_uploader", uploader)


def _audit_screen(monkeypatch, audit=None, model_tree=None):
    import anthropic
    monkeypatch.setattr(anthropic, "Anthropic", lambda *a, **k: _FakeClient())
    _fake_uploads(monkeypatch, audit=audit, model_tree=model_tree)
    return _sign_in(_router(), "alice")


def _sidebar_nav(at):
    """(page links, navigation buttons) the sidebar shows right now."""
    links = [p.label for p in at.sidebar.get("page_link")]
    buttons = [b.key for b in at.sidebar.button if (b.key or "").startswith("auth_nav_")]
    return links, buttons


def _on_audit_screen(at):
    """The audit screen is up (its uploaders are faked in these tests, so look for the Delivery section)."""
    return any(ti.key == "email_results_address" for ti in at.text_input)


def _leave_dialog(at):
    """(message shown, button labels) of the leave dialog, or ([], []) when it is not open."""
    texts = [m.value for m in at.markdown if m.value in (LEAVE_MESSAGE, LEAVE_QUESTION)]
    return texts, [b.label for b in at.button if (b.key or "").startswith("auth_leave_")]


def test_with_nothing_to_lose_the_sidebar_links_navigate_at_once(data, monkeypatch):
    at = _audit_screen(monkeypatch)
    assert _sidebar_nav(at) == (["About this app and data handling"], [])        # alice is not a manager
    assert _leave_dialog(at) == ([], [])
    at = _page(at, "about")
    assert at.markdown[0].value.startswith("# How this app handles your data")


@pytest.mark.parametrize("key", ["audit_prompt", "model_info", "summary_prompt"])
def test_an_edited_prompt_asks_before_leaving_and_go_back_keeps_it(data, monkeypatch, key):
    at = _audit_screen(monkeypatch)
    at.text_area(key=key).input("my own wording").run()
    assert not at.exception, at.exception
    assert _sidebar_nav(at) == ([], ["auth_nav_about"])
    assert [b.label for b in at.sidebar.button if b.key == "auth_nav_about"] == ["About this app and data handling"]

    at.button(key="auth_nav_about").click().run()
    assert not at.exception, at.exception
    assert at.session_state[PENDING_PAGE] == "about"
    assert _leave_dialog(at) == ([LEAVE_MESSAGE, LEAVE_QUESTION], ["Go back", "Proceed"])
    assert _on_audit_screen(at)                                            # still the audit screen, underneath

    # Go back closes the dialog and stays on the audit screen. In a browser the edited text is kept
    # too (checked by hand on Streamlit 1.52.2); AppTest resets keyed widgets after the full rerun that
    # Go back asks for from inside the dialog, so the kept text is not asserted here.
    at.button(key="auth_leave_back").click().run()
    assert not at.exception, at.exception
    assert PENDING_PAGE not in at.session_state
    at.run()
    assert _leave_dialog(at) == ([], [])
    assert _on_audit_screen(at)


def test_proceed_opens_the_chosen_page(data, monkeypatch):
    at = _audit_screen(monkeypatch)
    at.text_area(key="audit_prompt").input("my own wording").run()
    at.button(key="auth_nav_about").click().run()
    at.button(key="auth_leave_proceed").click().run()
    assert not at.exception, at.exception
    assert PENDING_PAGE not in at.session_state
    assert at.markdown[0].value.startswith("# How this app handles your data")
    assert not _on_audit_screen(at)


def test_a_manager_is_asked_before_either_page(data, monkeypatch):
    import anthropic
    monkeypatch.setattr(anthropic, "Anthropic", lambda *a, **k: _FakeClient())
    _fake_uploads(monkeypatch)
    at = _sign_in(_router(), "maria")
    assert _sidebar_nav(at) == (["Usage statistics", "About this app and data handling"], [])
    at.text_area(key="summary_prompt").input("my own wording").run()
    assert _sidebar_nav(at) == ([], ["auth_nav_reports", "auth_nav_about"])
    at.button(key="auth_nav_reports").click().run()
    assert _leave_dialog(at) == ([LEAVE_MESSAGE, LEAVE_QUESTION], ["Go back", "Proceed"])
    at.button(key="auth_leave_proceed").click().run()
    assert not at.exception, at.exception
    assert [t.value for t in at.title] == ["Usage statistics"]


def test_an_uploaded_model_tree_asks_before_leaving(data, monkeypatch):
    at = _audit_screen(monkeypatch, model_tree=_Upload(MODEL_XML, "dining_model.xml"))
    assert _events(data) == [("alice", "login"), ("alice", "upload_model_tree")]
    at.run()        # the faked uploader stores its file during the run; the real widget has it before the run starts
    assert _sidebar_nav(at) == ([], ["auth_nav_about"])
    at.button(key="auth_nav_about").click().run()
    assert _leave_dialog(at) == ([LEAVE_MESSAGE, LEAVE_QUESTION], ["Go back", "Proceed"])


def test_an_uploaded_audit_file_asks_before_leaving(data, monkeypatch):
    at = _audit_screen(monkeypatch, audit=_Upload(AUDIT_XLSX, "dining_audit.xlsx"))
    assert _events(data) == [("alice", "login"), ("alice", "upload_audit_file")]
    at.run()        # see above
    assert _sidebar_nav(at) == ([], ["auth_nav_about"])
    at.button(key="auth_nav_about").click().run()
    assert _leave_dialog(at) == ([LEAVE_MESSAGE, LEAVE_QUESTION], ["Go back", "Proceed"])


# --- the last run's status line -------------------------------------------------------------------

def _status_lines(at):
    return [i.value for i in at.info if i.value.startswith("**Last run**")]


def test_the_last_run_status_line_survives_a_visit_to_another_page(data, monkeypatch):
    at = _audit_screen(monkeypatch)
    assert _status_lines(at) == []
    # What _finalize_job_into_session stores for a run that finished and was emailed (Oct 26, 2026, 2:35 PM Eastern).
    at.session_state["audit_last_run"] = {"finished_at": 1793039700.0, "outcome": "completed", "has_output": True,
                                          "email_status": "sent:alice@example.test"}
    at.session_state["audit_output_bytes"] = AUDIT_XLSX
    at.session_state["audit_output_filename"] = "dining_audit_completed.xlsx"
    at.run()
    assert not at.exception, at.exception
    assert _status_lines(at) == ["**Last run** finished Oct 26, 2026 at 2:35 PM Eastern — audit completed. "
                                 "Results were emailed to alice@example.test."]
    assert [b.label for b in at.get("download_button")] == ["Download completed audit (.xlsx)"]

    at = _page(at, "about")
    assert at.markdown[0].value.startswith("# How this app handles your data")
    assert _status_lines(at) == []
    at = _page(at, "")
    assert _on_audit_screen(at)
    assert _status_lines(at) == ["**Last run** finished Oct 26, 2026 at 2:35 PM Eastern — audit completed. "
                                 "Results were emailed to alice@example.test."]
    assert [b.label for b in at.get("download_button")] == ["Download completed audit (.xlsx)"]


def test_the_status_line_describes_stopped_and_errored_runs(data, monkeypatch):
    at = _audit_screen(monkeypatch)
    at.session_state["audit_last_run"] = {"finished_at": 1793039700.0, "outcome": "stopped", "has_output": True,
                                          "email_status": "failed:mail server unreachable"}
    at.session_state["audit_output_bytes"] = AUDIT_XLSX
    at.session_state["audit_is_partial"] = True
    at.run()
    assert _status_lines(at) == ["**Last run** finished Oct 26, 2026 at 2:35 PM Eastern — audit stopped; partial results "
                                 "are available below. Emailing the results failed (mail server unreachable)."]
    assert [b.label for b in at.get("download_button")] == ["Download audit checkpoint (.xlsx)"]

    at.session_state["audit_last_run"] = {"finished_at": 1793039700.0, "outcome": "errored", "has_output": False,
                                          "email_status": "skipped:no output to send"}
    at.session_state["audit_output_bytes"] = None
    at.run()
    assert _status_lines(at) == ["**Last run** finished Oct 26, 2026 at 2:35 PM Eastern — audit errored; no results file "
                                 "was produced. Results were not emailed (no output to send)."]
    assert not at.get("download_button")


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

    # The persistent status line says when the run finished, that it completed, and where the email went.
    last_run = app.session_state["audit_last_run"]
    assert last_run["outcome"] == "completed" and last_run["email_status"] == "sent:alice@example.test"
    assert abs(last_run["finished_at"] - time.time()) < 60
    status = _status_lines(app)
    assert len(status) == 1 and status[0].endswith("Eastern — audit completed. Results were emailed to alice@example.test.")

    # AppTest cannot click a download button; the wrapper kept each button's click callback.
    callbacks = app.session_state["test_download_callbacks"]
    assert "final_output_download" in callbacks
    on_click, args = callbacks["final_output_download"]
    on_click(*args)
    assert _events(data)[-1] == ("alice", "download")
    assert sorted(e for _, e in _events(data)) == sorted(["upload_audit_file", "upload_model_tree", "run_audit", "email_results", "download"])
    assert {u for u, _ in _events(data)} == {"alice"}

    # On the next run the one-time outcome messages are gone, the status line and download stay.
    app.run()
    assert not app.success
    assert _status_lines(app) == status
    assert "final_output_download" in app.session_state["test_download_callbacks"]

    # Uploading a different file starts over: the output and the status line go with it.
    app.session_state["audit_source_name"] = "an_earlier_file.xlsx"
    app.run()
    assert not app.exception, app.exception
    assert _status_lines(app) == []
    assert "audit_last_run" not in app.session_state and "audit_output_bytes" not in app.session_state


def test_a_run_without_email_and_a_failed_send_record_no_email_event(data, monkeypatch):
    app = _signed_in_script(data, "alice")
    assert _events(data) == [("alice", "upload_audit_file")]       # no model tree this time
    app.checkbox(key="email_results_enabled").uncheck().run()
    _run_audit(app)
    assert [e for _, e in _events(data)] == ["upload_audit_file", "run_audit"]
    assert _status_lines(app)[0].endswith("— audit completed. Email was not requested.")

    monkeypatch.setenv("TEST_EMAIL_FAILS", "1")
    app = _signed_in_script(data, "maria")
    _confirm_email(app, "maria@example.test")
    _run_audit(app)
    assert any("emailing the results failed" in w.value for w in app.warning)
    assert [e for u, e in _events(data) if u == "maria"] == ["upload_audit_file", "run_audit"]
    assert _status_lines(app)[0].endswith("— audit completed. Emailing the results failed (mail server unreachable).")


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
