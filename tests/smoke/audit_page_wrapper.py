"""Run the real audit screen (audit_page.py) with fixture uploads, a fake model provider, a fake
background run, and a fake mailer, so tests can exercise every usage event where it is logged.

Nothing leaves the process: the provider client, the background run, and the mailer are replaced
with fakes while the environment variable TEST_FAKE_IO is set (the test fixture sets it; without it
the originals run, so the rest of the suite is unaffected by these process-wide patches). Session-state
flags set by the test: `test_auth_required` (use the real sign-in gate; the test seeds DATA_DIR and
the session itself), `test_no_audit` (no audit upload), `test_model_tree` (also upload a model tree).
Download buttons keep working; their click callbacks are recorded in
`st.session_state["test_download_callbacks"]` because AppTest cannot click them.
"""
import io
import os
import sys
import types
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
# Sign-in off unless a test asks for the real gate.
os.environ["AUTH_MODE"] = "required" if st.session_state.get("test_auth_required") else "disabled"

import anthropic  # noqa: E402
import audit_email  # noqa: E402
import audit_worker  # noqa: E402
from tests.conftest import _make_workbook_bytes  # noqa: E402

AUDIT_BYTES = _make_workbook_bytes({
    "Sentences": [
        ["#", "Sentences", "Category"],
        [1, "The food was great", "Food"],
        [2, "Service was slow", "Service"],
        [3, "Nice ambiance", "Food"],
    ]
})
OUTPUT_BYTES = _make_workbook_bytes({
    "Sentences": [
        ["ID", "Sentence", "Topic", "Audit", "Explanation"],
        [1, "The food was great", "Food", "YES", "Correct"],
        [2, "Service was slow", "Service", "NO", "Wrong topic"],
        [3, "Nice ambiance", "Food", "YES", "Accurate"],
    ],
    "Findings": [
        ["Topic", "Description", "Accuracy", "Issues"],
        ["Food", "About food", 0.66, ""],
        ["Service", "About service", 0.0, ""],
    ],
})
MODEL_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<model name="Dining Model" description="Feedback about dining">
  <tree>
    <node id="root" name="Dining Model">
      <node id="n1" name="Food" description="About food"/>
      <node id="n2" name="Service" description="About service"/>
    </node>
  </tree>
</model>
"""


class Upload(io.BytesIO):
    def __init__(self, data: bytes, name: str):
        super().__init__(data)
        self.name = name


def _file_uploader(label, *args, **kwargs):
    if "audit file" in label:
        return None if st.session_state.get("test_no_audit") else Upload(AUDIT_BYTES, "dining_audit.xlsx")
    if "model tree" in label:
        return Upload(MODEL_XML, "dining_model.xml") if st.session_state.get("test_model_tree") else None
    return None


if not getattr(st.file_uploader, "_test_wrapper", False):
    _file_uploader._test_wrapper = True
    st.file_uploader = _file_uploader

    _real_download_button = st.download_button

    def _download_button(label, *args, **kwargs):
        on_click = kwargs.get("on_click")
        if callable(on_click):
            callbacks = dict(st.session_state.get("test_download_callbacks") or {})
            callbacks[kwargs.get("key") or label] = (on_click, tuple(kwargs.get("args", ())))
            st.session_state["test_download_callbacks"] = callbacks
        return _real_download_button(label, *args, **kwargs)

    _download_button._test_wrapper = True
    st.download_button = _download_button

    def _faking() -> bool:
        return bool(os.environ.get("TEST_FAKE_IO"))

    # The sidebar asks the provider for its model list on every run; answer locally.
    class _FakeAnthropic:
        def __init__(self, *args, **kwargs):
            self.models = types.SimpleNamespace(list=lambda: [types.SimpleNamespace(id="claude-test-model")])

    _real_anthropic = anthropic.Anthropic

    def _anthropic_client(*args, **kwargs):
        return _FakeAnthropic() if _faking() else _real_anthropic(*args, **kwargs)

    anthropic.Anthropic = _anthropic_client

    # The background job runs an instant fake audit and summary instead of calling a model.
    _real_start = audit_worker.JobRegistry.start

    def _fake_run(*, progress_fn, save_progress_fn, **kwargs):
        progress_fn(1, 2, "Food")
        save_progress_fn(OUTPUT_BYTES)
        progress_fn(2, 2, "Service")
        return OUTPUT_BYTES

    def _fake_summarize(**kwargs):
        return OUTPUT_BYTES

    def _start(self, params, **kwargs):
        if _faking():
            kwargs = dict(run_fn=_fake_run, summarize_fn=_fake_summarize)
        return _real_start(self, params, **kwargs)

    audit_worker.JobRegistry.start = _start

    # The mailer records the send (worker thread; no Streamlit state) instead of talking to a server.
    audit_email.test_sent = []
    _real_send = audit_email.send_audit_email

    def _fake_send(**kwargs):
        if not _faking():
            return _real_send(**kwargs)
        if os.environ.get("TEST_EMAIL_FAILS"):
            raise RuntimeError("mail server unreachable")
        audit_email.test_sent.append({k: v for k, v in kwargs.items() if k != "smtp_password"})

    audit_email.send_audit_email = _fake_send

    # The sign-in check behind "Confirm email address" answers locally too (TEST_EMAIL_CHECK_FAILS
    # makes it report a rejected sign-in).
    _real_check = audit_email.cached_smtp_check

    def _fake_check(**kwargs):
        if not _faking():
            return _real_check(**kwargs)
        if os.environ.get("TEST_EMAIL_CHECK_FAILS"):
            return "the email server rejected the sign-in (test)"
        return None

    audit_email.cached_smtp_check = _fake_check

# The fixture is already in the app's own layout, so skip the CX Designer export reformat (tested in
# tests/test_validation.py and the audit engine tests) the way a resumed checkpoint does.
if st.session_state.get("audit_source_name") == "dining_audit.xlsx":
    st.session_state["reformatted_audit_bytes"] = AUDIT_BYTES

page = ROOT / "audit_page.py"
exec(compile(page.read_text(encoding="utf-8"), str(page), "exec"), {"__name__": "__main__", "__file__": str(page)})
