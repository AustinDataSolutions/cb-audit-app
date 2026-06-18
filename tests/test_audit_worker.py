"""Tests for audit_worker.py — background-execution lifecycle.

The worker is Streamlit-free, so these tests drive it directly with injected
fake run/summarize functions (no LLM calls, no browser). Each test uses its own
JobRegistry() instance rather than the module-global singleton to stay isolated.
"""
from __future__ import annotations

import threading

import pytest


def _make_params(worker, *, include_summary=False, email=None):
    return worker.JobParams(
        audit_kwargs={},
        summary_kwargs={},
        completed_filename="out_completed.xlsx",
        checkpoint_filename="out_checkpoint.xlsx",
        include_summary=include_summary,
        email=email or worker.EmailPayload(enabled=False, recipient="", smtp={}),
    )


def _join(job, timeout=5):
    assert job.thread is not None
    job.thread.join(timeout=timeout)
    assert not job.thread.is_alive(), "worker thread did not finish in time"


def test_happy_path(worker_module):
    w = worker_module
    reg = w.JobRegistry()

    def fake_run(*, progress_fn, save_progress_fn, **kw):
        progress_fn(1, 2, "A")
        save_progress_fn(b"PARTIAL")
        progress_fn(2, 2, "B")
        return b"FINAL_AUDIT"

    job = reg.start(_make_params(w), run_fn=fake_run)
    _join(job)
    snap = job.snapshot()
    assert snap.status == w.JobStatus.DONE
    assert snap.output_bytes == b"FINAL_AUDIT"
    assert snap.output_filename == "out_completed.xlsx"
    assert snap.is_partial is False
    assert (snap.progress_current, snap.progress_total) == (2, 2)
    assert reg.active() is None  # released on completion


def test_summary_runs_and_replaces_output(worker_module):
    w = worker_module
    reg = w.JobRegistry()

    def fake_run(*, progress_fn, **kw):
        progress_fn(1, 1, "A")
        return b"AUDIT"

    def fake_sum(*, audit_excel_input, progress_fn, **kw):
        assert audit_excel_input == b"AUDIT"  # audit output is fed to summary
        progress_fn(1, 1, "A")
        return b"SUMMARY"

    job = reg.start(_make_params(w, include_summary=True), run_fn=fake_run, summarize_fn=fake_sum)
    _join(job)
    snap = job.snapshot()
    assert snap.status == w.JobStatus.DONE
    assert snap.output_bytes == b"SUMMARY"  # summary replaces audit bytes
    assert snap.is_partial is False


def test_summary_gets_only_accepted_callbacks(worker_module):
    """Regression: summarize_audit_report has no save_progress_fn.

    The fake mirrors the real summarizer signature exactly (the callbacks it
    accepts, no save_progress_fn, no **kwargs catch-all), so passing the full
    run_audit callback set would raise TypeError — which is what happened in
    production before the fix.
    """
    w = worker_module
    reg = w.JobRegistry()

    def fake_run(*, progress_fn, save_progress_fn, **kw):
        progress_fn(1, 1, "A")
        return b"AUDIT"

    def fake_sum(
        *,
        audit_excel_input,
        msg_template=None,
        llm_provider="anthropic",
        model_name=None,
        max_tokens=0,
        accuracy_threshold=0.8,
        model_info="",
        anthropic_api_key=None,
        openai_api_key=None,
        log_fn=None,
        warn_fn=None,
        progress_fn=None,
        check_stop_fn=None,
        status_fn=None,
    ):
        return b"SUMMARY"

    job = reg.start(_make_params(w, include_summary=True), run_fn=fake_run, summarize_fn=fake_sum)
    _join(job)
    snap = job.snapshot()
    assert snap.status == w.JobStatus.DONE, snap.error_message
    assert snap.output_bytes == b"SUMMARY"


def test_summary_stop_keeps_audit_output(worker_module):
    w = worker_module
    reg = w.JobRegistry()

    def fake_run(**kw):
        return b"AUDIT"

    def fake_sum(**kw):
        raise w._summary_stop_exception()("stop summary")

    job = reg.start(_make_params(w, include_summary=True), run_fn=fake_run, summarize_fn=fake_sum)
    _join(job)
    snap = job.snapshot()
    # Summary stopped → keep the audit-only result, run still counts as DONE.
    assert snap.status == w.JobStatus.DONE
    assert snap.output_bytes == b"AUDIT"
    assert snap.is_partial is False


def test_stop_promotes_checkpoint(worker_module):
    w = worker_module
    reg = w.JobRegistry()
    reached = threading.Event()
    release = threading.Event()

    def fake_run(*, progress_fn, save_progress_fn, check_stop_fn, **kw):
        progress_fn(1, 5, "A")
        save_progress_fn(b"CKPT")
        reached.set()
        release.wait(timeout=5)
        if check_stop_fn():
            raise w.AuditStopRequested("stopped")
        return b"SHOULD_NOT_REACH"

    job = reg.start(_make_params(w), run_fn=fake_run)
    assert reached.wait(timeout=5)
    job.stop_event.set()
    release.set()
    _join(job)
    snap = job.snapshot()
    assert snap.status == w.JobStatus.STOPPED
    assert snap.output_bytes == b"CKPT"  # checkpoint promoted as partial output
    assert snap.output_filename == "out_checkpoint.xlsx"
    assert snap.is_partial is True
    assert snap.checkpoint_topic == 1
    assert reg.active() is None


def test_error_promotes_checkpoint(worker_module):
    w = worker_module
    reg = w.JobRegistry()

    def fake_run(*, progress_fn, save_progress_fn, **kw):
        progress_fn(1, 3, "A")
        save_progress_fn(b"CKPT")
        raise RuntimeError("boom")

    job = reg.start(_make_params(w), run_fn=fake_run)
    _join(job)
    snap = job.snapshot()
    assert snap.status == w.JobStatus.ERROR
    assert snap.error_type == "RuntimeError"
    assert "boom" in (snap.error_message or "")
    assert snap.output_bytes == b"CKPT"
    assert snap.is_partial is True
    assert reg.active() is None


def test_one_job_at_a_time(worker_module):
    w = worker_module
    reg = w.JobRegistry()
    started = threading.Event()
    release = threading.Event()

    def gated_run(**kw):
        started.set()
        release.wait(timeout=5)
        return b"FINAL"

    job1 = reg.start(_make_params(w), run_fn=gated_run)
    assert started.wait(timeout=5)
    with pytest.raises(w.AlreadyRunning):
        reg.start(_make_params(w), run_fn=gated_run)
    release.set()
    _join(job1)
    # Slot freed → a fresh start now succeeds.
    job2 = reg.start(_make_params(w), run_fn=lambda **kw: b"FINAL2")
    _join(job2)
    assert job2.snapshot().output_bytes == b"FINAL2"


def test_active_reattach(worker_module):
    w = worker_module
    reg = w.JobRegistry()
    started = threading.Event()
    release = threading.Event()

    def gated_run(**kw):
        started.set()
        release.wait(timeout=5)
        return b"FINAL"

    job = reg.start(_make_params(w), run_fn=gated_run)
    assert started.wait(timeout=5)
    # A fresh script run / session re-finds the live job via the registry.
    assert reg.active() is job
    assert reg.get(job.run_id) is job
    release.set()
    _join(job)
    assert reg.active() is None


def test_email_sent_on_completion(worker_module, monkeypatch):
    w = worker_module
    reg = w.JobRegistry()
    import audit_email

    calls = {}

    def fake_send(**kwargs):
        calls.update(kwargs)

    monkeypatch.setattr(audit_email, "send_audit_email", fake_send)
    email = w.EmailPayload(
        enabled=True,
        recipient="a@b.com",
        smtp={"host": "h", "port": 25, "user": "u", "password": "p", "sender": "s@b.com"},
    )
    job = reg.start(_make_params(w, email=email), run_fn=lambda **kw: b"FINAL")
    _join(job)
    snap = job.snapshot()
    assert snap.email_status == "sent:a@b.com"
    assert calls["recipient"] == "a@b.com"
    assert calls["smtp_host"] == "h"
    assert calls["attachment_bytes"] == b"FINAL"


def test_email_skipped_when_disabled(worker_module):
    w = worker_module
    reg = w.JobRegistry()
    job = reg.start(_make_params(w), run_fn=lambda **kw: b"FINAL")
    _join(job)
    assert job.snapshot().email_status == "skipped:not enabled"


def test_snapshot_is_threadsafe_smoke(worker_module):
    """Hammer set_progress from a thread while snapshotting; must not raise."""
    w = worker_module
    job = w.AuditJob("rid", _make_params(w))
    stop = threading.Event()

    def writer():
        i = 0
        while not stop.is_set():
            i += 1
            job.set_progress(i, 1000, "cat")

    t = threading.Thread(target=writer, daemon=True)
    t.start()
    try:
        for _ in range(2000):
            snap = job.snapshot()
            assert snap.progress_total == 1000
    finally:
        stop.set()
        t.join(timeout=5)
