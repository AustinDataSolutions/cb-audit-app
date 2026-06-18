"""Background execution of the audit + summary pipeline.

The audit and its follow-on summary are long-running (20-30 min). Running them
synchronously on Streamlit's script-runner thread ties their lifetime to the
browser websocket: a backgrounded tab, network blip, or proxy idle-timeout
triggers a reconnect, Streamlit starts a fresh script run, and the in-flight
work is abandoned mid-flight.

This module moves the work onto a daemon thread owned by a **process-global
registry** (``JobRegistry``), so the run survives reconnects and Streamlit
reruns — any new script run re-attaches to the same job by id. The worker is
deliberately **Streamlit-free**: it never touches ``st.*``, ``st.session_state``
or ``st.secrets`` (those are unsafe off the main thread and recreated on
reconnect). It only mutates a thread-safe :class:`AuditJob`; the UI polls that
object and does all drawing on the main thread.

What this does NOT survive: process death (container restart/redeploy, OOM,
Community-Cloud sleep with no viewer) kills the daemon thread. That is backstopped
by the workbook checkpoints (resume via ``detect_partial_audit``) and the
email-on-finish sent from here, which lands even with no browser attached.
"""
from __future__ import annotations

import enum
import importlib.util
import logging
import os
import threading
import uuid
from dataclasses import dataclass, field
from typing import Any, Callable, Optional

from audit import run_audit, AuditStopRequested, _is_retryable_llm_error

logger = logging.getLogger(__name__)


class JobStatus(enum.Enum):
    PENDING = "pending"
    RUNNING_AUDIT = "running_audit"
    RUNNING_SUMMARY = "running_summary"
    DONE = "done"
    STOPPED = "stopped"
    ERROR = "error"


_TERMINAL = {JobStatus.DONE, JobStatus.STOPPED, JobStatus.ERROR}
_MAX_LOG_LINES = 200


# ---------------------------------------------------------------------------
# Lazy loaders for the real run/summarize functions
# ---------------------------------------------------------------------------
# The summarizer lives in a hyphenated filename and can't be a normal import;
# load it the same way audit-app.py does. Resolving lazily keeps this module
# importable (and unit-testable with injected fakes) without pulling the
# summarizer in at import time.

_summarizer_module = None


def _load_summarizer_module():
    # Cached: re-execing the module would create a *new* SummaryStopRequested
    # class each call, so the `except` in _run_pipeline would never match the
    # instance raised by summarize_audit_report. Load once, reuse the classes.
    global _summarizer_module
    if _summarizer_module is None:
        module_path = os.path.join(os.path.dirname(__file__), "audit-report-summarizer.py")
        spec = importlib.util.spec_from_file_location("audit_report_summarizer", module_path)
        if spec is None or spec.loader is None:
            raise ImportError("Unable to load audit-report-summarizer module.")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        _summarizer_module = module
    return _summarizer_module


def _default_summarize_fn(**kwargs):
    return _load_summarizer_module().summarize_audit_report(**kwargs)


def _summary_stop_exception():
    """The summarizer's stop exception class (loaded lazily)."""
    return _load_summarizer_module().SummaryStopRequested


# ---------------------------------------------------------------------------
# Email payload + parameters frozen on the main thread at launch
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class EmailPayload:
    """Email settings resolved on the main thread (from st.secrets/session).

    Frozen and passed into the worker so delivery never reads Streamlit state
    off-thread and fires even when no browser is attached at finish.
    """
    enabled: bool
    recipient: str
    smtp: dict  # host/port/user/password/sender — plain resolved values


@dataclass(frozen=True)
class JobParams:
    """Everything the pipeline needs, captured on the main thread at launch."""
    audit_kwargs: dict  # kwargs for run_audit (minus callbacks)
    summary_kwargs: dict  # kwargs for summarize_audit_report (minus input/callbacks)
    completed_filename: str
    checkpoint_filename: str
    include_summary: bool
    email: EmailPayload
    # Summary-only run (re-uploaded completed audit): skip the audit phase and
    # summarize summary_only_input directly.
    summary_only: bool = False
    summary_only_input: Optional[bytes] = None


@dataclass(frozen=True)
class JobSnapshot:
    """Consistent point-in-time view of an AuditJob for the UI to render."""
    run_id: str
    status: JobStatus
    progress_current: int
    progress_total: int
    progress_category: str
    status_message: str
    checkpoint_bytes: Optional[bytes]
    checkpoint_topic: Optional[int]
    output_bytes: Optional[bytes]
    output_filename: Optional[str]
    is_partial: bool
    error_type: Optional[str]
    error_message: Optional[str]
    error_retryable: bool
    summary_error: Optional[str]
    warnings: tuple
    email_status: Optional[str]


# ---------------------------------------------------------------------------
# The shared job object
# ---------------------------------------------------------------------------

class AuditJob:
    """Thread-safe state for one audit run.

    Single-writer-per-field by convention: the worker thread writes everything
    except ``stop_event`` (written by the UI). The lock exists to make
    ``snapshot()`` atomic and to publish ``bytes``/object writes safely across
    threads.
    """

    def __init__(self, run_id: str, params: JobParams):
        self.run_id = run_id
        self.params = params
        self._lock = threading.Lock()
        self.stop_event = threading.Event()  # UI-written; threading.Event is thread-safe
        self.thread: Optional[threading.Thread] = None

        # worker-written, UI-read (guarded by _lock)
        self.status = JobStatus.PENDING
        self.progress_current = 0
        self.progress_total = 0
        self.progress_category = ""
        self.status_message = ""
        self.checkpoint_bytes: Optional[bytes] = None
        self.checkpoint_topic: Optional[int] = None
        self.output_bytes: Optional[bytes] = None
        self.output_filename: Optional[str] = None
        self.is_partial = False
        self.error_type: Optional[str] = None
        self.error_message: Optional[str] = None
        self.error_retryable: bool = False
        self.summary_error: Optional[str] = None
        self.warnings: list = []
        self.log_lines: list = []
        self.email_status: Optional[str] = None

    # -- worker-side mutators -------------------------------------------------
    def set_status(self, status: JobStatus):
        with self._lock:
            self.status = status

    def set_status_message(self, message: str):
        with self._lock:
            self.status_message = message or ""

    def set_progress(self, current: int, total: int, category: str):
        with self._lock:
            self.progress_current = current
            self.progress_total = total
            self.progress_category = category

    def set_checkpoint(self, partial_bytes: bytes):
        with self._lock:
            self.checkpoint_bytes = partial_bytes
            # The checkpoint covers whatever topic we're currently on, matching
            # the old _save_progress semantics (progress_fn fires before each
            # category, save_progress_fn at the checkpoint).
            self.checkpoint_topic = self.progress_current

    def append_warning(self, message: str):
        with self._lock:
            self.warnings.append(str(message))

    def append_log(self, message: str):
        with self._lock:
            self.log_lines.append(str(message))
            if len(self.log_lines) > _MAX_LOG_LINES:
                del self.log_lines[: len(self.log_lines) - _MAX_LOG_LINES]

    def set_output(self, output_bytes: bytes, filename: str, is_partial: bool):
        with self._lock:
            self.output_bytes = output_bytes
            self.output_filename = filename
            self.is_partial = is_partial

    def set_done(self):
        with self._lock:
            self.status = JobStatus.DONE

    def set_stopped(self):
        """Promote the latest checkpoint to the (partial) output and mark stopped."""
        with self._lock:
            if self.checkpoint_bytes is not None:
                self.output_bytes = self.checkpoint_bytes
                self.output_filename = self.params.checkpoint_filename
                self.is_partial = True
            self.status = JobStatus.STOPPED

    def set_error(self, exc: BaseException):
        with self._lock:
            self.error_type = type(exc).__name__
            self.error_message = str(exc)
            self.error_retryable = _is_retryable_llm_error(exc)
            # Surface whatever partial progress exists as a downloadable checkpoint.
            if self.output_bytes is None and self.checkpoint_bytes is not None:
                self.output_bytes = self.checkpoint_bytes
                self.output_filename = self.params.checkpoint_filename
                self.is_partial = True
            self.status = JobStatus.ERROR

    def set_summary_error(self, exc: BaseException):
        """Record a failed summary phase without failing the whole run.

        The audit already succeeded and is stored as the output; the summary is
        a best-effort add-on, so a summary failure leaves the completed audit
        intact and the run terminates DONE (with this note surfaced by the UI).
        """
        with self._lock:
            self.summary_error = f"{type(exc).__name__}: {exc}"

    def set_email_status(self, status: str):
        with self._lock:
            self.email_status = status

    # -- UI-side reader -------------------------------------------------------
    def snapshot(self) -> JobSnapshot:
        with self._lock:
            return JobSnapshot(
                run_id=self.run_id,
                status=self.status,
                progress_current=self.progress_current,
                progress_total=self.progress_total,
                progress_category=self.progress_category,
                status_message=self.status_message,
                checkpoint_bytes=self.checkpoint_bytes,
                checkpoint_topic=self.checkpoint_topic,
                output_bytes=self.output_bytes,
                output_filename=self.output_filename,
                is_partial=self.is_partial,
                error_type=self.error_type,
                error_message=self.error_message,
                error_retryable=self.error_retryable,
                summary_error=self.summary_error,
                warnings=tuple(self.warnings),
                email_status=self.email_status,
            )

    def is_terminal(self) -> bool:
        with self._lock:
            return self.status in _TERMINAL


# ---------------------------------------------------------------------------
# Callbacks: bind run_audit / summarizer callbacks to the job (no st.*)
# ---------------------------------------------------------------------------

def make_callbacks(job: AuditJob) -> dict:
    """Build the callback dict run_audit/summarize_audit_report expect.

    Each callback mutates the job instead of drawing, so the same functions run
    unchanged on a background thread. ``check_stop_fn`` reads the stop Event.
    """
    return dict(
        progress_fn=lambda current, total, category: job.set_progress(current, total, category),
        status_fn=lambda message: job.set_status_message(message or ""),
        save_progress_fn=lambda partial_bytes: job.set_checkpoint(partial_bytes),
        log_fn=lambda message: job.append_log(str(message)),
        warn_fn=lambda message: job.append_warning(str(message)),
        check_stop_fn=lambda: job.stop_event.is_set(),
    )


# ---------------------------------------------------------------------------
# Email (best-effort, never raises)
# ---------------------------------------------------------------------------

def _send_email(job: AuditJob, context_label: str):
    """Email the job's current output, recording the outcome on the job.

    Mirrors the old _maybe_email_results semantics but runs on the worker and
    touches no Streamlit state. Best-effort: failures are recorded, not raised.
    """
    import audit_email

    payload = job.params.email
    snap = job.snapshot()
    if not payload.enabled:
        job.set_email_status("skipped:not enabled")
        return
    if not audit_email.is_valid_email(payload.recipient):
        job.set_email_status("skipped:invalid or missing address")
        return
    smtp = payload.smtp or {}
    if not all(smtp.get(k) for k in ("host", "user", "password", "sender")):
        job.set_email_status("skipped:SMTP not configured")
        return
    if not snap.output_bytes:
        job.set_email_status("skipped:no output to send")
        return
    filename = snap.output_filename or "audit.xlsx"
    try:
        audit_email.send_audit_email(
            smtp_host=smtp["host"],
            smtp_port=smtp.get("port", 587),
            smtp_user=smtp["user"],
            smtp_password=smtp["password"],
            sender=smtp["sender"],
            recipient=payload.recipient,
            subject=f"Audit results ({context_label}): {filename}",
            body=(
                f"Your audit run has finished ({context_label}).\n\n"
                f"The workbook is attached: {filename}\n"
            ),
            attachment_bytes=snap.output_bytes,
            attachment_filename=filename,
        )
        job.set_email_status(f"sent:{payload.recipient}")
        logger.info("Audit results emailed to %s (%s)", payload.recipient, context_label)
    except Exception as exc:  # best-effort
        logger.error("Failed to email audit results: %s", exc, exc_info=True)
        job.set_email_status(f"failed:{exc}")


# ---------------------------------------------------------------------------
# The pipeline: audit -> summary -> email, all on the worker thread
# ---------------------------------------------------------------------------

def _run_pipeline(
    job: AuditJob,
    registry: "JobRegistry",
    run_fn: Callable = run_audit,
    summarize_fn: Callable = _default_summarize_fn,
):
    params = job.params
    callbacks = make_callbacks(job)
    try:
        if params.summary_only:
            # Summary-only run (re-uploaded completed audit): skip the audit
            # phase; the uploaded workbook IS the audit and feeds the summary.
            audit_bytes = params.summary_only_input
            job.set_output(audit_bytes, params.completed_filename, is_partial=False)
        else:
            job.set_status(JobStatus.RUNNING_AUDIT)
            audit_bytes = run_fn(**params.audit_kwargs, **callbacks)
            # Tentative final output is the audit workbook; replaced if summary runs.
            job.set_output(audit_bytes, params.completed_filename, is_partial=False)

        if params.include_summary:
            job.set_status(JobStatus.RUNNING_SUMMARY)
            # summarize_audit_report has no save_progress_fn (only run_audit
            # checkpoints), so pass the callback subset it actually accepts.
            summary_callbacks = {
                k: v for k, v in callbacks.items() if k != "save_progress_fn"
            }
            try:
                summary_bytes = summarize_fn(
                    audit_excel_input=audit_bytes,
                    **params.summary_kwargs,
                    **summary_callbacks,
                )
                job.set_output(summary_bytes, params.completed_filename, is_partial=False)
            except _summary_stop_exception():
                # User stopped the summary: keep the audit-only result.
                logger.warning("Summary stopped by user; keeping audit-only output")
            except Exception as summary_exc:
                # Summary is a best-effort add-on; the audit already succeeded
                # and is preserved as the output. Record the failure and finish
                # DONE rather than failing the whole run.
                logger.error(
                    "Summary failed (run %s): %s: %s",
                    job.run_id, type(summary_exc).__name__, summary_exc, exc_info=True,
                )
                job.set_summary_error(summary_exc)

        job.set_done()
        _send_email(
            job,
            "completed" if job.snapshot().summary_error is None
            else "completed — summary could not be generated",
        )
    except AuditStopRequested:
        logger.warning("Audit stopped by user request (run %s)", job.run_id)
        job.set_stopped()
        _send_email(job, "stopped before finishing — partial results")
    except Exception as exc:
        logger.error("Audit pipeline failed (run %s): %s: %s", job.run_id, type(exc).__name__, exc, exc_info=True)
        job.set_error(exc)
        _send_email(job, "stopped on an error — partial checkpoint")
    finally:
        registry.release(job.run_id)


# ---------------------------------------------------------------------------
# Process-global registry — survives reconnects/reruns within the process
# ---------------------------------------------------------------------------

class AlreadyRunning(RuntimeError):
    """Raised when starting a job while another is still active."""


class JobRegistry:
    def __init__(self):
        self._lock = threading.Lock()
        self._jobs: dict[str, AuditJob] = {}
        self._active_id: Optional[str] = None

    def get(self, run_id: Optional[str]) -> Optional[AuditJob]:
        if not run_id:
            return None
        with self._lock:
            return self._jobs.get(run_id)

    def active(self) -> Optional[AuditJob]:
        """The currently-active (non-terminal) job, if any."""
        with self._lock:
            if self._active_id is None:
                return None
            return self._jobs.get(self._active_id)

    def start(
        self,
        params: JobParams,
        run_fn: Callable = run_audit,
        summarize_fn: Callable = _default_summarize_fn,
    ) -> AuditJob:
        """Create + launch a job. Refuses if one is already active."""
        with self._lock:
            if self._active_id is not None:
                existing = self._jobs.get(self._active_id)
                if existing is not None and not existing.is_terminal():
                    raise AlreadyRunning("An audit is already running.")
            run_id = uuid.uuid4().hex
            job = AuditJob(run_id, params)
            self._jobs[run_id] = job
            self._active_id = run_id

        thread = threading.Thread(
            target=_run_pipeline,
            args=(job, self, run_fn, summarize_fn),
            name=f"audit-job-{run_id[:8]}",
            daemon=True,
        )
        job.thread = thread
        thread.start()
        logger.info("Started audit job %s", run_id)
        return job

    def release(self, run_id: str):
        """Clear the active slot once a job reaches a terminal state."""
        with self._lock:
            if self._active_id == run_id:
                self._active_id = None


_REGISTRY = JobRegistry()


def get_registry() -> JobRegistry:
    return _REGISTRY
