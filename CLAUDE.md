# CLAUDE.md

## Project overview
Streamlit app for auditing Qualtrics/CX Designer topic models using LLM-assisted
accuracy checks. The app ingests a Qualtrics audit export (.xlsx), optionally
enriches it with an XML model tree, and runs category-level audits using an LLM
(Anthropic or OpenAI). It produces a completed audit workbook with per-topic
accuracy metrics and can optionally generate summaries of issues for low-accuracy
topics. Hosted on Render behind individual accounts; `audit-app.py` is the
standalone "moved" notice served by the legacy Community Cloud deploy and must
import nothing from the app.

### User workflow
1. Sign in (accounts come from the shared `ads_auth` package)
2. Upload audit `.xlsx` exported from Qualtrics/CX Designer
3. (Optional) Upload model tree `.xml` for topic selection and descriptions
4. Choose topics, adjust prompt, set LLM provider/model/limits in sidebar
5. Run audit — download completed workbook when done, or have it emailed
6. (Optional) Generate summaries for low-accuracy topics

### Configuration
Everything is an environment variable (`core/config.py` documents the names). On
Render they are set on the service; locally a git-ignored `.env` is loaded
(`.env.example` lists them). Streamlit secrets are not used; there is no shared
password. `config.yaml` holds default provider/model and limits; `prompts.yaml`
holds the audit and summary prompt templates. The model names in `config.yaml`
are fallbacks: the sidebar preselects the newest model of that family from the
provider's live list (`core/model_choice.py`), so bumping them is only needed
when a family is renamed or retired.

## Governing documents
- `DATA_HANDLING.md`: the disclosure users read inside the app (About page). It
  is user-facing copy only: plain language, no instructions to agents or
  maintainers, no file or code names. Keep it true.
- `docs/render-deployment-runbook.md`: creating and operating the Render service,
  written for the app's owner. `docs/OPERATIONS.md`: operator commands.

## Rules
- **Accounts, sessions, usage events, backups, the sign-in gate, the Usage
  statistics and About pages, preflight, and the operator CLI come from the
  `ads-auth` package** (checkout `~/dev/aarp-2026/ads-auth`, pinned by tag in
  `requirements.txt`). This app keeps `core/config.py` (its `Settings` extends
  `ads_auth.AppSettings` and declares the event vocabulary), `app.py` (the
  router), `manage.py`, `scripts/`, and `DATA_HANDLING.md`. Fix shared behavior
  in the package and bump the pin; do not re-implement it here. Bumping the pin
  to a release whose CHANGELOG entry is marked "Disclosure: affected" requires a
  `DATA_HANDLING.md` update in the same commit.
  The package's stay-signed-in cookie is on whenever a cookie secret is present
  (`ADS_AUTH_COOKIE_SECRET`, else `STREAMLIT_SERVER_COOKIE_SECRET`, which Render sets).
- **Disclosure stays current.** Any change that stores, logs, transmits, or
  retains data differently (new table, new file on disk, new third-party call,
  new prompt content, new retention window, new cookie, new usage event name)
  must update `DATA_HANDLING.md` in the same commit, including its "Last
  updated" date. The About page renders that file verbatim. Treat a stale
  disclosure as a failing test (`tests/test_config.py` checks the event list).
- **Usage events** are declared in `core/config.py` (`USAGE_EVENTS`); `log_event`
  rejects undeclared names. Log at the place the action happens: uploads when a
  new file is read, `run_audit` when a job is started, `email_results` from the
  worker after a successful send (`JobParams.on_email_sent`), `download` as the
  button's `on_click`. Events hold timestamp, account, and event type only.
- Render deploys manually (`autoDeployTrigger: off`, Blueprint Auto Sync off).
  Pushing `main` also redeploys the Community Cloud notice from `audit-app.py`
  until that app is deleted.
- Never commit `.env`, `.data/`, `*.db`, real audit exports, credentials, or
  password hashes. Test fixtures are synthetic.
- Never log or print credentials, uploaded content, sentences, prompts, model
  responses, email addresses, or usernames with content.
- `core/` imports no Streamlit and does no I/O at import time.
- Run the full test suite before claiming a change works: `venv/bin/python -m pytest`.
- Python 3.12.3 (`.python-version`), Streamlit pinned at 1.52.2 until a
  deliberate upgrade.

## Development setup
```bash
python -m venv venv
venv/bin/python -m pip install -r requirements.txt
cp .env.example .env   # then fill in the values; AUTH_MODE=disabled skips sign-in locally
```

## Running the app
```bash
venv/bin/python -m streamlit run app.py
```
With `AUTH_MODE=required` locally, create an account first:
`venv/bin/python manage.py create-user <username>`.

## CLI entry points (non-UI)
```bash
python manage.py --help                                  # accounts, backups, usage events
python audit.py                                          # run audit from config.yaml + inputs/
python audit-report-summarizer.py --input path/to.xlsx   # summarize existing audit (in-place unless --output)
```

## Testing
```bash
venv/bin/python -m pytest -q            # fast summary
venv/bin/python -m pytest -v            # verbose with test names
venv/bin/python -m pytest tests/test_audit.py  # single file
```
Tests use programmatically-generated Excel fixtures (no external files needed).
`tests/test_app_router.py` drives `app.py` and, through
`tests/smoke/audit_page_wrapper.py`, the real audit screen with fixture uploads,
a fake provider client, an instant fake run, and a fake mailer.

## Key architecture notes

### File naming
- `audit-report-summarizer.py` has a hyphenated filename. Import it with
  `importlib.util` (see `tests/conftest.py` for the pattern and
  `audit_worker._load_summarizer_module()`).

### Module roles
| File | Role |
|------|------|
| `app.py` | Router: `st.navigation` over the audit screen (file page) and the package's About and Usage statistics pages; prepares the database and the housekeeping thread once per process |
| `audit_page.py` | Streamlit UI: sign-in gate and account sidebar (from `ads_auth`), launches/polls the background job, renders progress + downloads, logs usage events |
| `core/config.py` | `Settings` (extends `ads_auth.AppSettings`): app name, title, events, provider keys, SMTP values, organization and audience, all from the environment |
| `core/model_choice.py` | The default-model rule: from the provider's model list, the newest model of the same family as the `config.yaml` name (Opus, Terra); the configured name is the fallback when the list is unavailable. Pure functions, tested with fake lists |
| `manage.py`, `scripts/preflight.py` | Wrappers around `ads_auth.cli` and `ads_auth.preflight` with this app's settings |
| `audit_worker.py` | Background execution: `JobRegistry` (process-global), `AuditJob`, `_run_pipeline` (audit→summary→email) |
| `audit.py` | Core audit engine (`run_audit`), LLM calls, retry logic, partial detection |
| `audit-report-summarizer.py` | Post-audit summary generation for low-accuracy topics |
| `audit_validation.py` | Input file format detection and header validation |
| `audit_reformat.py` | Pre-processing to normalize input files to standard format |
| `audit-app.py` | Community Cloud "moved" notice; imports only Streamlit |

### Background-worker execution model
The audit + summary + email run on a daemon thread owned by the process-global
`JobRegistry` in `audit_worker.py`, NOT on Streamlit's script-runner thread.
This lets a run survive browser disconnects and Streamlit reruns: any new script
run re-attaches to the live job via `registry.active()`.
- The worker is **Streamlit-free** — it never touches `st.*` or
  `st.session_state` (unsafe off the main thread, recreated on reconnect). It
  only mutates a thread-safe `AuditJob`; the UI polls it via `snapshot()`.
- The UI's `@st.fragment(run_every="2s")` panel reads the job and redraws; on a
  terminal status it `st.rerun(scope="app")`s out of poll mode and the main run
  promotes the output + shows the outcome.
- Callbacks (`progress_fn`/`status_fn`/`save_progress_fn`/`log_fn`/`warn_fn`/
  `check_stop_fn`) are bound to job methods in `make_callbacks()`, so `run_audit`
  /`summarize_audit_report` run unchanged on the worker. Stop is a
  `threading.Event` on the job.
- Email (SMTP config + recipient) is frozen on the main thread at launch into
  `JobParams.email` and sent from the worker, so it lands even with no browser.
  `JobParams.on_email_sent` (the UI passes a usage-event logger) runs on the
  worker after a successful send and never changes the outcome.
- **Not covered:** process death (container restart/redeploy, OOM) kills the
  daemon thread — backstopped by checkpoint files (resume via
  `detect_partial_audit`) and the worker's email-on-finish.

### LLM call pattern
All LLM calls go through `_call_llm_with_status()` which runs the API call in a
background thread and shows periodic "waiting for response" status updates. Both
`audit.py` and `audit-report-summarizer.py` have their own copy of this helper.
LLM calls have a 300s timeout (`DEFAULT_LLM_TIMEOUT`). Retryable errors (429,
503, 529, timeouts) retry up to 3 times with delays of 5s, 60s, 120s.

### Streamlit session state keys
- `audit_run_requested`: True between the "Run audit" click and the launch
- `active_run_id`: the only handle from the session to the background job. Set
  when this session launches a run. On reconnect it is re-derived from
  `registry.active()` **only if** the `?run=<run_id>` URL query param matches —
  see run ownership below.
- `finalized_run_id`: idempotency guard so a terminal job is promoted once
- `audit_output_bytes` / `audit_output_filename`: final or promoted partial
  output for download
- `audit_is_partial`: whether the current output is incomplete
- `audit_last_run`: when the last run finished, how it ended, and the email
  outcome, written by `_finalize_job_into_session` and shown as a status line
  above the final download button. Session state survives a visit to the About
  or Usage statistics page (widget state does not), so the line is still there
  on return; a new upload clears it with the output.
- `auth_*`: the package's session keys (user id, session version, sign-in time;
  `auth_cookie_action` and `auth_cookie_blocked` drive the stay-signed-in cookie);
  sign-out clears the whole session state, uploads and results included.
- `auth_pending_page`: the package's record of a sidebar page link clicked while
  the screen had inputs to lose (see "Leaving the audit screen").

Live progress, the latest checkpoint bytes, the stop signal, and email status
all live on the `AuditJob` (not session_state). Within `run_audit`, checkpoints
are still written every `CHECKPOINT_INTERVAL` categories, on the last category,
and on retry/stop/error.

### Leaving the audit screen
Streamlit drops widget state (uploads, edited prompts, the typed email address)
when the user opens another page. The audit page passes
`confirm_leave=_has_unsaved_inputs` to the package's `render_sidebar`: while an
audit file or model tree is uploaded (widget keys `audit_upload`,
`model_tree_upload`) or the audit prompt, model notes, or summary prompt differ
from `prompts.yaml` (keys `audit_prompt`, `model_info`, `summary_prompt`), the
sidebar's page links become buttons and `confirm_navigation` opens a "Leave this
page?" dialog with `LEAVE_MESSAGE`, Go back, and Proceed (`st.switch_page`).
Both pieces live in `ads_auth.ui.gate` (v1.1.3); only the check and the wording
are this app's.

### Run ownership (single active run)
The `JobRegistry` is process-global and allows **one active audit at a time**
across the whole deployment (one Render instance = single process). Several
people can be signed in at once, so "who owns the running audit" matters:
- The launching session stamps `?run=<run_id>` into the URL query string and
  sets `active_run_id`. That token survives a websocket reconnect **within the
  same browser tab**, so only the owner's tab re-attaches to the live run (full
  progress panel + Stop + checkpoint download).
- A different session (another account, or the owner after losing the tab/URL)
  does **not** carry the token, so it can't adopt the run. A full-page gate near
  the top of `main()` renders only the title + a read-only "an audit is already
  running" banner/progress (`_busy_gate`, polls every 2s) and `return`s before
  the upload area, LLM settings, or Run button — the rest of the UI is hidden
  until the active run ends. It cannot stop or download the run. The account
  sidebar (Log out, About) is drawn before the gate and stays available.
- Idle sessions render `_collision_heartbeat` (polls every 2s) so the gate
  appears within ~2s of another session starting a run, without needing a click.
- `registry.start()` raises `AlreadyRunning` if a run is active; the launch path
  backs out (does **not** hijack the existing run) and falls through to the busy
  banner. Email delivery still lands the result for whoever configured it.
- **Accepted residual:** ownership is still per browser tab, not per account. A
  reconnecting owner who lost the URL gets the busy banner like anyone else and
  relies on email for the result. Tying ownership to the signed-in account is a
  possible follow-up now that accounts exist.

## Conventions
- Callbacks follow the pattern: `log_fn`, `warn_fn`, `progress_fn`, `status_fn`,
  `save_progress_fn`, `check_stop_fn`. The UI wires these to Streamlit widgets;
  the CLI uses print/no-ops.
- Excel workbooks always have four sheets: Findings, Sentences, Audit Settings,
  Errors.
- The `_is_retryable_llm_error()` function in `audit.py` is the single source of
  truth for which errors trigger retry logic.
