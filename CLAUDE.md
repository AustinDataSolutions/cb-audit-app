# CLAUDE.md

## Project overview
Streamlit app for auditing Qualtrics/CX Designer topic models using LLM-assisted
accuracy checks. The app ingests a Qualtrics audit export (.xlsx), optionally
enriches it with an XML model tree, and runs category-level audits using an LLM
(Anthropic or OpenAI). It produces a completed audit workbook with per-topic
accuracy metrics and can optionally generate summaries of issues for low-accuracy
topics.

### User workflow
1. Upload audit `.xlsx` exported from Qualtrics/CX Designer
2. (Optional) Upload model tree `.xml` for topic selection and descriptions
3. Choose topics, adjust prompt, set LLM provider/model/limits in sidebar
4. Run audit — download completed workbook when done
5. (Optional) Generate summaries for low-accuracy topics

### Configuration
- `.streamlit/secrets.toml`: APP_PASSWORD, ANTHROPIC_API_KEY, OPENAI_API_KEY,
  and optional SMTP settings (SMTP_HOST/PORT/USER/PASSWORD, EMAIL_FROM) that
  enable the "Email results to me" delivery option
- `config.yaml`: default provider/model and limits for UI and CLI
- `prompts.yaml`: audit and summary prompt templates

## Development setup
```bash
source venv/bin/activate
pip install -r requirements.txt
```

## Running the app
```bash
streamlit run audit-app.py
```

## CLI entry points (non-UI)
```bash
python audit.py                                          # run audit from config.yaml + inputs/
python audit-report-summarizer.py --input path/to.xlsx   # summarize existing audit (in-place unless --output)
```

## Testing
```bash
pytest -q            # fast summary
pytest -v            # verbose with test names
pytest tests/test_audit.py  # single file
```
Tests use programmatically-generated Excel fixtures (no external files needed).

## Key architecture notes

### File naming
- `audit-report-summarizer.py` has a hyphenated filename. Import it with
  `importlib.util` (see `tests/conftest.py` for the pattern and
  `audit_worker._load_summarizer_module()`).

### Module roles
| File | Role |
|------|------|
| `audit-app.py` | Streamlit UI: launches/polls the background job, renders progress + downloads |
| `audit_worker.py` | Background execution: `JobRegistry` (process-global), `AuditJob`, `_run_pipeline` (audit→summary→email) |
| `audit.py` | Core audit engine (`run_audit`), LLM calls, retry logic, partial detection |
| `audit-report-summarizer.py` | Post-audit summary generation for low-accuracy topics |
| `audit_validation.py` | Input file format detection and header validation |
| `audit_reformat.py` | Pre-processing to normalize input files to standard format |

### Background-worker execution model
The audit + summary + email run on a daemon thread owned by the process-global
`JobRegistry` in `audit_worker.py`, NOT on Streamlit's script-runner thread.
This lets a run survive browser disconnects and Streamlit reruns: any new script
run re-attaches to the live job via `registry.active()`.
- The worker is **Streamlit-free** — it never touches `st.*`, `st.session_state`,
  or `st.secrets` (unsafe off the main thread, recreated on reconnect). It only
  mutates a thread-safe `AuditJob`; the UI polls it via `snapshot()`.
- The UI's `@st.fragment(run_every="2s")` panel reads the job and redraws; on a
  terminal status it `st.rerun(scope="app")`s out of poll mode and the main run
  promotes the output + shows the outcome.
- Callbacks (`progress_fn`/`status_fn`/`save_progress_fn`/`log_fn`/`warn_fn`/
  `check_stop_fn`) are bound to job methods in `make_callbacks()`, so `run_audit`
  /`summarize_audit_report` run unchanged on the worker. Stop is a
  `threading.Event` on the job.
- Email (SMTP config + recipient) is frozen on the main thread at launch into
  `JobParams.email` and sent from the worker, so it lands even with no browser.
- **Not covered:** process death (container restart/redeploy, OOM, Community-Cloud
  sleep with no viewer) kills the daemon thread — backstopped by checkpoint files
  (resume via `detect_partial_audit`) and the worker's email-on-finish.

### LLM call pattern
All LLM calls go through `_call_llm_with_status()` which runs the API call in a
background thread and shows periodic "waiting for response" status updates. Both
`audit.py` and `audit-report-summarizer.py` have their own copy of this helper.
LLM calls have a 300s timeout (`DEFAULT_LLM_TIMEOUT`). Retryable errors (429,
503, 529, timeouts) retry up to 3 times with delays of 30s, 60s, 120s.

### Streamlit session state keys
- `audit_run_requested`: True between the "Run audit" click and the launch
- `active_run_id`: the only handle from the session to the background job;
  re-derived from `registry.active()` on reconnect
- `finalized_run_id`: idempotency guard so a terminal job is promoted once
- `audit_output_bytes` / `audit_output_filename`: final or promoted partial
  output for download
- `audit_is_partial`: whether the current output is incomplete

Live progress, the latest checkpoint bytes, the stop signal, and email status
all live on the `AuditJob` (not session_state). Within `run_audit`, checkpoints
are still written every `CHECKPOINT_INTERVAL` categories, on the last category,
and on retry/stop/error.

## Conventions
- Callbacks follow the pattern: `log_fn`, `warn_fn`, `progress_fn`, `status_fn`,
  `save_progress_fn`, `check_stop_fn`. The UI wires these to Streamlit widgets;
  the CLI uses print/no-ops.
- Excel workbooks always have four sheets: Findings, Sentences, Audit Settings,
  Errors.
- The `_is_retryable_llm_error()` function in `audit.py` is the single source of
  truth for which errors trigger retry logic.
