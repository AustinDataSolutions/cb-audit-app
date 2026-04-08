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
- `.streamlit/secrets.toml`: APP_PASSWORD, ANTHROPIC_API_KEY, OPENAI_API_KEY
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
  `importlib.util` (see `tests/conftest.py` for the pattern and `audit-app.py`
  `_load_summarizer_module()`).

### Module roles
| File | Role |
|------|------|
| `audit-app.py` | Streamlit UI, session state management, progress callbacks |
| `audit.py` | Core audit engine (`run_audit`), LLM calls, retry logic, partial detection |
| `audit-report-summarizer.py` | Post-audit summary generation for low-accuracy topics |
| `audit_validation.py` | Input file format detection and header validation |
| `audit_reformat.py` | Pre-processing to normalize input files to standard format |

### LLM call pattern
All LLM calls go through `_call_llm_with_status()` which runs the API call in a
background thread and shows periodic "waiting for response" status updates. Both
`audit.py` and `audit-report-summarizer.py` have their own copy of this helper.
LLM calls have a 300s timeout (`DEFAULT_LLM_TIMEOUT`). Retryable errors (429,
503, 529, timeouts) retry up to 3 times with delays of 30s, 60s, 120s.

### Streamlit session state keys
- `audit_in_progress`: True while audit is running
- `partial_audit_bytes`: latest saved workbook bytes (updated after each category)
- `audit_output_bytes`: final or promoted partial output for download
- `audit_stop_requested` / `summary_stop_requested`: stop signal flags
- `audit_is_partial`: whether current output is incomplete

### Stuck state recovery
If `audit_in_progress=True` but no audit is starting (`should_run_audit=False`),
the app detects this as an interrupted run and promotes `partial_audit_bytes` to
downloadable output. This handles Streamlit reruns that kill a blocking LLM call
before the `finally` block executes.

## Conventions
- Callbacks follow the pattern: `log_fn`, `warn_fn`, `progress_fn`, `status_fn`,
  `save_progress_fn`, `check_stop_fn`. The UI wires these to Streamlit widgets;
  the CLI uses print/no-ops.
- Excel workbooks always have four sheets: Findings, Sentences, Audit Settings,
  Errors.
- The `_is_retryable_llm_error()` function in `audit.py` is the single source of
  truth for which errors trigger retry logic.
