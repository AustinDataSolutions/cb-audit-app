# audit-app
Streamlit app for automating audits of Qualtrics/CX Designer topic models with
LLM-assisted accuracy checks. The audit screen lives in `audit_page.py`; `app.py`
is the router that serves it beside the shared About and Usage statistics pages.
Hosted on Render behind individual accounts (`DATA_HANDLING.md` is the
user-facing description of what the app stores and sends).

## Overview
This app ingests a Qualtrics audit export (.xlsx), optionally enriches it with
an XML model tree, and runs category-level audits using an LLM. It produces a
completed audit workbook with per-topic accuracy metrics and can optionally
generate summaries of issues for low-accuracy topics.

## Features
- Individual accounts, sessions, usage statistics, and an About page from the
  shared `ads-auth` package; provider selection (Anthropic/OpenAI).
- Upload-based workflow for audit spreadsheets and optional model tree XML.
- Topic selection via tree view, plus LLM prompt customization.
- Built-in safeguards: limits for categories, sentences, and tokens.
- Partial audit detection with in-progress downloads and resume support.
- Optional summary generation for low-accuracy topics.
- Optional email delivery of the finished workbook.

## Repository layout
- `app.py`: router (`streamlit run app.py`).
- `audit_page.py`: the audit screen.
- `core/config.py`: settings, read from environment variables.
- `manage.py`: operator commands (accounts, backups, usage events).
- `audit.py`: audit engine and CLI runner.
- `audit-report-summarizer.py`: summary generator (CLI and module).
- `audit_worker.py`: background run, summary, and email delivery.
- `audit_reformat.py`, `audit_validation.py`: input helpers.
- `config.yaml`: defaults for app/CLI. `prompts.yaml`: prompt templates.
- `render.yaml`, `scripts/`: Render Blueprint, build, start, preflight.
- `docs/render-deployment-runbook.md`, `docs/OPERATIONS.md`: deploying and operating.
- `audit-app.py`: the "moved" notice served by the old Community Cloud deploy.

## Requirements
- Python 3.12 (see `.python-version`)
- Dependencies listed in `requirements.txt` (the shared `ads-auth` package is
  installed from its GitHub tag, so installing needs network access and git)

## Quickstart
```bash
python -m venv venv
venv/bin/python -m pip install -r requirements.txt
cp .env.example .env            # fill in the values
venv/bin/python -m streamlit run app.py
```
With `AUTH_MODE=disabled` in `.env` the sign-in screen is skipped and nothing is
recorded (local development only; Render refuses it). With
`AUTH_MODE=required`, create an account first:
`venv/bin/python manage.py create-user <username>`.

## Configuration
### Environment variables
There is no shared password and no Streamlit secrets file. Every setting is an
environment variable, set on the Render service in the dashboard or in a local
git-ignored `.env`; `.env.example` lists them:

- `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`: the provider keys (one provider per run).
- `ORGANIZATION`, `AUDIENCE`: named in the prompts and the page title.
- `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `EMAIL_FROM`: optional;
  together they enable "Email results to me".
- `DATA_DIR`, `AUTH_MODE`, `USAGE_EXCLUDE_USERS`: the accounts database location
  and sign-in mode (see `core/config.py`).

### App defaults
`config.yaml` controls default provider/model and limits shown in the UI:
```yaml
app_defaults:
  llm_provider: "anthropic"
  model_name_anthropic: "claude-opus-5-5"
  model_name_openai: "gpt-6.1-terra"
  max_categories: 500
  max_sentences_per_category: 300
  max_tokens: 10000
```
The model names are fallbacks. The sidebar asks the chosen provider for its
model list and preselects the newest model of the same family as the configured
name: the newest Opus model for Anthropic, the newest Terra model for OpenAI
(the rule lives in `core/model_choice.py`). So the default keeps up with new
releases without an edit here; a short note under the model picker appears only
when the configured name is no longer in the provider's list. The configured
name is used as-is only when the list cannot be fetched. Every listed model stays selectable.
To prefer another family, name a model of that family (for example
`claude-sonnet-5-5`), and the rule follows it.

### Prompt templates
`prompts.yaml` contains the templates used for auditing and summaries. You can
edit these to tune instructions or formatting. What the prompts carry is
described for users in `DATA_HANDLING.md`; keep the two in step.

## Using the app
1. Sign in with your account.
2. Upload the audit `.xlsx` exported from Qualtrics/CX Designer.
3. (Optional) Upload the model tree `.xml` to select topics and add descriptions.
4. Choose topics to audit and adjust the audit prompt if needed.
5. Set LLM provider, model, and limits in the sidebar.
6. Run the audit and download the completed workbook, or have it emailed.
7. (Optional) Re-upload a completed audit to regenerate its summaries.

## CLI utilities (optional)
The repo includes Python entry points for non-UI runs.

### Run audit from config
`audit.py` can run an audit using values in `config.yaml` and files in `inputs/`.
```bash
python audit.py
```
This will read `config.yaml` -> `cli_audit` and write to `outputs/`. Create the
`inputs/` and `outputs/` directories if they do not exist.

### Summarize an existing audit
```bash
python audit-report-summarizer.py --input path/to/completed_audit.xlsx
```
This updates the workbook in place unless you pass `--output`.

## Testing
The test suite uses pytest.
```bash
venv/bin/python -m pytest -q
```

## Troubleshooting
- "API key is not configured on the server": set `ANTHROPIC_API_KEY` or
  `OPENAI_API_KEY` on the Render service (or in `.env`) and redeploy.
- Wrong file format: ensure the audit export is a valid `.xlsx` and includes a
  Sentences/Topics (or Findings) sheet layout expected by the app.
- Model tree mismatch: upload the correct XML to align category names.
- Locked out or forgotten password: an administrator runs
  `python manage.py set-password <username>` (see `docs/OPERATIONS.md`).
