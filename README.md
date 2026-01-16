# audit-app
Streamlit app for automating audits of Qualtrics/CX Designer topic models with
LLM-assisted accuracy checks. The core application lives in `audit-app.py`.

## Overview
This app ingests a Qualtrics audit export (.xlsx), optionally enriches it with
an XML model tree, and runs category-level audits using an LLM. It produces a
completed audit workbook with per-topic accuracy metrics and can optionally
generate summaries of issues for low-accuracy topics.

## Features
- Streamlit UI with password gate and provider selection (Anthropic/OpenAI).
- Upload-based workflow for audit spreadsheets and optional model tree XML.
- Topic selection via tree view, plus LLM prompt customization.
- Built-in safeguards: limits for categories, sentences, and tokens.
- Partial audit detection with in-progress downloads and resume support.
- Optional summary generation for low-accuracy topics.

## Repository layout
- `audit-app.py`: Streamlit app entry point.
- `audit.py`: audit engine and CLI runner.
- `audit-report-summarizer.py`: summary generator (CLI and module).
- `audit_reformat.py`: input reformat helper.
- `audit_validation.py`: input validation utilities.
- `config.yaml`: defaults for app/CLI.
- `prompts.yaml`: audit and summary prompt templates.

## Requirements
- Python and pip
- Dependencies listed in `requirements.txt`

## Quickstart
```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run audit-app.py
```

## Configuration
### Streamlit secrets
The app expects a password and optionally API keys via Streamlit secrets. Create
`.streamlit/secrets.toml` (or set Streamlit secrets in your deployment) with:
```toml
APP_PASSWORD = "your-password"
ANTHROPIC_API_KEY = "your-anthropic-key"
OPENAI_API_KEY = "your-openai-key"
```

### App defaults
`config.yaml` controls default provider/model and limits shown in the UI:
```yaml
app_defaults:
  llm_provider: "anthropic"
  model_name_anthropic: "claude-opus-4-5"
  model_name_openai: "gpt-5-nano"
  max_categories: 500
  max_sentences_per_category: 50
  max_tokens: 10000
```

### Prompt templates
`prompts.yaml` contains the templates used for auditing and summaries. You can
edit these to tune instructions or formatting.

## Using the app
1. Upload the audit `.xlsx` exported from Qualtrics/CX Designer.
2. (Optional) Upload the model tree `.xml` to select topics and add descriptions.
3. Choose topics to audit and adjust the audit prompt if needed.
4. Set LLM provider, model, API key, and limits in the sidebar.
5. Run the audit and download the completed workbook.
6. (Optional) Generate summaries for low-accuracy topics.

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
pytest -q
```

## Troubleshooting
- Missing API key: set `ANTHROPIC_API_KEY` or `OPENAI_API_KEY` in Streamlit
  secrets or your environment.
- Wrong file format: ensure the audit export is a valid `.xlsx` and includes a
  Sentences/Topics (or Findings) sheet layout expected by the app.
- Model tree mismatch: upload the correct XML to align category names.
