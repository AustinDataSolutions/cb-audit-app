from __future__ import annotations

import argparse
from datetime import datetime
from io import BytesIO
import os
import re

import anthropic
from dotenv import load_dotenv
from openpyxl import Workbook
import pandas as pd
import yaml

try:
    from openai import OpenAI
except ImportError:  # Optional dependency
    OpenAI = None


DEFAULT_MODEL = "claude-sonnet-4-5"
DEFAULT_OPENAI_MODEL = "gpt-5-nano"
DEFAULT_MAX_TOKENS = 10000
DEFAULT_ACCURACY_THRESHOLD = 0.80


def _build_completed_filename(input_name):
    base, ext = os.path.splitext(input_name)
    for suffix in ("_sortable", "_completed"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
    if not ext:
        ext = ".xlsx"
    return f"{base}_completed{ext}"


def _build_summary_filename(input_name):
    base, ext = os.path.splitext(input_name)
    for suffix in ("_sortable", "_completed", "_summary"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
    if not ext:
        ext = ".xlsx"
    return f"{base}_summary{ext}"


def _load_prompts_config(prompts_path, config_key):
    with open(prompts_path, 'r') as f:
        prompts = yaml.safe_load(f)
    return prompts[config_key]


def _get_llm_client(llm_provider, anthropic_api_key=None, openai_api_key=None):
    load_dotenv()

    if llm_provider == "anthropic":
        api_key = anthropic_api_key or os.getenv('ANTHROPIC_API_KEY')
        if not api_key:
            raise RuntimeError("ANTHROPIC_API_KEY environment variable not set")
        return anthropic.Anthropic(api_key=api_key)

    if llm_provider == "openai":
        if OpenAI is None:
            raise RuntimeError("openai package is not installed")
        api_key = openai_api_key or os.getenv('OPENAI_API_KEY')
        if not api_key:
            raise RuntimeError("OPENAI_API_KEY environment variable not set")
        return OpenAI(api_key=api_key)

    raise ValueError("llm_provider must be 'anthropic' or 'openai'")


def _coerce_audit_bytes(audit_excel_input):
    if audit_excel_input is None:
        raise ValueError("audit_excel_input is required")

    if isinstance(audit_excel_input, (bytes, bytearray)):
        return bytes(audit_excel_input)

    if isinstance(audit_excel_input, BytesIO):
        return audit_excel_input.getvalue()

    if hasattr(audit_excel_input, "read"):
        return audit_excel_input.read()

    if isinstance(audit_excel_input, (str, os.PathLike)):
        with open(audit_excel_input, "rb") as f:
            return f.read()

    raise TypeError("audit_excel_input must be bytes, BytesIO, file-like, or a path")


def _build_audit_findings(df):
    if len(df.columns) < 5:
        raise ValueError("Audit report must contain at least five columns.")

    id_col = df.columns[0]
    sentence_col = df.columns[1]
    category_col = df.columns[2]
    judgment_col = df.columns[3]
    explanation_col = df.columns[4]

    audit_findings = {}
    for _, row in df.iterrows():
        category = row[category_col]
        sentence_id = row[id_col]
        sentence = row[sentence_col]
        judgment = row[judgment_col]
        explanation = row[explanation_col]
        if pd.isna(category) or pd.isna(sentence) or pd.isna(sentence_id) or pd.isna(judgment) or pd.isna(explanation):
            continue
        if category not in audit_findings:
            audit_findings[category] = {}
        audit_findings[category][sentence_id] = (judgment, explanation)
    return audit_findings


def _parse_llm_summary(response_text, log_fn):
    pattern = r"SUMMARY:\s*(.+?)\s*RECOMMENDATION:\s*(.+)"
    matches = re.findall(pattern, response_text, re.IGNORECASE | re.DOTALL)

    if matches:
        summary, recommendation = matches[0]
        return summary.strip(), recommendation.strip()

    log_fn("WARNING: REGEX FAILED TO PARSE LLM RESPONSE")
    summary_match = re.search(r"SUMMARY:\s*(.+?)(?=RECOMMENDATION:|$)", response_text, re.IGNORECASE | re.DOTALL)
    rec_match = re.search(r"RECOMMENDATION:\s*(.+?)$", response_text, re.IGNORECASE | re.DOTALL)

    if summary_match and rec_match:
        return summary_match.group(1).strip(), rec_match.group(1).strip()

    return "REGEX FAILED TO PARSE LLM RESPONSE", response_text


def summarize_audit_report(
    audit_excel_input,
    msg_template=None,
    prompts_path=None,
    output_path=None,
    anthropic_api_key=None,
    openai_api_key=None,
    llm_provider="anthropic",
    model_name=DEFAULT_MODEL,
    max_tokens=DEFAULT_MAX_TOKENS,
    accuracy_threshold=DEFAULT_ACCURACY_THRESHOLD,
    log_fn=None,
    warn_fn=None,
    progress_fn=None,
):
    if log_fn is None:
        log_fn = lambda *_args, **_kwargs: None
    if warn_fn is None:
        warn_fn = log_fn

    script_dir = os.path.dirname(os.path.abspath(__file__))
    prompts_path = prompts_path or os.path.join(script_dir, "prompts.yaml")
    if msg_template is None:
        summarizer_config = _load_prompts_config(prompts_path, "audit-report-summarizer")
        msg_template = summarizer_config["rewards_msg_template"]

    audit_bytes = _coerce_audit_bytes(audit_excel_input)
    df = pd.read_excel(BytesIO(audit_bytes))
    audit_findings = _build_audit_findings(df)

    client = _get_llm_client(llm_provider, anthropic_api_key, openai_api_key)
    if model_name is None:
        model_name = DEFAULT_MODEL if llm_provider == "anthropic" else DEFAULT_OPENAI_MODEL

    wb = Workbook()
    ws = wb.active
    ws.append(["Category", "Accuracy", "Issues", "Recommendation"])

    total_categories = len(audit_findings)
    categories_checked = 1
    for category, findings in audit_findings.items():
        if progress_fn:
            progress_fn(categories_checked, total_categories, category)
        # log_fn(f"Reviewing audit findings for category '{category}' ({categories_checked} of {total_categories})...")
        inaccurate_sent_explanations = ""
        sent_count = 0
        wrong_count = 0

        for _, (judgment, explanation) in findings.items():
            sent_count += 1
            if judgment == "NO":
                wrong_count += 1
                inaccurate_sent_explanations += f"{explanation}\n"

        accuracy = round(((sent_count - wrong_count) / sent_count), 2) if sent_count else 0
        # log_fn(f"Detected {wrong_count} explanations out of {sent_count} sentences audited ({round(accuracy * 100)}% accuracy)")

        if accuracy < accuracy_threshold:
            message_content = msg_template.format(
                category=category,
                inaccurate_sent_explanations=inaccurate_sent_explanations,
            )
            # log_fn("Sending explanations to LLM for summarization...")
            if llm_provider == "anthropic":
                message = client.messages.create(
                    model=model_name,
                    max_tokens=max_tokens,
                    messages=[
                        {"role": "user", "content": message_content}
                    ]
                )
                response_text = message.content[0].text
            elif llm_provider == "openai":
                response = client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "user", "content": message_content}
                    ]
                )
                response_text = response.choices[0].message.content
            else:
                raise ValueError("llm_provider must be 'anthropic' or 'openai'")
            summary, recommendation = _parse_llm_summary(response_text, warn_fn)
            ws.append([category, accuracy, summary, recommendation])
        else:
            ws.append([category, accuracy, "", ""])

        categories_checked += 1

    output = BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)
    output_bytes = output.getvalue()

    if output_path:
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        with open(output_path, "wb") as f:
            f.write(output_bytes)

    return output_bytes


def _resolve_output_path(outputs_dir, input_path=None):
    if input_path:
        input_name = os.path.basename(input_path)
        output_filename = _build_summary_filename(input_name)
    else:
        timestamp = datetime.now().strftime("%y%m%d%H%M")
        output_filename = f"audit_summary_{timestamp}.xlsx"
    return os.path.join(outputs_dir, output_filename)


def main():
    parser = argparse.ArgumentParser(description="Summarize a completed audit report.")
    parser.add_argument("--input", dest="input_path", help="Path to the completed audit .xlsx file")
    parser.add_argument("--output", dest="output_path", help="Path to write the summary .xlsx file")
    parser.add_argument("--prompts", dest="prompts_path", help="Path to prompts.yaml")
    parser.add_argument("--api-key", dest="api_key", help="Anthropic API key override")
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    prompts_path = args.prompts_path or os.path.join(script_dir, "prompts.yaml")
    summarizer_config = _load_prompts_config(prompts_path, "audit-report-summarizer")

    input_path = args.input_path
    if not input_path:
        audit_file_name = summarizer_config.get("audit_file")
        if not audit_file_name:
            raise ValueError("No input path provided and audit_file is missing in prompts.yaml")
        input_path = os.path.join(script_dir, "inputs", audit_file_name)

    outputs_dir = os.path.join(script_dir, "outputs")
    output_path = args.output_path or _resolve_output_path(outputs_dir, input_path)

    def _log(msg):
        print(msg)

    summarize_audit_report(
        audit_excel_input=input_path,
        msg_template=summarizer_config.get("rewards_msg_template"),
        prompts_path=prompts_path,
        output_path=output_path,
        anthropic_api_key=args.api_key,
        log_fn=_log,
    )

    print(f"Workbook saved to {output_path}")


if __name__ == "__main__":
    main()
