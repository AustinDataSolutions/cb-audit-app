from io import BytesIO
import importlib.util
import os

import pandas as pd


def _load_summarizer_module():
    module_path = os.path.join(
        os.path.dirname(__file__),
        "..",
        "audit-report-summarizer.py",
    )
    spec = importlib.util.spec_from_file_location("audit_report_summarizer", module_path)
    if spec is None or spec.loader is None:
        raise ImportError("Unable to load audit-report-summarizer module.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_coerce_audit_bytes_accepts_bytes():
    summarizer = _load_summarizer_module()
    data = b"test-bytes"
    assert summarizer._coerce_audit_bytes(data) == data


def test_coerce_audit_bytes_accepts_bytesio():
    summarizer = _load_summarizer_module()
    data = b"test-bytesio"
    assert summarizer._coerce_audit_bytes(BytesIO(data)) == data


def test_coerce_audit_bytes_accepts_file_like(tmp_path):
    summarizer = _load_summarizer_module()
    data = b"test-file-like"
    file_path = tmp_path / "sample.bin"
    file_path.write_bytes(data)
    with file_path.open("rb") as handle:
        assert summarizer._coerce_audit_bytes(handle) == data


def test_coerce_audit_bytes_accepts_path(tmp_path):
    summarizer = _load_summarizer_module()
    data = b"test-path"
    file_path = tmp_path / "sample.bin"
    file_path.write_bytes(data)
    assert summarizer._coerce_audit_bytes(str(file_path)) == data


def test_parse_llm_summary_parses_sections():
    summarizer = _load_summarizer_module()
    response_text = "SUMMARY: issue one; issue two RECOMMENDATION: tighten rules"
    summary, recommendation = summarizer._parse_llm_summary(response_text, lambda *_args: None)
    assert summary == "issue one; issue two"
    assert recommendation == "tighten rules"


def test_parse_llm_summary_reports_failure():
    summarizer = _load_summarizer_module()
    response_text = "No expected sections here."
    summary, recommendation = summarizer._parse_llm_summary(response_text, lambda *_args: None)
    assert summary == "REGEX FAILED TO PARSE LLM RESPONSE"
    assert recommendation == response_text


def test_build_audit_findings_skips_missing_rows():
    summarizer = _load_summarizer_module()
    df = pd.DataFrame(
        [
            [1, "Sentence one", "Category A", "NO", "Wrong category"],
            [2, None, "Category A", "YES", "Looks right"],
            [3, "Sentence three", "Category B", "YES", "Accurate"],
        ]
    )
    findings = summarizer._build_audit_findings(df)
    assert "Category A" in findings
    assert "Category B" in findings
    assert 1 in findings["Category A"]
    assert 2 not in findings["Category A"]
    assert 3 in findings["Category B"]
