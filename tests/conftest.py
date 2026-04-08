"""Shared fixtures for the audit-app test suite."""
from __future__ import annotations

import importlib.util
import os
from io import BytesIO

import pytest
from openpyxl import Workbook


# ---------------------------------------------------------------------------
# Module loaders (the source files aren't proper packages, so we use importlib)
# ---------------------------------------------------------------------------

def _load_module(filename, module_name):
    module_path = os.path.join(os.path.dirname(__file__), "..", filename)
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load {filename}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture
def audit_module():
    return _load_module("audit.py", "audit")


@pytest.fixture
def summarizer_module():
    return _load_module("audit-report-summarizer.py", "audit_report_summarizer")


@pytest.fixture
def validation_module():
    return _load_module("audit_validation.py", "audit_validation")


# ---------------------------------------------------------------------------
# Workbook builders
# ---------------------------------------------------------------------------

def _make_workbook_bytes(sheets: dict) -> bytes:
    """Build an xlsx from *sheets* = {sheet_name: [row, row, ...]} and return bytes."""
    wb = Workbook()
    first = True
    for name, rows in sheets.items():
        if first:
            ws = wb.active
            ws.title = name
            first = False
        else:
            ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append(row)
    out = BytesIO()
    wb.save(out)
    wb.close()
    out.seek(0)
    return out.getvalue()


@pytest.fixture
def make_workbook_bytes():
    """Expose _make_workbook_bytes as a fixture."""
    return _make_workbook_bytes


@pytest.fixture
def input_format_bytes():
    """Sample audit input file in input format (#, Sentences, Category)."""
    return _make_workbook_bytes(
        {
            "Sentences": [
                ["#", "Sentences", "Category"],
                [1, "The food was great", "Food"],
                [2, "Service was slow", "Service"],
                [3, "Nice ambiance", "Food"],
            ]
        }
    )


@pytest.fixture
def output_format_bytes():
    """Sample audit output file (ID, Sentence, Topic, Audit, Explanation)."""
    return _make_workbook_bytes(
        {
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
        }
    )


@pytest.fixture
def partial_output_bytes():
    """Partial audit: some sentences judged, some not."""
    return _make_workbook_bytes(
        {
            "Sentences": [
                ["ID", "Sentence", "Topic", "Audit", "Explanation"],
                [1, "The food was great", "Food", "YES", "Correct"],
                [2, "Service was slow", "Service", "", ""],
                [3, "Nice ambiance", "Food", "YES", "Accurate"],
                [4, "Drinks were cold", "Drinks", "", ""],
            ],
            "Findings": [
                ["Topic", "Description", "Accuracy", "Issues"],
                ["Food", "About food", "", ""],
                ["Service", "About service", "", "Not yet audited"],
                ["Drinks", "About drinks", "", "Not yet audited"],
            ],
        }
    )
