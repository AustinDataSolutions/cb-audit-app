"""Tests for audit_validation.py — input file validation."""
from __future__ import annotations

from io import BytesIO

import pytest
from openpyxl import Workbook

from tests.conftest import _load_module, _make_workbook_bytes

validation = _load_module("audit_validation.py", "audit_validation")


# ===========================================================================
# validate_audit_sentences_sheet — input format
# ===========================================================================

class TestValidateInputFormat:
    def test_basic_input_format(self):
        wb_bytes = _make_workbook_bytes({
            "Sentences": [
                ["#", "Sentences", "Category"],
                [1, "The food was great", "Food"],
                [2, "Service was slow", "Service"],
            ]
        })
        sheet, row_idx, indices, warnings, is_output = (
            validation.validate_audit_sentences_sheet(wb_bytes)
        )
        assert sheet == "Sentences"
        assert row_idx == 0
        assert is_output is False
        assert len(warnings) == 0

    def test_headers_in_second_row(self):
        wb_bytes = _make_workbook_bytes({
            "Sentences": [
                ["Some title row", "", ""],
                ["#", "Sentences", "Category"],
                [1, "The food was great", "Food"],
            ]
        })
        sheet, row_idx, indices, warnings, is_output = (
            validation.validate_audit_sentences_sheet(wb_bytes)
        )
        assert row_idx == 1
        assert is_output is False

    def test_falls_back_to_first_sheet(self):
        wb_bytes = _make_workbook_bytes({
            "MyData": [
                ["#", "Sentences", "Category"],
                [1, "Food is great", "Food"],
            ]
        })
        sheet, _, _, _, _ = validation.validate_audit_sentences_sheet(wb_bytes)
        assert sheet == "MyData"

    def test_no_data_rows_raises(self):
        wb_bytes = _make_workbook_bytes({
            "Sentences": [
                ["#", "Sentences", "Category"],
            ]
        })
        with pytest.raises(ValueError, match="at least one data row"):
            validation.validate_audit_sentences_sheet(wb_bytes)

    def test_missing_required_headers_raises(self):
        wb_bytes = _make_workbook_bytes({
            "Sentences": [
                ["Col1", "Col2", "Col3"],
                [1, "data", "data"],
            ]
        })
        with pytest.raises(ValueError, match="must include headers"):
            validation.validate_audit_sentences_sheet(wb_bytes)


# ===========================================================================
# validate_audit_sentences_sheet — output format
# ===========================================================================

class TestValidateOutputFormat:
    def test_new_output_format_with_id(self):
        wb_bytes = _make_workbook_bytes({
            "Sentences": [
                ["ID", "Sentence", "Topic", "Audit", "Explanation"],
                [1, "Food was great", "Food", "YES", "Correct"],
            ]
        })
        sheet, row_idx, indices, warnings, is_output = (
            validation.validate_audit_sentences_sheet(wb_bytes)
        )
        assert is_output is True
        assert row_idx == 0

    def test_old_output_format_without_id(self):
        wb_bytes = _make_workbook_bytes({
            "Sentences": [
                ["Sentence", "Topic", "Audit", "Explanation"],
                ["Food was great", "Food", "YES", "Correct"],
            ]
        })
        sheet, row_idx, indices, warnings, is_output = (
            validation.validate_audit_sentences_sheet(wb_bytes)
        )
        assert is_output is True

    def test_empty_file_raises(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sentences"
        # Completely empty — no headers at all
        out = BytesIO()
        wb.save(out)
        wb.close()
        out.seek(0)
        with pytest.raises(ValueError):
            validation.validate_audit_sentences_sheet(out.getvalue())


# ===========================================================================
# _normalize_header
# ===========================================================================

class TestNormalizeHeader:
    def test_string(self):
        assert validation._normalize_header("  Sentences  ") == "sentences"

    def test_none(self):
        assert validation._normalize_header(None) == ""

    def test_number(self):
        assert validation._normalize_header(42) == "42"

    def test_nan(self):
        import math
        assert validation._normalize_header(float("nan")) == ""


# ===========================================================================
# _find_header_row
# ===========================================================================

class TestFindHeaderRow:
    def test_input_format_first_row(self):
        import pandas as pd
        df = pd.DataFrame([["#", "Sentences", "Category"]])
        row_idx, indices, is_output = validation._find_header_row(df)
        assert row_idx == 0
        assert is_output is False

    def test_output_format_new(self):
        import pandas as pd
        df = pd.DataFrame([["ID", "Sentence", "Topic", "Audit", "Explanation"]])
        row_idx, indices, is_output = validation._find_header_row(df)
        assert row_idx == 0
        assert is_output is True

    def test_no_matching_headers_raises(self):
        import pandas as pd
        df = pd.DataFrame([["A", "B", "C"], ["D", "E", "F"]])
        with pytest.raises(ValueError, match="must include headers"):
            validation._find_header_row(df)
