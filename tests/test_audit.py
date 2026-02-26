from io import BytesIO
import importlib.util
import os

from openpyxl import Workbook, load_workbook
from openpyxl.styles import numbers


def _load_audit_module():
    module_path = os.path.join(
        os.path.dirname(__file__),
        "..",
        "audit.py",
    )
    spec = importlib.util.spec_from_file_location("audit", module_path)
    if spec is None or spec.loader is None:
        raise ImportError("Unable to load audit module.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _build_audit_workbook(sentences_data):
    """Build a minimal audit workbook from a list of (id, sentence, topic, audit, explanation) tuples."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sentences"
    ws.append(["ID", "Sentence", "Topic", "Audit", "Explanation"])
    for row in sentences_data:
        ws.append(list(row))

    ws_findings = wb.create_sheet("Findings")
    ws_findings.append(["Topic", "Description", "Accuracy", "Issues"])

    output = BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)
    return output.getvalue()


# ---------------------------------------------------------------------------
# _parse_llm_response
# ---------------------------------------------------------------------------

def test_parse_llm_response_valid():
    audit = _load_audit_module()
    response = (
        "ID: 1 - Judgment: YES - Reasoning: Looks accurate\n"
        "ID: 2 - Judgment: NO - Reasoning: Wrong category\n"
        "ID: 3 - Judgment: YES - Reasoning: Good fit\n"
    )
    result = audit._parse_llm_response(response)
    assert len(result) == 3
    assert result["1"] == ("YES", "Looks accurate")
    assert result["2"] == ("NO", "Wrong category")
    assert result["3"] == ("YES", "Good fit")


def test_parse_llm_response_case_insensitive():
    audit = _load_audit_module()
    response = (
        "ID: 1 - Judgment: yes - Reasoning: Fine\n"
        "ID: 2 - Judgment: No - Reasoning: Not good\n"
    )
    result = audit._parse_llm_response(response)
    assert result["1"] == ("YES", "Fine")
    assert result["2"] == ("NO", "Not good")


def test_parse_llm_response_empty_response():
    audit = _load_audit_module()
    result = audit._parse_llm_response("")
    assert result == {}


def test_parse_llm_response_malformed():
    audit = _load_audit_module()
    result = audit._parse_llm_response("This is not a valid response at all.")
    assert result == {}


def test_parse_llm_response_partial():
    audit = _load_audit_module()
    response = (
        "ID: 1 - Judgment: YES - Reasoning: Good\n"
        "This line is garbage\n"
        "ID: 3 - Judgment: NO - Reasoning: Bad\n"
    )
    result = audit._parse_llm_response(response)
    assert len(result) == 2
    assert "1" in result
    assert "2" not in result
    assert "3" in result


def test_parse_llm_response_missing_ids():
    """Parsed result may contain IDs that don't match what was sent; that's a parsing concern,
    not a _parse_llm_response concern. The function just returns what it finds."""
    audit = _load_audit_module()
    response = "ID: 999 - Judgment: YES - Reasoning: Phantom sentence\n"
    result = audit._parse_llm_response(response)
    assert result["999"] == ("YES", "Phantom sentence")


# ---------------------------------------------------------------------------
# _apply_precision_formula
# ---------------------------------------------------------------------------

def test_precision_formula_excludes_unaudited():
    audit = _load_audit_module()
    wb = Workbook()
    ws_sentences = wb.active
    ws_sentences.title = "Sentences"
    ws_sentences.append(["ID", "Sentence", "Topic", "Audit", "Explanation"])
    ws_sentences.append([1, "S1", "TopicA", "YES", "Good"])
    ws_sentences.append([2, "S2", "TopicA", "NO", "Bad"])
    ws_sentences.append([3, "S3", "TopicA", "", ""])  # un-audited

    ws_findings = wb.create_sheet("Findings")
    ws_findings.append(["Topic", "Description", "Accuracy", "Issues"])
    ws_findings.append(["TopicA", "Desc", "", ""])

    audit._apply_precision_formula(ws_findings, 2, "Sentences")

    formula = ws_findings.cell(row=2, column=3).value
    # Numerator: COUNTIFS matching TopicA and YES
    assert "YES" in formula
    # Denominator: COUNTIFS matching TopicA and non-empty (excludes un-audited)
    assert '<>"' in formula or "<>" in formula
    # Should NOT use plain COUNTIF for denominator (which would count all rows)
    # The formula has two COUNTIFS calls
    assert formula.count("COUNTIFS") == 2
    wb.close()


def test_precision_formula_format():
    audit = _load_audit_module()
    wb = Workbook()
    ws_findings = wb.active
    ws_findings.title = "Findings"
    ws_findings.append(["Topic", "Description", "Accuracy", "Issues"])
    ws_findings.append(["TopicA", "Desc", "", ""])

    ws_sentences = wb.create_sheet("Sentences")

    audit._apply_precision_formula(ws_findings, 2, "Sentences")

    cell = ws_findings.cell(row=2, column=3)
    assert cell.number_format == numbers.FORMAT_PERCENTAGE
    wb.close()


# ---------------------------------------------------------------------------
# _write_settings_sheet / _update_setting
# ---------------------------------------------------------------------------

def test_write_settings_sheet():
    audit = _load_audit_module()
    wb = Workbook()
    ws = audit._ensure_settings_sheet(wb)
    settings = {"Model": "gpt-4o", "Provider": "openai"}
    audit._write_settings_sheet(ws, settings)

    assert ws.cell(row=2, column=1).value == "Model"
    assert ws.cell(row=2, column=2).value == "gpt-4o"
    assert ws.cell(row=3, column=1).value == "Provider"
    assert ws.cell(row=3, column=2).value == "openai"
    wb.close()


def test_write_settings_sheet_replaces_existing():
    audit = _load_audit_module()
    wb = Workbook()
    ws = audit._ensure_settings_sheet(wb)

    audit._write_settings_sheet(ws, {"Key1": "Value1", "Key2": "Value2"})
    audit._write_settings_sheet(ws, {"Key3": "Value3"})

    # Should only have header + 1 data row now
    assert ws.max_row == 2
    assert ws.cell(row=2, column=1).value == "Key3"
    assert ws.cell(row=2, column=2).value == "Value3"
    wb.close()


def test_update_setting():
    audit = _load_audit_module()
    wb = Workbook()
    ws = audit._ensure_settings_sheet(wb)
    audit._write_settings_sheet(ws, {"Model": "gpt-4o", "Provider": "openai"})

    audit._update_setting(ws, "Model", "claude-sonnet-4-20250514")
    assert ws.cell(row=2, column=2).value == "claude-sonnet-4-20250514"
    # Other settings unchanged
    assert ws.cell(row=3, column=2).value == "openai"
    wb.close()


def test_update_setting_missing_key():
    audit = _load_audit_module()
    wb = Workbook()
    ws = audit._ensure_settings_sheet(wb)
    audit._write_settings_sheet(ws, {"Model": "gpt-4o"})

    # Should not raise
    audit._update_setting(ws, "Nonexistent", "value")
    # Original data unchanged
    assert ws.cell(row=2, column=2).value == "gpt-4o"
    wb.close()


# ---------------------------------------------------------------------------
# _write_errors_sheet
# ---------------------------------------------------------------------------

def test_write_errors_sheet_creates_sheet():
    audit = _load_audit_module()
    wb = Workbook()
    audit._write_errors_sheet(wb, ["Warning one", "Warning two"])

    assert "Errors" in wb.sheetnames
    ws = wb["Errors"]
    assert ws.cell(row=1, column=1).value == "Type"
    assert ws.cell(row=1, column=2).value == "Message"
    assert ws.cell(row=2, column=1).value == "Warning"
    assert ws.cell(row=2, column=2).value == "Warning one"
    assert ws.cell(row=3, column=1).value == "Warning"
    assert ws.cell(row=3, column=2).value == "Warning two"
    wb.close()


def test_write_errors_sheet_empty():
    audit = _load_audit_module()
    wb = Workbook()
    audit._write_errors_sheet(wb, [])
    assert "Errors" not in wb.sheetnames
    wb.close()


def test_write_errors_sheet_removes_empty():
    audit = _load_audit_module()
    wb = Workbook()
    # Create sheet with warnings first
    audit._write_errors_sheet(wb, ["Some warning"])
    assert "Errors" in wb.sheetnames

    # Now call with empty list — should remove the sheet
    audit._write_errors_sheet(wb, [])
    assert "Errors" not in wb.sheetnames
    wb.close()


def test_write_errors_sheet_replaces_on_rewrite():
    audit = _load_audit_module()
    wb = Workbook()

    audit._write_errors_sheet(wb, ["First warning"])
    audit._write_errors_sheet(wb, ["First warning", "Second warning"])

    ws = wb["Errors"]
    # Header + 2 data rows
    assert ws.max_row == 3
    assert ws.cell(row=2, column=2).value == "First warning"
    assert ws.cell(row=3, column=2).value == "Second warning"
    wb.close()


# ---------------------------------------------------------------------------
# _format_categories_selected
# ---------------------------------------------------------------------------

def test_format_categories_all():
    audit = _load_audit_module()
    result = audit._format_categories_selected(None, 50, 50)
    assert result == "All (50)"


def test_format_categories_all_empty_list():
    audit = _load_audit_module()
    result = audit._format_categories_selected([], 50, 50)
    assert result == "All (50)"


def test_format_categories_subset():
    audit = _load_audit_module()
    result = audit._format_categories_selected(["TopicA", "TopicB"], 2, 100)
    assert result == "2 of 100"


# ---------------------------------------------------------------------------
# detect_partial_audit
# ---------------------------------------------------------------------------

def test_detect_partial_completed():
    """All sentences judged — should NOT be marked as partial."""
    audit = _load_audit_module()
    wb_bytes = _build_audit_workbook([
        (1, "Sentence one", "TopicA", "YES", "Good"),
        (2, "Sentence two", "TopicA", "NO", "Bad"),
    ])
    result = audit.detect_partial_audit(wb_bytes)
    assert result["is_partial"] is False
    assert "TopicA" in result["completed_categories"]


def test_detect_partial_incomplete():
    """Some sentences blank — should be marked as partial."""
    audit = _load_audit_module()
    wb_bytes = _build_audit_workbook([
        (1, "Sentence one", "TopicA", "YES", "Good"),
        (2, "Sentence two", "TopicA", "", ""),
        (3, "Sentence three", "TopicB", "", ""),
    ])
    result = audit.detect_partial_audit(wb_bytes)
    assert result["is_partial"] is True
    assert "TopicA" in result["incomplete_categories"]
    assert "TopicB" in result["unjudged_categories"]


def test_detect_partial_not_audit_file():
    """File without expected sheets — should not be detected as partial."""
    audit = _load_audit_module()
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Column1", "Column2"])
    ws.append(["value1", "value2"])

    output = BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)

    result = audit.detect_partial_audit(output.getvalue())
    assert result["is_partial"] is False
