"""Tests for audit-app.py — Streamlit UI logic (non-widget functions only)."""
from __future__ import annotations

import os
from datetime import datetime
from io import BytesIO
from unittest.mock import MagicMock, patch

import pytest

from tests.conftest import _load_module


# We cannot import audit-app.py directly because it calls st.set_page_config()
# at module level. Instead we test the specific utility functions via importlib
# with Streamlit mocked out.

def _load_app_module():
    """Load audit-app.py with Streamlit fully mocked to prevent side effects."""
    import sys
    import importlib.util

    # Create a mock streamlit module to prevent any Streamlit calls at import time
    mock_st = MagicMock()
    mock_st.secrets = {"APP_PASSWORD": "test", "ORGANIZATION": "TestOrg", "AUDIENCE": "testers"}
    mock_st.session_state = {"authenticated": True}
    mock_st.set_page_config = MagicMock()
    mock_st.cache_data = MagicMock()

    # Mock streamlit_tree_select too
    mock_tree_select = MagicMock()

    saved_modules = {}
    modules_to_mock = {
        "streamlit": mock_st,
        "streamlit_tree_select": mock_tree_select,
    }
    for mod_name, mock_mod in modules_to_mock.items():
        if mod_name in sys.modules:
            saved_modules[mod_name] = sys.modules[mod_name]
        sys.modules[mod_name] = mock_mod

    try:
        module_path = os.path.join(os.path.dirname(__file__), "..", "audit-app.py")
        spec = importlib.util.spec_from_file_location("audit_app", module_path)
        if spec is None or spec.loader is None:
            raise ImportError("Unable to load audit-app.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        for mod_name, original in saved_modules.items():
            sys.modules[mod_name] = original
        for mod_name in modules_to_mock:
            if mod_name not in saved_modules:
                sys.modules.pop(mod_name, None)


# ===========================================================================
# _build_completed_filename
# ===========================================================================

class TestBuildCompletedFilename:
    def test_basic_xlsx(self):
        app = _load_app_module()
        mock_file = MagicMock()
        mock_file.name = "my_audit.xlsx"
        result = app._build_completed_filename(mock_file)
        suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
        assert result == f"my_audit_completed_{suffix}.xlsx"

    def test_strips_sortable_suffix(self):
        app = _load_app_module()
        mock_file = MagicMock()
        mock_file.name = "my_audit_sortable.xlsx"
        result = app._build_completed_filename(mock_file)
        suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
        assert result == f"my_audit_completed_{suffix}.xlsx"

    def test_no_extension(self):
        app = _load_app_module()
        mock_file = MagicMock()
        mock_file.name = "audit_file"
        result = app._build_completed_filename(mock_file)
        suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
        assert result == f"audit_file_completed_{suffix}.xlsx"

    def test_no_name_attribute(self):
        app = _load_app_module()
        mock_file = MagicMock(spec=[])  # no .name attribute
        result = app._build_completed_filename(mock_file)
        suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
        assert result == f"completed_audit_completed_{suffix}.xlsx"

    def test_strips_checkpoint_suffix_legacy_date_only(self):
        """Old checkpoints used `_checkpoint_YYYY-MM-DD` (no time). Should still strip."""
        app = _load_app_module()
        mock_file = MagicMock()
        mock_file.name = "audit_3_Fraud_07_04_26_checkpoint_2026-04-09.xlsx"
        result = app._build_completed_filename(mock_file)
        suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
        assert result == f"audit_3_Fraud_07_04_26_completed_{suffix}.xlsx"

    def test_strips_checkpoint_suffix_with_time(self):
        """New checkpoints include `_HHMM`. Should strip that too on re-upload."""
        app = _load_app_module()
        mock_file = MagicMock()
        mock_file.name = "audit_3_Fraud_07_04_26_checkpoint_2026-04-09_1530.xlsx"
        result = app._build_completed_filename(mock_file)
        suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
        assert result == f"audit_3_Fraud_07_04_26_completed_{suffix}.xlsx"

    def test_strips_multiple_checkpoint_suffixes_mixed_formats(self):
        app = _load_app_module()
        mock_file = MagicMock()
        mock_file.name = (
            "audit_3_Fraud_07_04_26"
            "_checkpoint_2026-04-09"  # legacy date-only
            "_checkpoint_2026-04-09_1530"  # new with time
            "_completed_2026-04-09_1645"
            ".xlsx"
        )
        result = app._build_completed_filename(mock_file)
        suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
        assert result == f"audit_3_Fraud_07_04_26_completed_{suffix}.xlsx"

    def test_strips_completed_suffix_on_reupload(self):
        app = _load_app_module()
        mock_file = MagicMock()
        mock_file.name = "my_audit_completed_2026-04-01.xlsx"
        result = app._build_completed_filename(mock_file)
        suffix = datetime.now().strftime("%Y-%m-%d_%H%M")
        assert result == f"my_audit_completed_{suffix}.xlsx"


# ===========================================================================
# _strip_status_suffixes
# ===========================================================================

class TestStripStatusSuffixes:
    def test_strips_legacy_date_only(self):
        app = _load_app_module()
        assert app._strip_status_suffixes("foo_checkpoint_2026-04-09") == "foo"
        assert app._strip_status_suffixes("foo_completed_2026-04-09") == "foo"

    def test_strips_new_with_time(self):
        app = _load_app_module()
        assert app._strip_status_suffixes("foo_checkpoint_2026-04-09_1530") == "foo"
        assert app._strip_status_suffixes("foo_completed_2026-04-09_1530") == "foo"

    def test_strips_repeated_suffixes(self):
        app = _load_app_module()
        result = app._strip_status_suffixes(
            "foo_checkpoint_2026-04-09_checkpoint_2026-04-09_1530_completed_2026-04-10_0900"
        )
        assert result == "foo"

    def test_leaves_unrelated_dates_alone(self):
        """A date inside the base name (not at the end with _checkpoint/_completed prefix) shouldn't be stripped."""
        app = _load_app_module()
        assert app._strip_status_suffixes("audit_2026-04-09_data") == "audit_2026-04-09_data"


# ===========================================================================
# _topic_key / _normalize_topic (app version)
# ===========================================================================

class TestTopicKeyApp:
    def test_basic(self):
        app = _load_app_module()
        assert app._topic_key("Food") == "food"

    def test_arrow_normalization(self):
        app = _load_app_module()
        assert app._topic_key("Parent --> Child") == "parent-->child"
        assert app._topic_key("Parent  -->  Child") == "parent-->child"

    def test_extra_whitespace(self):
        app = _load_app_module()
        assert app._topic_key("  Food   Quality  ") == "food quality"


# ===========================================================================
# _completed / _partial / _in_progress filename replacements
# ===========================================================================

class TestFilenameReplacements:
    """Test that filename manipulation patterns used in the app work correctly."""

    def test_completed_to_partial(self):
        filename = "my_audit_completed_2024-01-15.xlsx"
        partial = filename.replace("_completed", "_partial")
        assert partial == "my_audit_partial_2024-01-15.xlsx"

    def test_completed_to_in_progress(self):
        filename = "my_audit_completed_2024-01-15.xlsx"
        in_progress = filename.replace("_completed", "_in_progress")
        assert in_progress == "my_audit_in_progress_2024-01-15.xlsx"

    def test_no_completed_suffix(self):
        filename = "raw_audit.xlsx"
        partial = filename.replace("_completed", "_partial")
        # No change if _completed not present
        assert partial == "raw_audit.xlsx"
