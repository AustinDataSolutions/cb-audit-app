"""core/logs.py: the shared server-log setup."""
import logging
import warnings

from core import logs


def test_configure_is_repeatable_and_quiets_routine_http_traffic():
    logs.configure()
    logs.configure()
    assert logging.getLogger("httpx").level == logging.WARNING
    assert logging.getLogger("httpcore").level == logging.WARNING
    # The root level is not checked: pytest installs its own handlers first, so basicConfig is a no-op here.


def test_configure_silences_only_the_openpyxl_default_style_notice():
    with warnings.catch_warnings(record=True) as seen:
        warnings.simplefilter("always")
        logs.configure()  # inserted ahead of "always"
        warnings.warn_explicit("Workbook contains no default style, apply openpyxl's default", UserWarning,
                               "stylesheet.py", 237, module="openpyxl.styles.stylesheet")
        warnings.warn_explicit("Data Validation extension is not supported", UserWarning,
                               "excel.py", 1, module="openpyxl.reader.excel")
        warnings.warn_explicit("Workbook contains no default style", UserWarning,
                               "other.py", 1, module="somewhere.else")
    assert [str(w.message) for w in seen] == [
        "Data Validation extension is not supported",
        "Workbook contains no default style",
    ]
