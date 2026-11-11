"""Server-log setup shared by the router and the audit screen. No I/O at import time.

What the log may hold is stated in DATA_HANDLING.md: account names and event types, run progress by
topic number and name, request timing and retries, and whether the results email was sent. Never
sentences, prompts, model answers, file names, email addresses, or passwords.
"""
from __future__ import annotations

import logging
import warnings

FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"

# Third-party loggers that report routine traffic at INFO (one line per model-list request, on every
# page load). Their warnings and errors still show.
QUIET_LOGGERS = ("httpx", "httpcore")

# openpyxl says this on every read of a workbook saved without a default style, which CX Designer
# exports are; the workbook reads fine. Four reads per upload made it most of the log.
OPENPYXL_NOTICE = "Workbook contains no default style"


def configure(level: int = logging.INFO) -> None:
    """Once per process, safe to repeat: the line format, the quiet third-party loggers, and the
    silenced openpyxl notice."""
    logging.basicConfig(level=level, format=FORMAT)
    for name in QUIET_LOGGERS:
        logging.getLogger(name).setLevel(logging.WARNING)
    warnings.filterwarnings("ignore", message=OPENPYXL_NOTICE, category=UserWarning, module=r"openpyxl\.")
