"""Automatic Audit: the router.

Run with `streamlit run app.py` (scripts/start.sh does so on Render after the preflight checks). The
audit screen itself is `audit_page.py`, registered here as a file page; the About and Usage statistics
pages come from the shared `ads_auth` package. Each page runs the sign-in gate itself, so this file
draws nothing of its own. `audit-app.py` is the standalone "moved" notice served by the legacy
Community Cloud deployment and imports nothing from here.
"""
from pathlib import Path

import streamlit as st

from ads_auth import db, maintenance
from ads_auth.ui import about as about_ui
from ads_auth.ui import reports as reports_ui
from ads_auth.ui.gate import require_login
from core.config import APP_TITLE, load_settings

settings = load_settings()

ABOUT_TITLE = "About this app and data handling"
REPORTS_TITLE = "Usage statistics"


@st.cache_resource
def _prepare_storage(data_dir: str, db_path: str) -> int:
    """Once per process: the data directory, the schema, and the daily backup and pruning thread."""
    Path(data_dir).mkdir(parents=True, exist_ok=True)
    version = db.migrate(db_path)
    maintenance.start_background(data_dir, db_path)
    return version


def about_page() -> None:
    # Readable without signing in, by design: the sign-in screen links here.
    about_ui.render(st, settings)


def reports_page() -> None:
    user = require_login(st, settings)
    reports_ui.render(st, settings, user)


def run() -> None:
    st.set_page_config(page_title=APP_TITLE, initial_sidebar_state="collapsed")
    _prepare_storage(str(settings.data_dir), str(settings.db_path))
    settings.pages["app"] = st.Page("audit_page.py", title=APP_TITLE, url_path="app", default=True)
    settings.pages["about"] = st.Page(about_page, title=ABOUT_TITLE, url_path="about")
    settings.pages["reports"] = st.Page(reports_page, title=REPORTS_TITLE, url_path="reports")
    st.navigation(list(settings.pages.values()), position="hidden").run()


if __name__ == "__main__":
    run()
