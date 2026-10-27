"""Runtime settings, read from environment variables.

`Settings` extends the `ads_auth.AppSettings` contract (app name, sign-in title, data directory,
disclosure path, auth mode, usage event vocabulary, page registry) with this app's own values: the
organization and audience named in the prompts, the two model-provider API keys, and the email
(SMTP) settings behind "Email results to me".

Every value is an environment variable. On Render they are set on the service (secrets are entered in
the dashboard, never committed). Locally a `.env` file in the repository root is loaded unless the app
runs on Render; `.env.example` lists the names. Streamlit secrets are no longer read.

Variables: `DATA_DIR` (where the accounts and usage database lives; default `.data/` in the repository,
git-ignored; `/var/data` on Render), `AUTH_MODE` (`required`, the default, or `disabled` for local
development only), `USAGE_EXCLUDE_USERS` (comma-separated accounts left out of the Usage statistics page
by default; default `ads-dev`), `ORGANIZATION`, `AUDIENCE`, `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`,
`SMTP_HOST`, `SMTP_PORT` (default 587), `SMTP_USER`, `SMTP_PASSWORD`, `EMAIL_FROM`, and
`ADS_AUTH_COOKIE_SECRET` (optional; falls back to `STREAMLIT_SERVER_COOKIE_SECRET`, which Render sets, and signs
the stay-signed-in cookie whenever one is present).
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from ads_auth import (AUTH_MODES, AppSettings, auth_mode_from_environ, cookie_secret_from_environ,
                      exclude_users_from_environ)

REPO_ROOT = Path(__file__).resolve().parent.parent
APP_SLUG = "automatic-audit"          # `app` column in the shared event export; CSV download prefix
APP_TITLE = "Automatic Audit"         # sign-in screen title, prefixed with the organization when one is set

# Placeholders used in the prompts when the organization or audience is not configured.
DEFAULT_ORGANIZATION = "the organization"
DEFAULT_AUDIENCE = "customers and users"
DEFAULT_SMTP_PORT = 587

# The usage events this app records, beside the package's login / login_failed / logout. Every name here
# is listed in DATA_HANDLING.md; adding one is a disclosure change in the same commit.
USAGE_EVENTS = ("upload_audit_file", "upload_model_tree", "run_audit", "email_results", "download")

# Caption of the Usage statistics page: what each of this app's events means and how it is counted.
USAGE_DEFINITIONS = (
    "**Active user**: an account with at least one sign-in, upload, audit run, emailed result, or download "
    "in the period. Failed sign-ins are not activity. **Upload audit file**: one per audit workbook or "
    "checkpoint the app reads. **Upload model tree**: one per model tree file the app reads. **Run audit**: "
    "one per run started, whether a new audit, a resumed checkpoint, or a summary-only rerun of a completed "
    "audit; whether the run finished is not recorded. **Email results**: one per results email actually "
    "sent. **Download**: one per workbook downloaded, finished or checkpoint. Weeks start on Monday, newest "
    "period first; periods use Eastern time. Accounts listed under \"Exclude accounts\" are left out of every "
    "table and download on this page. File names, file contents, sentences, prompts, model responses, and "
    "email addresses are never recorded."
)

# One page registry for the whole process. The router (app.py) fills it; every Settings instance, including
# the one the audit page builds for its own gate and sidebar, shares it so page links resolve.
PAGES: dict[str, object] = {}

__all__ = ["AUTH_MODES", "APP_SLUG", "APP_TITLE", "DEFAULT_AUDIENCE", "DEFAULT_ORGANIZATION", "PAGES",
           "REPO_ROOT", "USAGE_DEFINITIONS", "USAGE_EVENTS", "Settings", "load_settings"]


@dataclass(frozen=True, kw_only=True)
class Settings(AppSettings):
    app_name: str = APP_SLUG
    title: str = APP_TITLE
    disclosure_path: Path = REPO_ROOT / "DATA_HANDLING.md"
    usage_exclude_users: tuple[str, ...] = ("ads-dev",)
    events: tuple[str, ...] = USAGE_EVENTS
    definitions: str = USAGE_DEFINITIONS
    pages: dict[str, object] = field(default_factory=lambda: PAGES, compare=False, repr=False)
    # This app's own values. Secrets are excluded from repr so they never reach a log line.
    organization: str = DEFAULT_ORGANIZATION
    audience: str = DEFAULT_AUDIENCE
    anthropic_api_key: str = field(default="", repr=False)
    openai_api_key: str = field(default="", repr=False)
    smtp_host: str = ""
    smtp_port: int = DEFAULT_SMTP_PORT
    smtp_user: str = field(default="", repr=False)
    smtp_password: str = field(default="", repr=False)
    email_from: str = ""

    def api_key(self, provider: str) -> str:
        """The configured key for `anthropic` or `openai`; empty when not configured."""
        return {"anthropic": self.anthropic_api_key, "openai": self.openai_api_key}.get(provider.lower(), "")

    @property
    def email_configured(self) -> bool:
        """True when every setting needed to actually send mail is present."""
        return all((self.smtp_host, self.smtp_user, self.smtp_password, self.email_from))

    @property
    def smtp(self) -> dict:
        """The SMTP values in the shape `audit_worker.EmailPayload` carries into the background job."""
        return {"host": self.smtp_host, "port": self.smtp_port, "user": self.smtp_user,
                "password": self.smtp_password, "sender": self.email_from}


def _text(environ, name: str, default: str = "") -> str:
    value = environ.get(name, "")
    return value.strip() if value and value.strip() else default


def load_settings(environ=None, *, dotenv_path: str | os.PathLike | None = ".env") -> Settings:
    """Build Settings from `environ` (defaults to os.environ).

    Performs no filesystem I/O. A `.env` file is loaded into os.environ first unless running on Render
    (`RENDER=true`) or `dotenv_path` is None. Values already in the environment take precedence over the file.
    """
    if environ is None:
        environ = os.environ
    is_render = environ.get("RENDER", "").lower() == "true"
    if dotenv_path is not None and not is_render and environ is os.environ:
        try:
            from dotenv import load_dotenv
            load_dotenv(dotenv_path, override=False)
        except ImportError:
            pass

    raw_dir = environ.get("DATA_DIR", "")
    data_dir = Path(raw_dir).expanduser() if raw_dir else REPO_ROOT / ".data"
    data_dir = data_dir if data_dir.is_absolute() else (REPO_ROOT / data_dir)
    # No I/O here: the router creates the directory when it starts.

    raw_port = _text(environ, "SMTP_PORT", str(DEFAULT_SMTP_PORT))
    if not raw_port.isdigit() or not 1 <= int(raw_port) <= 65535:
        raise ValueError("SMTP_PORT must be a whole number from 1 to 65535.")

    organization = _text(environ, "ORGANIZATION", DEFAULT_ORGANIZATION)
    title = APP_TITLE if organization == DEFAULT_ORGANIZATION else f"{organization} {APP_TITLE}"

    return Settings(
        title=title,
        data_dir=data_dir,
        is_render=is_render,
        auth_mode=auth_mode_from_environ(environ),
        cookie_secret=cookie_secret_from_environ(environ),
        usage_exclude_users=exclude_users_from_environ(environ, default="ads-dev"),
        organization=organization,
        audience=_text(environ, "AUDIENCE", DEFAULT_AUDIENCE),
        anthropic_api_key=_text(environ, "ANTHROPIC_API_KEY"),
        openai_api_key=_text(environ, "OPENAI_API_KEY"),
        smtp_host=_text(environ, "SMTP_HOST"),
        smtp_port=int(raw_port),
        smtp_user=_text(environ, "SMTP_USER"),
        smtp_password=_text(environ, "SMTP_PASSWORD"),
        email_from=_text(environ, "EMAIL_FROM"),
    )
