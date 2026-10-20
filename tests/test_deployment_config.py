"""Render Blueprint, start and build scripts, Streamlit server config, pins, the local env example, and
the standalone Community Cloud notice."""
import ast
import os
import tomllib
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent


def _service():
    blueprint = yaml.safe_load((ROOT / "render.yaml").read_text())
    assert list(blueprint) == ["services"]
    assert len(blueprint["services"]) == 1
    return blueprint["services"][0]


def test_single_web_service_runs_the_router_through_preflight():
    svc = _service()
    assert svc["name"] == "aarp-automatic-audit"
    assert svc["type"] == "web" and svc["runtime"] == "python"
    assert svc["region"] == "virginia" and svc["plan"] == "0.5c-512mb" and svc["numInstances"] == 1
    assert svc["autoDeployTrigger"] == "off"
    assert svc["buildCommand"] == "./scripts/render-build.sh"
    assert svc["startCommand"] == "./scripts/start.sh"
    assert svc["healthCheckPath"] == "/_stcore/health"
    assert svc["branch"] == "main"
    start = (ROOT / "scripts" / "start.sh").read_text()
    assert os.access(ROOT / "scripts" / "start.sh", os.X_OK)
    assert start.index("python scripts/preflight.py") < start.index("exec streamlit run app.py")
    assert "--server.address 0.0.0.0" in start and '"${PORT:-8501}"' in start and "--server.headless true" in start


def test_disk_holds_the_accounts_database_and_matches_data_dir():
    svc = _service()
    env = {e["key"]: e for e in svc["envVars"]}
    assert svc["disk"] == {"name": "automatic-audit-data", "mountPath": "/var/data", "sizeGB": 1}
    assert env["DATA_DIR"] == {"key": "DATA_DIR", "value": "/var/data"}
    assert env["AUTH_MODE"] == {"key": "AUTH_MODE", "value": "required"}
    assert env["STREAMLIT_SERVER_COOKIE_SECRET"] == {"key": "STREAMLIT_SERVER_COOKIE_SECRET", "generateValue": True}


def test_keys_and_mail_settings_are_dashboard_values_without_committed_secrets():
    env = {e["key"]: e for e in _service()["envVars"]}
    for key in ("ANTHROPIC_API_KEY", "OPENAI_API_KEY", "SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "EMAIL_FROM", "AUDIENCE"):
        assert env[key] == {"key": key, "sync": False}, key
    assert env["ORGANIZATION"] == {"key": "ORGANIZATION", "value": "AARP"}
    assert env["SMTP_PORT"] == {"key": "SMTP_PORT", "value": "587"}
    assert "APP_PASSWORD" not in env                     # the shared password is gone
    plain = {"DATA_DIR", "AUTH_MODE", "PYTHONUNBUFFERED", "ORGANIZATION", "SMTP_PORT"}
    for e in env.values():
        assert "value" not in e or e["key"] in plain, e["key"]


def test_build_script_installs_requirements_and_never_touches_the_disk():
    script = ROOT / "scripts" / "render-build.sh"
    assert os.access(script, os.X_OK)
    text = script.read_text()
    assert "set -euo pipefail" in text and "requirements.txt" in text
    assert "compileall -q app.py audit_page.py manage.py core scripts" in text
    assert "/var/data" not in text.replace("touch /var/data", "")


def test_python_pin_and_requirements():
    assert (ROOT / ".python-version").read_text().strip() == "3.12.3"
    reqs = (ROOT / "requirements.txt").read_text()
    assert "streamlit==1.52.2" in reqs
    assert "ads-auth @ git+https://github.com/AustinDataSolutions/ads-auth.git@v1.1.1" in reqs
    for pkg in ("anthropic==", "openai==", "python-dotenv==", "pandas==", "openpyxl=="):
        assert pkg in reqs


def test_streamlit_server_config():
    cfg = tomllib.loads((ROOT / ".streamlit" / "config.toml").read_text())
    assert cfg["browser"]["gatherUsageStats"] is False
    assert cfg["server"]["headless"] is True and cfg["server"]["runOnSave"] is False
    assert cfg["server"]["fileWatcherType"] == "none"
    assert cfg["server"]["maxUploadSize"] == 50 and cfg["server"]["maxMessageSize"] == 50
    assert cfg["server"]["websocketPingInterval"] == 120
    assert "enableCORS" not in cfg["server"] and "enableXsrfProtection" not in cfg["server"]


def test_local_env_example_replaces_the_secrets_file_and_the_shared_password():
    assert not (ROOT / ".streamlit" / "secrets.toml.example").exists()
    example = (ROOT / ".env.example").read_text()
    assert "APP_PASSWORD" not in example
    assert "AUTH_MODE=disabled" in example and "DATA_DIR=.data" in example
    assert "Never use this mode on Render" in example
    for key in ("ORGANIZATION", "AUDIENCE", "ANTHROPIC_API_KEY", "OPENAI_API_KEY", "SMTP_HOST", "SMTP_PORT",
                "SMTP_USER", "SMTP_PASSWORD", "EMAIL_FROM"):
        assert f"\n{key}=" in example, key
    ignored = (ROOT / ".gitignore").read_text().splitlines()
    for pattern in (".env", ".data/", "*.db", ".streamlit/secrets.toml"):
        assert pattern in ignored, pattern


def test_no_module_reads_streamlit_secrets_or_the_shared_password():
    for path in sorted(ROOT.glob("*.py")) + sorted((ROOT / "core").glob("*.py")):
        text = path.read_text(encoding="utf-8")
        assert "st.secrets" not in text, path.name
        assert "APP_PASSWORD" not in text, path.name


def test_community_cloud_notice_is_standalone():
    source = (ROOT / "audit-app.py").read_text()
    assert "https://aarp-automatic-audit.onrender.com/" in source
    imported = set()
    for node in ast.walk(ast.parse(source)):
        if isinstance(node, ast.Import):
            imported.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom):
            imported.add((node.module or "").split(".")[0])
    assert imported == {"streamlit"}  # no app, core, audit, or provider client imports
