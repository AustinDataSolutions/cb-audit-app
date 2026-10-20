"""The operator wrapper (manage.py) runs the package's commands with this app's settings."""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _manage(tmp_path, *args):
    env = {"PATH": "/usr/bin:/bin", "DATA_DIR": str(tmp_path / "data"), "AUTH_MODE": "required", "RENDER": "true",
           "HOME": str(tmp_path)}
    return subprocess.run([sys.executable, str(ROOT / "manage.py"), *args], cwd=ROOT, env=env,
                          capture_output=True, text=True, timeout=120)


def test_help_lists_the_account_backup_and_event_commands(tmp_path):
    result = _manage(tmp_path, "--help")
    assert result.returncode == 0, result.stderr
    for command in ("create-user", "set-password", "set-role", "list-users", "import-users", "backup-db",
                    "restore-db", "prune-events", "export-events"):
        assert command in result.stdout, command


def test_list_users_creates_the_database_under_data_dir(tmp_path):
    result = _manage(tmp_path, "list-users")
    assert result.returncode == 0, result.stderr
    assert (tmp_path / "data" / "app.db").exists()


def test_create_user_with_a_temporary_password_then_list(tmp_path):
    result = _manage(tmp_path, "create-user", "pat", "--full-name", "Pat Example", "--role", "manager", "--temporary")
    assert result.returncode == 0, result.stderr
    result = _manage(tmp_path, "list-users")
    assert result.returncode == 0, result.stderr
    assert "pat" in result.stdout and "manager" in result.stdout
