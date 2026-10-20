#!/usr/bin/env bash
# Render build step. Keep this script idempotent and free of secrets.
# It must not touch /var/data (the disk is runtime-only) or call a model provider.
# requirements.txt pulls the shared ads-auth package from its public GitHub tag over https, so the
# build needs network access and git, both of which Render's Python build image provides.
set -euo pipefail
cd "$(dirname "$0")/.."
python -m pip install --disable-pip-version-check -r requirements.txt
python -m compileall -q app.py audit_page.py manage.py core scripts
echo "Build complete."
