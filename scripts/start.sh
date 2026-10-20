#!/usr/bin/env bash
# Production start: refuse to serve unless preflight passes, then run Streamlit through the router.
set -euo pipefail
cd "$(dirname "$0")/.."
python scripts/preflight.py
exec streamlit run app.py --server.address 0.0.0.0 --server.port "${PORT:-8501}" --server.headless true
