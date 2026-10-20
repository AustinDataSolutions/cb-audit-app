#!/usr/bin/env python
"""Startup checks run before Streamlit (see scripts/start.sh). Exit non-zero to refuse to serve.

The checks live in `ads_auth.preflight`; this wrapper supplies the app's settings.
Never prints secrets. Safe to run repeatedly.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ads_auth import preflight  # noqa: E402
from core.config import load_settings  # noqa: E402


def run(environ=None, *, out=print) -> int:
    environ = os.environ if environ is None else environ
    dotenv_path = None if environ.get("RENDER", "").lower() == "true" else ".env"
    return preflight.run(lambda: load_settings(environ, dotenv_path=dotenv_path), out=out)


if __name__ == "__main__":
    sys.exit(run())
