#!/usr/bin/env python
"""Operator CLI: accounts, backups, and usage events, from `ads_auth.cli` with this app's settings.

Run from the repository root (locally) or the Render Shell. Passwords are never accepted as
arguments; they are prompted for or generated and shown once. `python manage.py --help` lists the commands.
"""
import sys

from ads_auth import cli
from core.config import load_settings

if __name__ == "__main__":
    sys.exit(cli.main(load_settings=load_settings))
