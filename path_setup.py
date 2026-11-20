"""Ensure the repository root is on ``sys.path`` (shared by CLI entrypoints)."""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent


def ensure_repo_root_on_sys_path() -> Path:
    """
    Insert this repository's root onto ``sys.path`` once.

    Call before importing ``src.*`` from scripts that are not installed as a package
    (e.g. ``main.py``). Tests rely on ``pytest.ini`` ``pythonpath`` instead.
    """
    root = str(REPO_ROOT)
    if root not in sys.path:
        sys.path.insert(0, root)
    return REPO_ROOT
