"""
Vercel serverless entrypoint.

The `functions` block in vercel.json must reference files under `api/`.
This module re-exports the Flask `app` from `application.py` at the repo root.
"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parent.parent
_root_str = str(_root)
if _root_str not in sys.path:
    sys.path.insert(0, _root_str)

from application import app  # noqa: E402
