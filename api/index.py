"""Vercel serverless entry: routes all traffic here via vercel.json rewrites."""
from __future__ import annotations

import sys
from pathlib import Path

_root = str(Path(__file__).resolve().parent.parent)
if _root not in sys.path:
    sys.path.insert(0, _root)

from factory import app  # noqa: E402
