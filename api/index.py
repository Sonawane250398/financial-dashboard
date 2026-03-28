"""
Vercel serverless entry for the Dash app.

Vercel's Flask integration expects a Flask WSGI application bound to the name `app`.
The Dash instance lives in `app.py`; its underlying Flask server is exported as `server`.
"""
from __future__ import annotations

import os
import sys

# Project root must be on sys.path so `import app` resolves to `app.py` at the repo root.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

# Relative paths in `app.py` (e.g. financial_sample.xlsx) assume the process cwd is the project root.
if os.getcwd() != _ROOT:
    os.chdir(_ROOT)

from app import server as app  # noqa: E402  — `app` is the Flask WSGI app Vercel invokes

__all__ = ["app"]
