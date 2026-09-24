"""Vercel entrypoint.

Vercel's FastAPI support looks for an ``app`` in a root ``index.py``; the real
application lives in ``backend/app.py``. Locally, keep running
``uvicorn backend.app:app``.
"""

from backend.app import app  # noqa: F401
