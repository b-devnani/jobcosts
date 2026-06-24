"""Tiny SQLite-backed store for the projects admins maintain.

A project supplies the four milestone dates (and its name) that the CSV does
not contain.  We keep the schema intentionally small and use the standard
library so the app runs with no extra dependencies.
"""

from __future__ import annotations

import os
import sqlite3
import threading
from pathlib import Path
from typing import Optional

DB_PATH = Path(os.environ.get("JOBCOSTS_DB", Path(__file__).resolve().parent / "jobcosts.db"))

_lock = threading.Lock()

# Column <-> milestone mapping, kept here so the API and converter agree.
DATE_FIELDS = (
    "orig_substantial_completion",
    "orig_final_completion",
    "current_substantial_completion",
    "current_final_completion",
)


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    with _lock, _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS projects (
                id                              INTEGER PRIMARY KEY AUTOINCREMENT,
                name                            TEXT    NOT NULL,
                orig_substantial_completion     TEXT,
                orig_final_completion           TEXT,
                current_substantial_completion  TEXT,
                current_final_completion        TEXT,
                created_at                      TEXT    NOT NULL DEFAULT (datetime('now')),
                updated_at                      TEXT    NOT NULL DEFAULT (datetime('now'))
            )
            """
        )


def _row_to_dict(row: sqlite3.Row) -> dict:
    return {k: row[k] for k in row.keys()}


def _fetch(conn: sqlite3.Connection, project_id: int) -> Optional[dict]:
    """Read a single project using an already-open connection."""
    row = conn.execute(
        "SELECT * FROM projects WHERE id = ?", (project_id,)
    ).fetchone()
    return _row_to_dict(row) if row else None


def list_projects() -> list[dict]:
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM projects ORDER BY name COLLATE NOCASE"
        ).fetchall()
        return [_row_to_dict(r) for r in rows]


def get_project(project_id: int) -> Optional[dict]:
    with _connect() as conn:
        return _fetch(conn, project_id)


def create_project(data: dict) -> dict:
    name = (data.get("name") or "").strip()
    if not name:
        raise ValueError("Project name is required.")
    values = [name] + [_norm_date(data.get(f)) for f in DATE_FIELDS]
    with _lock, _connect() as conn:
        cur = conn.execute(
            f"""INSERT INTO projects (name, {", ".join(DATE_FIELDS)})
                VALUES (?, ?, ?, ?, ?)""",
            values,
        )
        return _fetch(conn, cur.lastrowid)


def update_project(project_id: int, data: dict) -> Optional[dict]:
    # The whole read-modify-write runs under the lock so concurrent updates to
    # the same project cannot clobber each other with stale field values.
    with _lock, _connect() as conn:
        existing = _fetch(conn, project_id)
        if existing is None:
            return None
        name = (data.get("name") or existing["name"]).strip()
        if not name:
            raise ValueError("Project name cannot be empty.")
        values = [name] + [
            _norm_date(data.get(f, existing[f])) for f in DATE_FIELDS
        ] + [project_id]
        conn.execute(
            f"""UPDATE projects
                   SET name = ?,
                       {", ".join(f + " = ?" for f in DATE_FIELDS)},
                       updated_at = datetime('now')
                 WHERE id = ?""",
            values,
        )
        return _fetch(conn, project_id)


def delete_project(project_id: int) -> bool:
    with _lock, _connect() as conn:
        cur = conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))
        return cur.rowcount > 0


def _norm_date(value) -> Optional[str]:
    """Store dates as ISO yyyy-mm-dd strings (or NULL)."""
    if value in (None, ""):
        return None
    return str(value).strip()
