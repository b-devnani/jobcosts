"""Tiny SQLite-backed store for the projects admins maintain.

A project supplies the milestone dates, the last-pay-app figures and its name —
the "remaining info" the CSV does not contain. Everything except the name is
optional, so admins can seed a project now and fill in the rest later.

On first run an empty database is seeded from ``seed/projects_seed.csv`` so the
dropdown is populated out of the box.
"""

from __future__ import annotations

import csv
import io
import os
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

DB_PATH = Path(os.environ.get("JOBCOSTS_DB", Path(__file__).resolve().parent / "jobcosts.db"))
SEED_PATH = Path(__file__).resolve().parent / "seed" / "projects_seed.csv"

_lock = threading.Lock()

# Editable project columns, in display order. ``name`` is required; the rest may
# be left blank and filled in later by an admin.
EDITABLE_FIELDS = (
    "project_number",
    "name",
    "orig_substantial_completion",
    "orig_final_completion",
    "current_substantial_completion",
    "current_final_completion",
    "contract_amount_last_pay_app",
    "month_last_pay_app",
)

# Fields the converter treats as dates (stored as ISO yyyy-mm-dd strings).
DATE_FIELDS = (
    "orig_substantial_completion",
    "orig_final_completion",
    "current_substantial_completion",
    "current_final_completion",
    "month_last_pay_app",
)


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    # Create the directory for the database file if it does not exist yet
    # (e.g. a freshly mounted persistent disk at /data).
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with _lock, _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS projects (
                id                              INTEGER PRIMARY KEY AUTOINCREMENT,
                project_number                  TEXT,
                name                            TEXT    NOT NULL,
                orig_substantial_completion     TEXT,
                orig_final_completion           TEXT,
                current_substantial_completion  TEXT,
                current_final_completion        TEXT,
                contract_amount_last_pay_app    TEXT,
                month_last_pay_app              TEXT,
                created_at                      TEXT    NOT NULL DEFAULT (datetime('now')),
                updated_at                      TEXT    NOT NULL DEFAULT (datetime('now'))
            )
            """
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS app_meta (key TEXT PRIMARY KEY, value TEXT)"
        )
        _migrate_columns(conn)
        _seed_if_needed(conn)


def _migrate_columns(conn: sqlite3.Connection) -> None:
    """Add any newer columns to a database created by an older schema."""
    existing = {row["name"] for row in conn.execute("PRAGMA table_info(projects)")}
    for col in EDITABLE_FIELDS:
        if col not in existing:
            conn.execute(f"ALTER TABLE projects ADD COLUMN {col} TEXT")


# --------------------------------------------------------------------------- #
# Reads
# --------------------------------------------------------------------------- #
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


# --------------------------------------------------------------------------- #
# Writes
# --------------------------------------------------------------------------- #
def _norm(value) -> Optional[str]:
    """Blank values become NULL; everything else is stored as a trimmed string."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _insert(conn: sqlite3.Connection, data: dict) -> dict:
    name = (data.get("name") or "").strip()
    if not name:
        raise ValueError("Project name is required.")
    values = [name if col == "name" else _norm(data.get(col)) for col in EDITABLE_FIELDS]
    placeholders = ", ".join("?" * len(EDITABLE_FIELDS))
    cur = conn.execute(
        f"INSERT INTO projects ({', '.join(EDITABLE_FIELDS)}) VALUES ({placeholders})",
        values,
    )
    return _fetch(conn, cur.lastrowid)


def create_project(data: dict) -> dict:
    with _lock, _connect() as conn:
        return _insert(conn, data)


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
        values = [
            name if col == "name" else _norm(data.get(col, existing[col]))
            for col in EDITABLE_FIELDS
        ] + [project_id]
        conn.execute(
            f"""UPDATE projects
                   SET {", ".join(col + " = ?" for col in EDITABLE_FIELDS)},
                       updated_at = datetime('now')
                 WHERE id = ?""",
            values,
        )
        return _fetch(conn, project_id)


def delete_project(project_id: int) -> bool:
    with _lock, _connect() as conn:
        cur = conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))
        return cur.rowcount > 0


# --------------------------------------------------------------------------- #
# Seeding
# --------------------------------------------------------------------------- #
def _iso_date(value: str) -> Optional[str]:
    """Convert a m/d/Y seed date into ISO yyyy-mm-dd (so date inputs display it)."""
    text = (value or "").strip()
    if not text:
        return None
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text  # leave anything unexpected as-is rather than dropping it


def parse_seed_csv(content: str) -> list[dict]:
    """Parse the Company-Home style export into project dicts.

    Expected columns: Project Number, Name, Original Substantial Completion,
    Original Final Completion, Current Substantial Completion, Current Final
    Completion, Contract amount on last pay app, Month of last pay app.
    The ``manual`` sentinel row (the built-in "manual input" option) is skipped.
    """
    reader = csv.reader(io.StringIO(content))
    rows = list(reader)
    projects: list[dict] = []
    for raw in rows[1:]:  # skip header
        if not any(cell.strip() for cell in raw):
            continue
        number = raw[0].strip() if len(raw) > 0 else ""
        name = raw[1].strip() if len(raw) > 1 else ""
        if not name or number.lower() == "manual":
            continue  # the manual-input sentinel is already a dropdown option
        projects.append(
            {
                "project_number": number or None,
                "name": name,
                "orig_substantial_completion": _iso_date(raw[2] if len(raw) > 2 else ""),
                "orig_final_completion": _iso_date(raw[3] if len(raw) > 3 else ""),
                "current_substantial_completion": _iso_date(raw[4] if len(raw) > 4 else ""),
                "current_final_completion": _iso_date(raw[5] if len(raw) > 5 else ""),
                "contract_amount_last_pay_app": (raw[6].strip() if len(raw) > 6 else "") or None,
                "month_last_pay_app": _iso_date(raw[7] if len(raw) > 7 else ""),
            }
        )
    return projects


def _seed_if_needed(conn: sqlite3.Connection) -> None:
    """Populate an empty database from the bundled seed file, exactly once.

    Set ``JOBCOSTS_SEED=0`` to skip seeding (used by tests that want an empty
    database).
    """
    already = conn.execute(
        "SELECT value FROM app_meta WHERE key = 'seeded'"
    ).fetchone()
    if already:
        return
    seeding_enabled = os.environ.get("JOBCOSTS_SEED", "1") != "0"
    count = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
    if seeding_enabled and count == 0 and SEED_PATH.exists():
        for project in parse_seed_csv(SEED_PATH.read_text()):
            _insert(conn, project)
    # Record that seeding has run so admin deletions are never re-seeded.
    conn.execute(
        "INSERT OR REPLACE INTO app_meta (key, value) VALUES ('seeded', '1')"
    )
