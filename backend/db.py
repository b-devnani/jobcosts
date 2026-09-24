"""Tiny store for the projects admins maintain.

A project supplies the milestone dates, the last-pay-app figures and its name —
the "remaining info" the CSV does not contain. Everything except the name is
optional, so admins can seed a project now and fill in the rest later.

Storage is SQLite by default. When ``DATABASE_URL`` (or ``POSTGRES_URL``) is set
-- e.g. a Neon or Supabase database attached to a Vercel project, where the
filesystem is not persistent -- the same schema lives in Postgres instead.

On first run an empty database is seeded from ``seed/projects_seed.csv`` so the
dropdown is populated out of the box.
"""

from __future__ import annotations

import csv
import io
import os
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator, Optional
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

DATABASE_URL = os.environ.get("DATABASE_URL") or os.environ.get("POSTGRES_URL")
USE_POSTGRES = bool(DATABASE_URL)


def _default_db_path() -> Path:
    # On Vercel only /tmp is writable (and it is per-instance, not persistent).
    if os.environ.get("VERCEL"):
        return Path("/tmp/jobcosts.db")
    return Path(__file__).resolve().parent / "jobcosts.db"


DB_PATH = Path(os.environ.get("JOBCOSTS_DB") or _default_db_path())
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


# SQL that differs between the two engines.
_ID_COLUMN = "SERIAL PRIMARY KEY" if USE_POSTGRES else "INTEGER PRIMARY KEY AUTOINCREMENT"
_NOW = (
    "to_char(now() AT TIME ZONE 'UTC', 'YYYY-MM-DD HH24:MI:SS')"
    if USE_POSTGRES
    else "datetime('now')"
)


def _postgres_url(url: str) -> str:
    """Drop query parameters some hosted integrations append that libpq
    rejects (Supabase's ``supa=``, Prisma's ``pgbouncer=``)."""
    parts = urlsplit(url)
    query = [
        (k, v)
        for k, v in parse_qsl(parts.query, keep_blank_values=True)
        if k not in {"supa", "pgbouncer"}
    ]
    return urlunsplit(parts._replace(query=urlencode(query)))


class _Conn:
    """Runs the module's SQL (written with ``?`` placeholders) on either engine."""

    def __init__(self, raw):
        self.raw = raw

    def execute(self, sql: str, params=()):
        if USE_POSTGRES:
            sql = sql.replace("?", "%s")
        return self.raw.execute(sql, params)


@contextmanager
def _connection() -> Iterator[_Conn]:
    """Open a connection, commit on success / roll back on error, always close."""
    if USE_POSTGRES:
        import psycopg
        from psycopg.rows import dict_row

        # prepare_threshold=None: server-side prepared statements break behind
        # transaction-mode poolers (Neon/Supabase pooled URLs).
        raw = psycopg.connect(
            _postgres_url(DATABASE_URL),
            row_factory=dict_row,
            prepare_threshold=None,
            connect_timeout=10,
        )
    else:
        raw = sqlite3.connect(DB_PATH)
        raw.row_factory = sqlite3.Row
        raw.execute("PRAGMA foreign_keys = ON")
    try:
        yield _Conn(raw)
        raw.commit()
    except BaseException:
        raw.rollback()
        raise
    finally:
        raw.close()


def init_db() -> None:
    if not USE_POSTGRES:
        # Create the directory for the database file if it does not exist yet
        # (e.g. a freshly mounted persistent disk at /data).
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with _lock, _connection() as conn:
        if USE_POSTGRES:
            # Serialise schema creation + seeding across concurrently starting
            # instances; released when this transaction commits.
            conn.execute("SELECT pg_advisory_xact_lock(724301)")
        conn.execute(
            f"""
            CREATE TABLE IF NOT EXISTS projects (
                id                              {_ID_COLUMN},
                project_number                  TEXT,
                name                            TEXT    NOT NULL,
                orig_substantial_completion     TEXT,
                orig_final_completion           TEXT,
                current_substantial_completion  TEXT,
                current_final_completion        TEXT,
                contract_amount_last_pay_app    TEXT,
                month_last_pay_app              TEXT,
                created_at                      TEXT    NOT NULL DEFAULT ({_NOW}),
                updated_at                      TEXT    NOT NULL DEFAULT ({_NOW})
            )
            """
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS app_meta (key TEXT PRIMARY KEY, value TEXT)"
        )
        _migrate_columns(conn)
        _seed_if_needed(conn)


def _migrate_columns(conn: _Conn) -> None:
    """Add any newer columns to a database created by an older schema."""
    if USE_POSTGRES:
        rows = conn.execute(
            "SELECT column_name AS name FROM information_schema.columns "
            "WHERE table_schema = current_schema() AND table_name = 'projects'"
        )
    else:
        rows = conn.execute("PRAGMA table_info(projects)")
    existing = {row["name"] for row in rows}
    for col in EDITABLE_FIELDS:
        if col not in existing:
            conn.execute(f"ALTER TABLE projects ADD COLUMN {col} TEXT")


# --------------------------------------------------------------------------- #
# Reads
# --------------------------------------------------------------------------- #
def _row_to_dict(row) -> dict:
    return {k: row[k] for k in row.keys()}


def _fetch(conn: _Conn, project_id: int) -> Optional[dict]:
    """Read a single project using an already-open connection."""
    row = conn.execute(
        "SELECT * FROM projects WHERE id = ?", (project_id,)
    ).fetchone()
    return _row_to_dict(row) if row else None


def list_projects() -> list[dict]:
    with _connection() as conn:
        rows = conn.execute("SELECT * FROM projects ORDER BY lower(name), id").fetchall()
        return [_row_to_dict(r) for r in rows]


def get_project(project_id: int) -> Optional[dict]:
    with _connection() as conn:
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


def _insert(conn: _Conn, data: dict) -> dict:
    name = (data.get("name") or "").strip()
    if not name:
        raise ValueError("Project name is required.")
    values = [name if col == "name" else _norm(data.get(col)) for col in EDITABLE_FIELDS]
    placeholders = ", ".join("?" * len(EDITABLE_FIELDS))
    sql = f"INSERT INTO projects ({', '.join(EDITABLE_FIELDS)}) VALUES ({placeholders})"
    if USE_POSTGRES:
        new_id = conn.execute(sql + " RETURNING id", values).fetchone()["id"]
    else:
        new_id = conn.execute(sql, values).lastrowid
    return _fetch(conn, new_id)


def create_project(data: dict) -> dict:
    with _lock, _connection() as conn:
        return _insert(conn, data)


def update_project(project_id: int, data: dict) -> Optional[dict]:
    # The whole read-modify-write runs under the lock so concurrent updates to
    # the same project cannot clobber each other with stale field values.
    with _lock, _connection() as conn:
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
                       updated_at = {_NOW}
                 WHERE id = ?""",
            values,
        )
        return _fetch(conn, project_id)


def delete_project(project_id: int) -> bool:
    with _lock, _connection() as conn:
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


def _seed_if_needed(conn: _Conn) -> None:
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
    count = conn.execute("SELECT COUNT(*) AS n FROM projects").fetchone()["n"]
    if seeding_enabled and count == 0 and SEED_PATH.exists():
        for project in parse_seed_csv(SEED_PATH.read_text()):
            _insert(conn, project)
    # Record that seeding has run so admin deletions are never re-seeded.
    conn.execute(
        "INSERT INTO app_meta (key, value) VALUES ('seeded', '1') "
        "ON CONFLICT (key) DO NOTHING"
    )
