"""Storage selection (SQLite vs Postgres) and Vercel-specific behaviour."""

import importlib
import os
import subprocess
import sys
from pathlib import Path

import pytest

from backend import db

REPO_ROOT = Path(__file__).resolve().parent.parent


def _reload_with(monkeypatch, **env):
    for var in ("DATABASE_URL", "POSTGRES_URL", "VERCEL", "JOBCOSTS_DB"):
        monkeypatch.delenv(var, raising=False)
    for key, value in env.items():
        monkeypatch.setenv(key, value)
    return importlib.reload(db)


def test_vercel_without_database_uses_tmp_sqlite(monkeypatch):
    mod = _reload_with(monkeypatch, VERCEL="1")
    assert not mod.USE_POSTGRES
    assert mod.DB_PATH == Path("/tmp/jobcosts.db")


def test_local_default_is_sqlite_next_to_the_code(monkeypatch):
    mod = _reload_with(monkeypatch)
    assert not mod.USE_POSTGRES
    assert mod.DB_PATH == Path(mod.__file__).resolve().parent / "jobcosts.db"


@pytest.mark.parametrize("var", ["DATABASE_URL", "POSTGRES_URL"])
def test_database_url_selects_postgres(monkeypatch, var):
    mod = _reload_with(monkeypatch, **{var: "postgresql://u:p@host/db"})
    assert mod.USE_POSTGRES
    assert mod.DATABASE_URL == "postgresql://u:p@host/db"


def test_postgres_url_drops_params_libpq_rejects():
    url = "postgres://u:p@h.pooler.supabase.com:6543/postgres?sslmode=require&supa=base-pooler.x&pgbouncer=true"
    assert db._postgres_url(url) == "postgres://u:p@h.pooler.supabase.com:6543/postgres?sslmode=require"


def test_vercel_entrypoint_exposes_the_app():
    import index
    from backend.app import app

    assert index.app is app


@pytest.mark.skipif(not os.environ.get("TEST_DATABASE_URL"), reason="needs TEST_DATABASE_URL")
def test_concurrent_cold_starts_seed_exactly_once():
    """Several instances starting at once must not seed the list twice."""
    import psycopg

    url = os.environ["TEST_DATABASE_URL"]
    with psycopg.connect(url, autocommit=True) as conn:
        conn.execute("DROP TABLE IF EXISTS projects, app_meta")

    env = {k: v for k, v in os.environ.items() if k not in ("JOBCOSTS_SEED", "POSTGRES_URL")}
    env["DATABASE_URL"] = url
    code = "from backend import db; db.init_db()"
    procs = [
        subprocess.Popen([sys.executable, "-c", code], cwd=REPO_ROOT, env=env)
        for _ in range(6)
    ]
    assert all(p.wait(timeout=60) == 0 for p in procs)

    with psycopg.connect(url) as conn:
        count = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
    assert count == 15
