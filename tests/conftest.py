"""Shared fixtures.

Tests that use ``storage`` run once per storage engine: SQLite always, and
Postgres when ``TEST_DATABASE_URL`` points at a disposable database (its
``projects`` and ``app_meta`` tables are dropped before each test).
"""

import os

import pytest

TEST_DATABASE_URL = os.environ.get("TEST_DATABASE_URL")


@pytest.fixture(params=["sqlite", "postgres"])
def storage(request, tmp_path, monkeypatch):
    """Point ``backend.db`` (once reloaded) at a fresh, empty database."""
    for var in ("DATABASE_URL", "POSTGRES_URL", "VERCEL", "JOBCOSTS_DB"):
        monkeypatch.delenv(var, raising=False)
    if request.param == "sqlite":
        monkeypatch.setenv("JOBCOSTS_DB", str(tmp_path / "test.db"))
    else:
        if not TEST_DATABASE_URL:
            pytest.skip("set TEST_DATABASE_URL to run the Postgres tests")
        import psycopg

        with psycopg.connect(TEST_DATABASE_URL, autocommit=True) as conn:
            conn.execute("DROP TABLE IF EXISTS projects, app_meta")
        monkeypatch.setenv("DATABASE_URL", TEST_DATABASE_URL)
    return request.param
