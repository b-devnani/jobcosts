"""Tests for project seeding and blank-friendly project creation."""

import importlib

import pytest

from backend import db


@pytest.fixture()
def seeded_db(tmp_path, monkeypatch):
    monkeypatch.setenv("JOBCOSTS_DB", str(tmp_path / "seed.db"))
    monkeypatch.delenv("JOBCOSTS_SEED", raising=False)  # seeding on (default)
    importlib.reload(db)
    db.init_db()
    return db


def test_seed_parses_and_skips_manual_row():
    projects = db.parse_seed_csv(db.SEED_PATH.read_text())
    names = [p["name"] for p in projects]
    assert "AUA Invest SW" in names
    assert "Other (manual input)" not in names  # the manual sentinel is skipped
    assert len(projects) == 15


def test_seed_converts_dates_to_iso():
    by_name = {p["name"]: p for p in db.parse_seed_csv(db.SEED_PATH.read_text())}
    uc = by_name["UC Bookstore HVAC"]
    assert uc["project_number"] == "25-201-001"
    assert uc["orig_substantial_completion"] == "2025-12-15"
    assert uc["orig_final_completion"] == "2026-03-31"
    assert uc["current_substantial_completion"] == "2026-06-30"
    assert uc["current_final_completion"] == "2026-06-30"
    # A project with only the original dates filled in.
    bronzeville = by_name["PBC Bronzeville Sr Ctr"]
    assert bronzeville["orig_substantial_completion"] == "2027-12-31"
    assert bronzeville["current_substantial_completion"] is None


def test_init_seeds_database(seeded_db):
    projects = seeded_db.list_projects()
    assert len(projects) == 15
    # Blank dates are stored as NULL, not empty strings.
    aua = next(p for p in projects if p["name"] == "AUA Invest SW")
    assert aua["orig_substantial_completion"] is None
    assert aua["project_number"] == "24-201-008"


def test_seed_runs_only_once(seeded_db):
    # Deleting a seeded project and re-initialising must not bring it back.
    projects = seeded_db.list_projects()
    seeded_db.delete_project(projects[0]["id"])
    seeded_db.init_db()
    assert len(seeded_db.list_projects()) == 14


def test_create_project_accepts_only_a_name(seeded_db):
    created = seeded_db.create_project({"name": "Bare Project"})
    assert created["name"] == "Bare Project"
    assert created["project_number"] is None
    assert created["orig_substantial_completion"] is None
    assert created["contract_amount_last_pay_app"] is None


def test_create_project_requires_a_name(seeded_db):
    with pytest.raises(ValueError):
        seeded_db.create_project({"project_number": "x", "name": "   "})
