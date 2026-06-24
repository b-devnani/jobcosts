# Job Cost Projection Tool

A small web tool that turns a Procore **budget-detail CSV** into a filled-in
**Job Cost Projection** Excel workbook.

A user uploads the CSV, picks a project from a dropdown, and downloads the
workbook. An admin maintains the project list (name + four milestone dates) in
a data grid.

---

## What it does

The conversion reproduces the manual Excel procedure:

| Step | Manual instruction | What the tool does |
|------|--------------------|--------------------|
| 6 | Delete columns A, D and E of the CSV | Drops *Cost Code Tier 1*, *Budget Code*, *Budget Code Description* |
| 7 | Keep columns A (cost code) – I (job to date) | Keeps the remaining 9 columns in template order |
| 8 | Paste as values into the template | Writes them into the `Job Cost` sheet starting at row 8 |
| 9 | Delete rows below the data | Removes unused template rows and repoints the totals / header formulas |

The **"remaining info"** that is not in the CSV comes from the selected
project:

| Template cell | Field |
|---------------|-------|
| `I3` | Original Substantial Comp Date |
| `I4` | Original Final Completion Date |
| `K3` | Current Substantial Comp Date |
| `K4` | Current Final Comp Date |
| `F3` | Contract amount on last pay app |
| `G3` | Month of last pay app |
| file name → title cell `A1` | Project name |

The project list is **seeded on first run** from `backend/seed/projects_seed.csv`
(a Company-Home style export of project numbers, names and any known dates), so
the dropdown is populated out of the box. Every field except the name is
optional — admins fill in the rest later. Seeding runs once; deleting a seeded
project does not bring it back.

> Procore placeholder rows (cost code `None`) are skipped automatically. The
> contract-amount cells (`C3`–`C5`) are template formulas that recalculate from
> the pasted data when the workbook is opened in Excel.

---

## Run it

```bash
pip install -r requirements.txt
./run.sh
# or:  python3 -m uvicorn backend.app:app --reload
```

Then open <http://127.0.0.1:8000>.

* **Generate** tab — choose a project (or "Manual entry"), drop in the CSV,
  click **Generate & Download**.
* **Admin** tab — sign in with the admin password, then add / edit / delete
  projects in the grid.

### Configuration

| Env var | Default | Purpose |
|---------|---------|---------|
| `ADMIN_PASSWORD` | `admin` | Password protecting project writes |
| `JOBCOSTS_DB` | `backend/jobcosts.db` | SQLite database path |
| `JOBCOSTS_SEED` | `1` | Set to `0` to skip seeding the project list |
| `PORT` / `HOST` | `8000` / `127.0.0.1` | Bind address for `run.sh` |

---

## Project layout

```
backend/
  app.py            FastAPI app: projects CRUD + /api/generate + static
  converter.py      CSV -> XLSX core logic (steps 6-9 + milestone dates)
  db.py             SQLite project store (stdlib sqlite3)
  template/         The Job Cost Projection template workbook
  seed/             projects_seed.csv used to seed the dropdown on first run
  static/           Frontend (index.html, app.js, styles.css)
tests/
  test_converter.py Unit tests for the conversion
  test_api.py       API integration tests
  sample/           Sample budget-detail CSV
```

## Deploy (host it online)

The app is one FastAPI process that serves both the API and the frontend, plus
a SQLite file for the project list. Three things matter when hosting:

1. **Bind to `0.0.0.0` and the platform port:**
   `uvicorn backend.app:app --host 0.0.0.0 --port $PORT`
2. **Persist the SQLite DB** on a disk/volume and point `JOBCOSTS_DB` at it
   (otherwise admin edits reset on each restart; seeded projects still return).
3. **Set a strong `ADMIN_PASSWORD`** and use HTTPS (the admin password travels
   in a request header). Managed hosts provide HTTPS automatically.

### Render (one click)

This repo ships a [`render.yaml`](render.yaml) blueprint. In Render: **New →
Blueprint → connect this repo**. It provisions a web service with a 1 GB
persistent disk at `/data`, sets `JOBCOSTS_DB=/data/jobcosts.db`, and generates
a strong `ADMIN_PASSWORD` (read it in the dashboard's Environment tab). A health
check is exposed at `/healthz`.

> The persistent disk needs a paid (starter+) instance. On the free tier, remove
> the `disk:` block and `JOBCOSTS_DB` — the DB then re-seeds on every restart.

### Anywhere with Docker (Fly.io, Railway, a VPS, …)

A [`Dockerfile`](Dockerfile) is included:

```bash
docker build -t jobcosts .
docker run -p 8000:8000 -e ADMIN_PASSWORD=change-me -v jobcosts-data:/data jobcosts
```

The named volume `jobcosts-data` keeps the project database across restarts.

## Tests

```bash
python3 -m pytest -q
```

## Notes & limits

* Admin auth is a single shared password (sent as `X-Admin-Password`). It gates
  project writes; swap in real auth before exposing publicly. The app logs a
  warning at startup if `ADMIN_PASSWORD` is left at the insecure default.
* The template holds up to **146** data rows; larger CSVs are rejected with a
  clear message. Uploads are also capped at **10 MB**.
* The uploaded CSV's key column headers are validated, so a non-Procore file is
  rejected rather than silently mapped into the wrong template columns.
* Date cells are written with the template's `mm-dd-yy` format.
