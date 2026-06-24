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

The workbook downloads as **`<Project Name> Job Costs MMDDYY.xlsx`** (today's
date). The sheet carries the BURLING logo and a themed title/summary header; the
table itself is left as the template ships it apart from column borders, a grey
header row, blue **Committed Costs** and green **Estimated Cost at Completion**
columns, and a bold totals row (top border + double bottom). It prints
landscape with narrow margins, fits all columns to one page wide, repeats rows
1–7 on every page, and shows a "Page x of x" footer. Values, formulas and the
currency/date number formats are never altered.

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

* **Generate** tab — choose a project, then click **Upload CSV** and pick the
  file. The workbook builds and downloads automatically (no extra clicks).
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

### Render (free, one click)

This repo ships a [`render.yaml`](render.yaml) blueprint configured for Render's
**free** tier. In Render: **New → Blueprint → connect this repo**. It provisions
a free web service, generates a strong `ADMIN_PASSWORD` (read it in the
dashboard's Environment tab), and exposes a `/healthz` check. You get a
`https://…onrender.com` URL with HTTPS, at $0.

Free-tier trade-offs:

* **No persistent disk** — the SQLite DB is ephemeral, so the dropdown re-seeds
  the 15 projects on every restart and admin edits are not retained. The
  conversion and the seeded dropdown work fine; type any one-off dates in the
  Generate form (no saved project needed).
* **Sleeps when idle** — the first request after ~15 min cold-starts (~30–60s).

To make admin edits **persist**, upgrade to a paid (starter+) instance and add a
disk — see the comment block at the top of `render.yaml`.

**Want the seeded dates to stick for free?** Edit
`backend/seed/projects_seed.csv` with the real completion dates and commit — the
dropdown then re-seeds with the correct data on every boot, no disk required.

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
