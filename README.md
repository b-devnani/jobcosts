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
| file name → title cell `A1` | Project name |

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
| `PORT` / `HOST` | `8000` / `127.0.0.1` | Bind address for `run.sh` |

---

## Project layout

```
backend/
  app.py            FastAPI app: projects CRUD + /api/generate + static
  converter.py      CSV -> XLSX core logic (steps 6-9 + milestone dates)
  db.py             SQLite project store (stdlib sqlite3)
  template/         The Job Cost Projection template workbook
  static/           Frontend (index.html, app.js, styles.css)
tests/
  test_converter.py Unit tests for the conversion
  test_api.py       API integration tests
  sample/           Sample budget-detail CSV
```

## Tests

```bash
python3 -m pytest -q
```

## Notes & limits

* Admin auth is a single shared password (sent as `X-Admin-Password`). It gates
  project writes; swap in real auth before exposing publicly.
* The template holds up to **146** data rows; larger CSVs are rejected with a
  clear message.
* Date cells are written with the template's `mm-dd-yy` format.
