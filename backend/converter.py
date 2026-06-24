"""Core logic: turn a Procore-style budget-detail CSV into a filled-in
Job Cost Projection workbook.

The transformation mirrors the manual Excel procedure documented in the
project brief (steps 6-9):

    6. Delete columns A, D and E of the exported CSV.
    7. Keep columns A (cost code) through I (job to date cost).
    8. Paste those columns, as values, into the template starting at row 8.
    9. Delete the unused template rows below the pasted data.

On top of the pasted data we also stamp the four project milestone dates and
name the workbook after the project, which is how the template's title cell
(an array formula reading the file name) shows the project name.
"""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

TEMPLATE_PATH = Path(__file__).resolve().parent / "template" / "Job_Cost_Projection_Template.xlsx"

SHEET_NAME = "Job Cost"

# Row of the table header inside the template; data starts on the next row.
HEADER_ROW = 7
FIRST_DATA_ROW = 8
# Last data row the pristine template ships with (row 154 holds the totals).
LAST_TEMPLATE_DATA_ROW = 153
TOTALS_ROW = 154

# Header columns that carry currency values and must be stored as numbers so
# the workbook's own formulas keep working.
NUMERIC_OUTPUT_COLUMNS = ("C", "D", "E", "F", "G", "H", "I")

DATE_FORMAT = "mm-dd-yy"

# Indexes (0-based) of the CSV columns that survive "delete columns A, D, E"
# and form template columns A..I.  Order matters: it is the template order.
#
#   CSV layout (0-based):
#     0 Cost Code Tier 1      <- deleted (col A)
#     1 Cost Code Tier 2      -> template A  Cost Code/Description
#     2 Cost Type             -> template B  CAT
#     3 Budget Code           <- deleted (col D)
#     4 Budget Code Desc.     <- deleted (col E)
#     5 Original Budget       -> template C
#     6 Budget Modifications  -> template D
#     7 Approved COs          -> template E
#     8 Revised Budget        -> template F
#     9 Committed Costs       -> template G
#     10 Direct Cost          -> template H
#     11 Job to date Costs    -> template I
CSV_COLUMN_ORDER = [1, 2, 5, 6, 7, 8, 9, 10, 11]
# Which of the kept columns are text vs. numeric (by template column letter).
TEXT_COLUMNS = {"A", "B"}

# A handful of header names we verify so a non-Procore CSV that merely has 12+
# columns cannot map the wrong fields into the template and silently corrupt it.
EXPECTED_HEADERS = {
    1: "cost code tier 2",
    2: "cost type",
    5: "original budget amount",
    11: "job to date costs",
}


class ConversionError(ValueError):
    """Raised when the uploaded CSV does not look like a budget-detail export."""


@dataclass
class ProjectInfo:
    """The "remaining info" that does not come from the CSV.

    ``name`` becomes the workbook file name (and therefore the title cell).
    The four dates land in the template's milestone cells.
    """

    name: str
    orig_substantial_completion: date | None = None
    orig_final_completion: date | None = None
    current_substantial_completion: date | None = None
    current_final_completion: date | None = None


def _to_number(raw: str) -> float | str | None:
    """Parse a CSV money cell into a float, treating blanks as ``None``."""
    if raw is None:
        return None
    text = raw.strip().strip('"').replace(",", "").replace("$", "")
    if text == "" or text.lower() == "none":
        return None
    try:
        return float(text)
    except ValueError:
        # Non-numeric content in a money column -> leave as-is so nothing is lost.
        return raw.strip()


def _clean_text(raw: str | None) -> str:
    return (raw or "").strip()


def parse_budget_csv(content: str | bytes) -> list[list]:
    """Apply steps 6-7: drop columns A/D/E and keep A..I for each real row.

    Returns a list of 9-element rows in template column order (A..I).  Rows
    whose cost code is blank or the Procore ``None`` placeholder are skipped.
    """
    if isinstance(content, bytes):
        content = content.decode("utf-8-sig")

    reader = csv.reader(io.StringIO(content))
    rows = list(reader)
    if not rows:
        raise ConversionError("The CSV file is empty.")

    header = [c.strip().strip('"') for c in rows[0]]
    if len(header) < 12:
        raise ConversionError(
            "Unexpected CSV format: expected a Procore budget-detail export with "
            f"at least 12 columns, found {len(header)}."
        )
    mismatches = [
        f"column {idx} should be '{name}' but is "
        f"'{header[idx] if idx < len(header) else ''}'"
        for idx, name in EXPECTED_HEADERS.items()
        if idx >= len(header) or header[idx].strip().lower() != name
    ]
    if mismatches:
        raise ConversionError(
            "Unexpected CSV format (is this a Procore budget-detail export?): "
            + "; ".join(mismatches)
        )

    out: list[list] = []
    for raw_row in rows[1:]:
        if not any(cell.strip() for cell in raw_row):
            continue  # fully blank line
        if len(raw_row) <= max(CSV_COLUMN_ORDER):
            continue  # malformed / short line

        cost_code = _clean_text(raw_row[1])
        if cost_code == "" or cost_code.lower() == "none":
            continue  # Procore placeholder / subtotal row

        record = []
        for src_idx, col_letter in zip(CSV_COLUMN_ORDER, "ABCDEFGHI"):
            if col_letter in TEXT_COLUMNS:
                record.append(_clean_text(raw_row[src_idx]))
            else:
                record.append(_to_number(raw_row[src_idx]))
        out.append(record)

    if not out:
        raise ConversionError("No data rows were found in the CSV.")
    return out


def _coerce_date(value) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
        raise ConversionError(f"Could not understand date value: {value!r}")
    raise ConversionError(f"Unsupported date value: {value!r}")


def _set_date(ws: Worksheet, coord: str, value: date | None) -> None:
    if value is None:
        return
    cell = ws[coord]
    cell.value = value
    cell.number_format = DATE_FORMAT


def _rewrite_formulas_for_totals(ws: Worksheet, data_rows: int) -> None:
    """After the surplus rows are deleted, repoint the formulas that referenced
    the original totals/summary rows so the workbook still recalculates."""
    totals_row = FIRST_DATA_ROW + data_rows  # new position of the totals row
    last_data = totals_row - 1

    # Totals row: SUM(col8:col<last_data>) for every numeric/derived column.
    for col in "CDEFGHIJKL":
        ws[f"{col}{totals_row}"] = f"=SUM({col}{FIRST_DATA_ROW}:{col}{last_data})"

    # Header contract-amount cells reference the totals row.
    ws["C3"] = f"=+C{totals_row}"   # Original Contract Amount  (Original Budget total)
    ws["C4"] = f"=+E{totals_row}"   # Approved PCCO's           (Approved COs total)
    ws["C5"] = f"=+F{totals_row}"   # Current Contract Amount   (Revised Budget total)

    # Summary block below the totals (PROJECT FEE ... Pending PCCO Fee) shifts up
    # by the number of deleted rows. Repoint its one internal formula.
    shift = LAST_TEMPLATE_DATA_ROW - last_data  # rows removed
    if shift:
        new_l159 = 159 - shift
        new_l160 = 160 - shift
        new_l161 = 161 - shift
        ws[f"L{new_l161}"] = f"=+L{new_l159}-L{new_l160}"


def build_workbook(csv_rows: Sequence[Sequence], project: ProjectInfo):
    """Return an openpyxl workbook: the template with data + dates filled in."""
    try:
        wb = load_workbook(TEMPLATE_PATH)
    except FileNotFoundError:
        raise ConversionError(f"Template workbook not found at {TEMPLATE_PATH}.")
    except Exception as exc:  # corrupt / unreadable workbook
        raise ConversionError(f"Template workbook could not be read: {exc}")
    if SHEET_NAME not in wb.sheetnames:
        raise ConversionError(
            f"Template is missing the required '{SHEET_NAME}' sheet."
        )
    ws = wb[SHEET_NAME]

    n = len(csv_rows)
    capacity = LAST_TEMPLATE_DATA_ROW - FIRST_DATA_ROW + 1  # 146 rows
    if n > capacity:
        raise ConversionError(
            f"The CSV has {n} rows but the template only has room for {capacity}."
        )

    # Step 8: paste columns A..I as values into the data rows.
    for i, record in enumerate(csv_rows):
        row = FIRST_DATA_ROW + i
        for col_letter, value in zip("ABCDEFGHI", record):
            ws[f"{col_letter}{row}"] = value

    # Step 9: delete the unused template rows below the pasted data, then fix
    # up the formulas that pointed at the (now moved) totals/summary rows.
    surplus_start = FIRST_DATA_ROW + n
    surplus_count = LAST_TEMPLATE_DATA_ROW - surplus_start + 1
    if surplus_count > 0:
        ws.delete_rows(surplus_start, surplus_count)
    _rewrite_formulas_for_totals(ws, n)

    # Stamp the four milestone dates from the project record.
    _set_date(ws, "I3", project.orig_substantial_completion)
    _set_date(ws, "I4", project.orig_final_completion)
    _set_date(ws, "K3", project.current_substantial_completion)
    _set_date(ws, "K4", project.current_final_completion)

    return wb


def safe_filename(name: str, max_length: int = 120) -> str:
    """Make a project name safe (and not too long) to use as a download name."""
    keep = "".join(c if c.isalnum() or c in " -_." else "_" for c in name).strip()
    base = keep or "Job Cost Projection"
    if len(base) + len(".xlsx") > max_length:
        base = base[: max_length - len(".xlsx")].strip()
    return base + ".xlsx"


def convert_csv_to_workbook_bytes(
    csv_content: str | bytes,
    name: str,
    orig_substantial_completion=None,
    orig_final_completion=None,
    current_substantial_completion=None,
    current_final_completion=None,
) -> tuple[bytes, str]:
    """High-level entry point used by the web layer.

    Returns ``(xlsx_bytes, download_filename)``.
    """
    project = ProjectInfo(
        name=name,
        orig_substantial_completion=_coerce_date(orig_substantial_completion),
        orig_final_completion=_coerce_date(orig_final_completion),
        current_substantial_completion=_coerce_date(current_substantial_completion),
        current_final_completion=_coerce_date(current_final_completion),
    )
    rows = parse_budget_csv(csv_content)
    wb = build_workbook(rows, project)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), safe_filename(name)
