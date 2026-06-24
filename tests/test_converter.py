"""Tests for the CSV -> Job Cost Projection workbook conversion."""

import io
import re
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from backend import converter

SAMPLE_CSV = Path(__file__).parent / "sample" / "budget_details_3.csv"


@pytest.fixture(scope="module")
def csv_text():
    return SAMPLE_CSV.read_text()


@pytest.fixture(scope="module")
def parsed(csv_text):
    return converter.parse_budget_csv(csv_text)


# --------------------------------------------------------------------------- #
# Parsing (steps 6-7)
# --------------------------------------------------------------------------- #
def test_parse_drops_placeholder_rows(parsed):
    # The sample has 64 data lines; one is the Procore "None" placeholder.
    assert len(parsed) == 63
    assert all(str(r[0]).strip().lower() not in ("", "none") for r in parsed)


def test_parse_keeps_nine_columns(parsed):
    assert all(len(r) == 9 for r in parsed)


def test_parse_column_mapping(parsed):
    first = parsed[0]
    # A=cost code, B=CAT, C=original budget ... I=job to date
    assert first[0] == "1-020 - Superintendent"
    assert first[1] == "L - Labor"
    assert first[2] == 134000.0          # Original Budget Amount
    assert first[3] == -30500.0          # Budget Modifications
    assert first[8] == 39358.6           # Job to date Costs


def test_numeric_columns_are_numbers(parsed):
    for row in parsed:
        for value in row[2:]:
            assert value is None or isinstance(value, (int, float))


def test_bad_csv_raises():
    with pytest.raises(converter.ConversionError):
        converter.parse_budget_csv("not,a,budget,export\n1,2,3,4")


def test_empty_csv_raises():
    with pytest.raises(converter.ConversionError):
        converter.parse_budget_csv("")


def test_spoofed_csv_with_wrong_columns_is_rejected():
    # 12 columns, column 1 is "Cost Code Tier 2", but the other key columns are
    # wrong -> must be rejected instead of silently mapping the wrong fields.
    header = (
        "Tier1,Cost Code Tier 2,WRONG Type,Budget Code,Desc,WRONG Budget,"
        "Mods,COs,Revised,Committed,Direct,WRONG JTD\n"
    )
    row = "1,1-020 - X,L,code,d,100,0,0,100,0,50,50\n"
    with pytest.raises(converter.ConversionError) as exc:
        converter.parse_budget_csv(header + row)
    assert "column 2" in str(exc.value) or "column 5" in str(exc.value)


# --------------------------------------------------------------------------- #
# Workbook building (steps 8-9 + milestones)
# --------------------------------------------------------------------------- #
@pytest.fixture(scope="module")
def workbook_bytes(csv_text):
    data, fname = converter.convert_csv_to_workbook_bytes(
        csv_text,
        name="Maple Street",
        orig_substantial_completion="2025-09-15",
        orig_final_completion="2025-11-30",
        current_substantial_completion="2025-10-15",
        current_final_completion="2025-12-20",
    )
    return data, fname


def _ws(workbook_bytes):
    data, _ = workbook_bytes
    return load_workbook(io.BytesIO(data))["Job Cost"]


def test_filename_is_project_name(workbook_bytes):
    _, fname = workbook_bytes
    assert fname == "Maple Street.xlsx"


def test_data_pasted_at_row_8(workbook_bytes):
    ws = _ws(workbook_bytes)
    assert ws["A8"].value == "1-020 - Superintendent"
    assert ws["B8"].value == "L - Labor"
    assert ws["C8"].value == 134000.0


def test_surplus_rows_deleted(workbook_bytes, parsed):
    ws = _ws(workbook_bytes)
    totals_row = 8 + len(parsed)  # 71
    # The row right after the data is the totals row, not an empty data row.
    assert ws[f"C{totals_row}"].value == f"=SUM(C8:C{totals_row - 1})"
    # No data rows survive beyond the totals row except the summary block.
    assert ws[f"A{totals_row - 1}"].value == parsed[-1][0]


def test_totals_and_header_formulas_repointed(workbook_bytes, parsed):
    ws = _ws(workbook_bytes)
    totals_row = 8 + len(parsed)
    assert ws["C3"].value == f"=+C{totals_row}"
    assert ws["C4"].value == f"=+E{totals_row}"
    assert ws["C5"].value == f"=+F{totals_row}"


def test_per_row_formulas_present(workbook_bytes, parsed):
    ws = _ws(workbook_bytes)
    last = 7 + len(parsed)
    assert ws[f"K{last}"].value == f"=+I{last}+J{last}"
    assert ws[f"L{last}"].value == f"=+F{last}-K{last}"


def test_milestone_dates_written(workbook_bytes):
    ws = _ws(workbook_bytes)
    assert ws["I3"].value.date() == date(2025, 9, 15)
    assert ws["I4"].value.date() == date(2025, 11, 30)
    assert ws["K3"].value.date() == date(2025, 10, 15)
    assert ws["K4"].value.date() == date(2025, 12, 20)
    assert ws["I3"].number_format == converter.DATE_FORMAT
    assert ws["K3"].number_format == converter.DATE_FORMAT


def test_no_dangling_formula_references(workbook_bytes):
    ws = _ws(workbook_bytes)
    max_row = ws.max_row
    ref = re.compile(r"\$?[A-Z]{1,2}\$?(\d+)")
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                for m in ref.finditer(cell.value):
                    assert int(m.group(1)) <= max_row, (cell.coordinate, cell.value)


def test_no_milestones_leaves_cells_blank(csv_text):
    data, _ = converter.convert_csv_to_workbook_bytes(csv_text, name="No Dates")
    ws = load_workbook(io.BytesIO(data))["Job Cost"]
    assert ws["I3"].value is None
    assert ws["K4"].value is None
    assert ws["F3"].value is None          # contract amount on last pay app
    assert ws["G3"].value == "(month)"     # template placeholder left intact


def test_last_pay_app_fields_written(csv_text):
    data, _ = converter.convert_csv_to_workbook_bytes(
        csv_text,
        name="Pay App",
        contract_amount_last_pay_app="1,234,567.89",
        month_last_pay_app="2026-05-31",
    )
    ws = load_workbook(io.BytesIO(data))["Job Cost"]
    assert ws["F3"].value == 1234567.89
    assert ws["G3"].value.date() == date(2026, 5, 31)
    assert ws["G3"].number_format == converter.DATE_FORMAT


def test_safe_filename():
    assert converter.safe_filename("A/B:C*?") == "A_B_C__.xlsx"
    assert converter.safe_filename("   ") == "Job Cost Projection.xlsx"


def test_safe_filename_truncates_long_names():
    out = converter.safe_filename("X" * 500)
    assert len(out) <= 120
    assert out.endswith(".xlsx")
