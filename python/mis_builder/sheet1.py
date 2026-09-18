"""Derive the 'Sheet1' reporting view from a built list of Working rows.

Sheet1 is two tables stacked with a gap between them:
  - Table 1: every 'Invoice' row, sorted by passenger name then invoice
    number, ending in a SUM(...) total row.
  - Table 2: every 'Credit Note' row, same sort, placed SHEET1_TABLE_GAP
    blank rows below Table 1's total, with its own SUM(...) total row and an
    extra helper column that turns the amount negative.
"""

from dataclasses import dataclass
from datetime import datetime

from .config import SHEET1_TABLE_GAP


_DATE_FORMATS = (
    "%d-%b-%Y",              # 16-AUG-2026
    "%d-%m-%Y %I:%M:%S %p",  # 17-08-2026 12:00:00 AM
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%Y-%m-%d",
)


def parse_invoice_date(value):
    """Return a python date/datetime for a source 'Reference Date' cell.

    Source cells are inconsistently either already-parsed datetimes (if
    Excel stored them as a real date) or text in one of a few formats.
    Returns None (and leaves the cell blank) if nothing matches, rather than
    guessing wrong.
    """
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


@dataclass
class Sheet1Row:
    name_of_passenger: str
    emp_id: str
    invoice_number: str
    invoice_date: object  # datetime or None
    gl_code: str
    narration: str
    vendor_id: str
    invoice_or_credit_note: str
    amounts: float
    cost_code: str  # for the K/L "GL code"-suffix helper columns


def _to_sheet1_row(w) -> Sheet1Row:
    return Sheet1Row(
        # Sheet1's name column is always upper-cased, even when the source
        # sheet stored it in title case (e.g. hotel sheets' "Employee Name").
        name_of_passenger=w.name_of_passenger.upper(),
        cost_code=w.cost_code,
        emp_id=w.emp_id,
        invoice_number=w.invoice_number,
        invoice_date=parse_invoice_date(w.invoice_date_raw),
        gl_code=w.gl_code,
        narration=w.narration,
        vendor_id=w.vendor_id,
        invoice_or_credit_note=w.invoice_or_credit_note,
        amounts=w.amounts,
    )


def _sort_key(row: Sheet1Row):
    return (row.name_of_passenger.upper(), row.invoice_number.upper())


@dataclass
class Sheet1Layout:
    table1: list  # list[Sheet1Row], the 'Invoice' block
    table2: list  # list[Sheet1Row], the 'Credit Note' block
    table1_start_row: int
    table1_total_row: int
    table2_start_row: int
    table2_total_row: int


def build_sheet1_layout(working_rows, table_gap: int = SHEET1_TABLE_GAP) -> Sheet1Layout:
    table1 = sorted(
        (_to_sheet1_row(w) for w in working_rows if w.invoice_or_credit_note == "Invoice"),
        key=_sort_key,
    )
    table2 = sorted(
        (_to_sheet1_row(w) for w in working_rows if w.invoice_or_credit_note == "Credit Note"),
        key=_sort_key,
    )

    table1_start_row = 2  # row 1 is the header
    table1_total_row = table1_start_row + len(table1)
    table2_start_row = table1_total_row + 1 + table_gap
    table2_total_row = table2_start_row + len(table2)

    return Sheet1Layout(
        table1=table1,
        table2=table2,
        table1_start_row=table1_start_row,
        table1_total_row=table1_total_row,
        table2_start_row=table2_start_row,
        table2_total_row=table2_total_row,
    )
