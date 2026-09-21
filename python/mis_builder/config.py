"""Static mapping config: which source-sheet column feeds which normalized
field, per source sheet, plus the per-period constants used to fill
'Working'. Mirrors MIS_Working_Column_Mapping.xlsx.

If a future month adds/renames a source sheet or column, update SOURCE_SHEETS
here -- nothing else in the codebase should need to change.
"""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class SourceSheet:
    name: str                     # exact tab name in the input workbook
    kind: str                     # 'Invoice' or 'Credit Note' -> Working col A
    nature_of_transaction: str    # Working col G (constant for this sheet)
    nature_of_ticket: str         # Working col H (constant for this sheet)
    # normalized field -> exact header text (row 3) to look up in this sheet
    columns: dict = field(default_factory=dict)


REQUIRED_FIELDS = (
    "name_of_passenger",
    "location",
    "cost_code",
    "emp_id",
    "amounts",
    "invoice_number",
    "invoice_date",
)

# Order here IS the row order used when building 'Working': every row from
# the first sheet, then every row from the second sheet, and so on.
SOURCE_SHEETS = [
    SourceSheet(
        name="DOMESTIC DEBIT",
        kind="Invoice",
        nature_of_transaction="Airways Ticket",
        nature_of_ticket="Airticket",
        columns={
            "name_of_passenger": "Passenger Name",
            "location": "Sector",
            "cost_code": "Cost Code",
            "emp_id": "Employee Code",
            "amounts": "Total Fare",
            "invoice_number": "Reference No.",
            "invoice_date": "Reference Date",
        },
    ),
    SourceSheet(
        name="MISCELLANEOUS DEBIT",
        kind="Invoice",
        nature_of_transaction="Airways Ticket",
        nature_of_ticket="Misc",
        columns={
            "name_of_passenger": "Passenger Name",
            "location": "Sector",
            "cost_code": "Cost Code",
            "emp_id": "Employee Code",
            "amounts": "Total Fare",
            "invoice_number": "Reference No.",
            "invoice_date": "Reference Date",
        },
    ),
    SourceSheet(
        name="RAIL DEBIT",
        kind="Invoice",
        nature_of_transaction="Railways Ticket",
        nature_of_ticket="Railticket",
        columns={
            "name_of_passenger": "Passenger Name",
            "location": "Sector",
            "cost_code": "Cost Code",
            "emp_id": "Employee Code",
            "amounts": "Total Fare",
            "invoice_number": "Reference No.",
            "invoice_date": "Reference Date",
        },
    ),
    SourceSheet(
        name="BUS DEBIT",
        kind="Invoice",
        nature_of_transaction="Bus Ticket",
        nature_of_ticket="Bus",
        columns={
            "name_of_passenger": "Passenger Name",
            "location": "Sector",
            "cost_code": "Cost Code",
            "emp_id": "Employee Code",
            "amounts": "Total Fare",
            "invoice_number": "Reference No.",
            "invoice_date": "Reference Date",
        },
    ),
    SourceSheet(
        name="HOTEL DEBIT",
        kind="Invoice",
        nature_of_transaction="Hotel ",
        nature_of_ticket="Debit",
        columns={
            "name_of_passenger": "Employee Name",
            "location": "Hotel State",
            "cost_code": "Cost Code",
            "emp_id": "Employee Id",
            "amounts": "Q2T Invoice Amount",
            "invoice_number": "Reference No.",
            "invoice_date": "Reference Date",
        },
    ),
    SourceSheet(
        name="DOMESTIC CREDIT",
        kind="Credit Note",
        nature_of_transaction="Airways Ticket",
        nature_of_ticket="Credit",
        columns={
            "name_of_passenger": "Passenger Name",
            "location": "Sector",
            "cost_code": "Cost Code",
            "emp_id": "Employee Code",
            "amounts": "Refund Amount",
            "invoice_number": "Refund Reference No.",
            "invoice_date": "Refund Reference Date",
        },
    ),
    SourceSheet(
        name="RAIL CREDIT",
        kind="Credit Note",
        nature_of_transaction="Railways Ticket",
        nature_of_ticket="Credit",
        columns={
            "name_of_passenger": "Passenger Name",
            "location": "Sector",
            "cost_code": "Cost Code",
            "emp_id": "Employee Code",
            "amounts": "Refund Amount",
            "invoice_number": "Refund Reference No.",
            "invoice_date": "Refund Reference Date",
        },
    ),
    SourceSheet(
        name="BUS CREDIT",
        kind="Credit Note",
        nature_of_transaction="Bus Ticket",
        nature_of_ticket="Credit",
        columns={
            "name_of_passenger": "Passenger Name",
            "location": "Sector",
            "cost_code": "Cost Code",
            "emp_id": "Employee Code",
            "amounts": "Refund Amount",
            "invoice_number": "Refund Reference No.",
            "invoice_date": "Refund Reference Date",
        },
    ),
    SourceSheet(
        name="HOTEL CREDIT",
        kind="Credit Note",
        nature_of_transaction="Hotel ",
        nature_of_ticket="Credit",
        columns={
            "name_of_passenger": "Employee Name",
            "location": "Hotel State",
            "cost_code": "Cost Code",
            "emp_id": "Employee Id",
            "amounts": "Refund Amount",
            "invoice_number": "Refund Reference No.",
            "invoice_date": "Refund Reference Date",
        },
    ),
]


@dataclass(frozen=True)
class PeriodConfig:
    """Values that change every MIS cycle -- override via CLI flags."""

    travel_month: str = "AUG-26"
    service_month: str = "AUG-26"
    submission_date: str = "06-05-2026"
    gl_prefix: str = "80003"
    vendor_id: str = "V03136"


# Header row is always row 3 (row 1 = title, row 2 = blank) across every
# source sheet observed so far.
HEADER_ROW = 3
TOTAL_MARKER = "total"

WORKING_SHEET_NAME = "Working"
SHEET1_NAME = "Sheet1"

WORKING_HEADERS = [
    "Invoice/ Credit Note", "Travel Month", "Name of Passenger", "Location",
    "Cost Code", "Emp ID", "Nature of Transaction", "Nature of Ticket",
    "Amounts", "Invoice Number", "Invoice Date", "gl", "GL Code",
    "Narration", "Vendor ID", "Submission Date", "GL Code", "BU",
    "Service Month",
]

SHEET1_HEADERS = [
    "Name of Passenger", "Emp ID", "Invoice Number", "Invoice Date",
    "GL Code", "Narration", "Vendor ID", "Invoice/ Credit Note",
    "Sum of Amounts",
]

# Blank rows left between Sheet1's Invoice table and its Credit Note table,
# matching the layout of the original workbook.
SHEET1_TABLE_GAP = 6
