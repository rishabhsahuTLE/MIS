"""The reusable end-to-end pipeline: an already-open workbook in, the same
workbook with 'Working'/'Sheet1' (re)built in place, plus a small report.

Both the CLI (mis_builder.cli) and the web API (api/index.py) call this, so
there is exactly one implementation of "read sources -> build Working ->
build Sheet1 -> write them into the workbook".
"""

from dataclasses import dataclass, field

from openpyxl import Workbook

from .config import PeriodConfig, SOURCE_SHEETS
from .extract import SheetLayoutError, read_source_sheet
from .formula_resolver import UnresolvedLookupError
from .sheet1 import build_sheet1_layout
from .working import build_working_rows
from .writer import build_output_workbook


@dataclass
class ProcessResult:
    rows_per_sheet: dict = field(default_factory=dict)   # sheet name -> row count
    sheets_not_found: list = field(default_factory=list)  # sheet names missing from the input
    unparsed_invoice_dates: list = field(default_factory=list)  # "<sheet> row <n>: ..." messages
    working_row_count: int = 0
    sheet1_invoice_count: int = 0
    sheet1_credit_note_count: int = 0


def process_workbook(wb: Workbook, wb_values: Workbook, period: PeriodConfig) -> ProcessResult:
    """Read the known source sheets out of `wb`, build Working + Sheet1 from
    them, and write those two sheets into `wb` (replacing them if they
    already exist). Raises SheetLayoutError if a present source sheet
    doesn't match the expected header layout, or UnresolvedLookupError if a
    mapped cell's formula can't be resolved to a real value.

    `wb_values` must be the same workbook re-opened with `data_only=True`,
    used to resolve formula cells to their real values.
    """
    result = ProcessResult()
    source_rows_by_sheet = {}

    for sheet_config in SOURCE_SHEETS:
        if sheet_config.name not in wb.sheetnames:
            result.sheets_not_found.append(sheet_config.name)
            continue
        rows = read_source_sheet(
            wb[sheet_config.name], wb_values[sheet_config.name], sheet_config
        )
        source_rows_by_sheet[sheet_config.name] = rows
        result.rows_per_sheet[sheet_config.name] = len(rows)

    working_rows, date_warnings = build_working_rows(source_rows_by_sheet, SOURCE_SHEETS, period)
    result.unparsed_invoice_dates = date_warnings
    sheet1_layout = build_sheet1_layout(working_rows)
    build_output_workbook(wb, working_rows, sheet1_layout)

    result.working_row_count = len(working_rows)
    result.sheet1_invoice_count = len(sheet1_layout.table1)
    result.sheet1_credit_note_count = len(sheet1_layout.table2)
    return result


__all__ = ["ProcessResult", "process_workbook", "SheetLayoutError", "UnresolvedLookupError"]
