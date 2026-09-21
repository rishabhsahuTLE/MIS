"""Read one source sheet (DOMESTIC DEBIT, HOTEL CREDIT, ...) into a list of
normalized row dicts, per the mapping in config.SourceSheet.
"""

from openpyxl.worksheet.worksheet import Worksheet

from .config import HEADER_ROW, REQUIRED_FIELDS, TOTAL_MARKER, SourceSheet
from .formula_resolver import resolve_cell


class SheetLayoutError(ValueError):
    """Raised when a source sheet doesn't match the expected header layout."""


def _normalize(text):
    if text is None:
        return ""
    return str(text).strip()


def _build_header_index(ws: Worksheet, header_row: int) -> dict:
    """Map normalized header text -> leftmost 1-based column index.

    Some sheets repeat a header name (e.g. HOTEL DEBIT has "Reference No."
    twice) -- the first occurrence (lowest column index) wins, and later
    duplicates are ignored, since we only ever look up a header by its exact
    expected text.
    """
    header_index = {}
    for cell in ws[header_row]:
        text = _normalize(cell.value)
        if not text:
            continue
        key = text.lower()
        if key not in header_index:
            header_index[key] = cell.column
    return header_index


def _row_is_stop_marker(ws: Worksheet, row: int, first_col: int, last_col: int) -> bool:
    values = [ws.cell(row=row, column=c).value for c in range(first_col, last_col + 1)]
    if all(v is None or _normalize(v) == "" for v in values):
        return True
    first_val = _normalize(values[0]).lower()
    if first_val == TOTAL_MARKER:
        return True
    return False


def read_source_sheet(
    ws: Worksheet, ws_values: Worksheet, sheet_config: SourceSheet
) -> list[dict]:
    """Read all data rows of a source sheet into normalized dicts.

    Stops at the first fully-blank row or a row whose first cell reads
    "Total" -- this also correctly excludes trailing junk (summary blocks,
    orphan rows) that sits below a sheet's real data table.

    `ws_values` must be the same sheet re-opened with `data_only=True`, used
    to resolve any mapped cell that holds a formula (a lookup from another
    cell/sheet) to its real value instead of a stale cached result.
    """
    header_index = _build_header_index(ws, HEADER_ROW)

    missing = []
    field_to_col = {}
    for field_name, header_text in sheet_config.columns.items():
        col = header_index.get(header_text.lower())
        if col is None:
            missing.append(header_text)
        else:
            field_to_col[field_name] = col

    if missing:
        raise SheetLayoutError(
            f"Sheet '{sheet_config.name}': expected header(s) {missing!r} not "
            f"found in row {HEADER_ROW}. Found headers: "
            f"{sorted(header_index)!r}"
        )
    for field_name in REQUIRED_FIELDS:
        if field_name not in field_to_col:
            raise SheetLayoutError(
                f"Sheet '{sheet_config.name}': no mapping configured for "
                f"required field '{field_name}'."
            )

    first_col = min(field_to_col.values())
    last_col = max(ws.max_column, max(field_to_col.values()))

    rows = []
    row_num = HEADER_ROW + 1
    while row_num <= ws.max_row:
        if _row_is_stop_marker(ws, row_num, first_col, last_col):
            break
        record = {
            field_name: resolve_cell(
                ws.cell(row=row_num, column=col),
                ws_values.cell(row=row_num, column=col),
            )
            for field_name, col in field_to_col.items()
        }
        rows.append(record)
        row_num += 1

    return rows
