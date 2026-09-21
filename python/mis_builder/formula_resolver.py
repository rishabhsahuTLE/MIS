"""Resolve a source cell to its real value, even when the cell holds a
formula that looks the value up from elsewhere (a plain cross-cell/
cross-sheet reference, or a VLOOKUP/HLOOKUP/INDEX-MATCH against a lookup
table) instead of trusting a stale cached result (which can be a leftover
`#N/A`/`#REF!`/etc. from the last time the source file was saved in Excel).
"""

import re

from openpyxl.utils.cell import range_boundaries

ERROR_VALUES = {"#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}

_CELL_REF_RE = re.compile(
    r"^(?:'(?P<sheet1>[^']+)'|(?P<sheet2>[A-Za-z_][A-Za-z0-9_. ]*))?!?"
    r"\$?(?P<col>[A-Za-z]{1,3})\$?(?P<row>\d+)$"
)
_LOOKUP_RE = re.compile(r"^(?P<fn>V|H)LOOKUP\((?P<args>.*)\)$", re.IGNORECASE | re.DOTALL)
_INDEX_RE = re.compile(r"^INDEX\((?P<args>.*)\)$", re.IGNORECASE | re.DOTALL)
_MATCH_RE = re.compile(r"^MATCH\((?P<args>.*)\)$", re.IGNORECASE | re.DOTALL)
_NUMBER_RE = re.compile(r"^-?\d+(\.\d+)?$")


class UnresolvedLookupError(ValueError):
    """Raised when a formula cell's value can't be resolved to a real answer."""


def resolve_cell(cell_formula, cell_value, *, _seen=None):
    """Return the real value of a cell read from a `data_only=False` and a
    `data_only=True` load of the same workbook.

    `cell_formula`/`cell_value` must be the *same coordinate* on the *same
    sheet name*, one from each workbook.
    """
    if _seen is None:
        _seen = frozenset()

    if cell_formula.data_type != "f":
        return cell_formula.value

    key = (cell_formula.parent.title, cell_formula.coordinate)
    if key in _seen:
        raise UnresolvedLookupError(f"Circular reference at {key[0]}!{key[1]}")

    cached = cell_value.value
    if cached is not None and not (isinstance(cached, str) and cached.strip() in ERROR_VALUES):
        return cached

    wb_formulas = cell_formula.parent.parent
    wb_values = cell_value.parent.parent
    return _resolve_formula(
        cell_formula.value, wb_formulas, wb_values,
        _seen | {key}, cell_formula.parent.title,
    )


def _resolve_formula(formula, wb_formulas, wb_values, seen, origin_sheet):
    text = formula.lstrip("=").strip()

    m = _CELL_REF_RE.match(text)
    if m:
        sheet = m.group("sheet1") or m.group("sheet2") or origin_sheet
        return _resolve_ref(sheet, f"{m.group('col')}{m.group('row')}", wb_formulas, wb_values, seen)

    m = _LOOKUP_RE.match(text)
    if m:
        return _eval_lookup(m.group("fn").upper(), m.group("args"), wb_formulas, wb_values, seen, origin_sheet)

    m = _INDEX_RE.match(text)
    if m:
        return _eval_index_match(m.group("args"), wb_formulas, wb_values, seen, origin_sheet)

    raise UnresolvedLookupError(
        f"Could not resolve formula '{formula}' at {origin_sheet}!"
        f"(unsupported formula shape)"
    )


def _resolve_ref(sheet_name, coord, wb_formulas, wb_values, seen):
    try:
        ws_f = wb_formulas[sheet_name]
        ws_v = wb_values[sheet_name]
    except KeyError as exc:
        raise UnresolvedLookupError(f"Referenced sheet '{sheet_name}' does not exist") from exc
    return resolve_cell(ws_f[coord], ws_v[coord], _seen=seen)


def _split_args(args_str):
    """Split a formula argument list on top-level commas only (ignoring
    commas nested inside parens or quoted strings)."""
    args, depth, current, in_quotes = [], 0, [], False
    for ch in args_str:
        if ch == '"':
            in_quotes = not in_quotes
            current.append(ch)
        elif in_quotes:
            current.append(ch)
        elif ch == "(":
            depth += 1
            current.append(ch)
        elif ch == ")":
            depth -= 1
            current.append(ch)
        elif ch == "," and depth == 0:
            args.append("".join(current).strip())
            current = []
        else:
            current.append(ch)
    if current:
        args.append("".join(current).strip())
    return args


def _eval_value_expr(expr, wb_formulas, wb_values, seen, origin_sheet):
    expr = expr.strip()
    if expr.startswith('"') and expr.endswith('"'):
        return expr[1:-1]
    if _NUMBER_RE.match(expr):
        return float(expr) if "." in expr else int(expr)
    m = _CELL_REF_RE.match(expr)
    if m:
        sheet = m.group("sheet1") or m.group("sheet2") or origin_sheet
        return _resolve_ref(sheet, f"{m.group('col')}{m.group('row')}", wb_formulas, wb_values, seen)
    raise UnresolvedLookupError(f"Could not evaluate lookup argument '{expr}'")


def _parse_range(expr, origin_sheet):
    expr = expr.strip()
    m = re.match(r"^(?:'(?P<sheet1>[^']+)'|(?P<sheet2>[A-Za-z_][A-Za-z0-9_. ]*))?!?(?P<range>.+)$", expr)
    sheet = (m.group("sheet1") or m.group("sheet2") or origin_sheet) if m else origin_sheet
    cell_range = m.group("range") if m else expr
    return sheet, cell_range


def _range_cells(ws, cell_range):
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    if min_row is None:
        min_row = 1
    if max_row is None:
        max_row = ws.max_row
    if min_col is None:
        min_col = 1
    if max_col is None:
        max_col = ws.max_column
    return min_col, min_row, max_col, max_row


def _values_equal(a, b):
    if isinstance(a, str) and isinstance(b, str):
        return a.strip().lower() == b.strip().lower()
    try:
        return float(a) == float(b)
    except (TypeError, ValueError):
        return a == b


def _eval_lookup(fn, args_str, wb_formulas, wb_values, seen, origin_sheet):
    args = _split_args(args_str)
    if len(args) < 3:
        raise UnresolvedLookupError(f"{fn}LOOKUP with too few arguments: {args_str}")
    lookup_value = _eval_value_expr(args[0], wb_formulas, wb_values, seen, origin_sheet)
    table_sheet, table_range = _parse_range(args[1], origin_sheet)
    index = int(_eval_value_expr(args[2], wb_formulas, wb_values, seen, origin_sheet))

    ws_f = wb_formulas[table_sheet]
    ws_v = wb_values[table_sheet]
    min_col, min_row, max_col, max_row = _range_cells(ws_v, table_range)

    if fn == "V":
        for row in range(min_row, max_row + 1):
            if _values_equal(ws_v.cell(row=row, column=min_col).value, lookup_value):
                target_col = min_col + index - 1
                return resolve_cell(
                    ws_f.cell(row=row, column=target_col),
                    ws_v.cell(row=row, column=target_col),
                    _seen=seen,
                )
    else:  # HLOOKUP
        for col in range(min_col, max_col + 1):
            if _values_equal(ws_v.cell(row=min_row, column=col).value, lookup_value):
                target_row = min_row + index - 1
                return resolve_cell(
                    ws_f.cell(row=target_row, column=col),
                    ws_v.cell(row=target_row, column=col),
                    _seen=seen,
                )

    raise UnresolvedLookupError(
        f"{fn}LOOKUP: value {lookup_value!r} not found in "
        f"'{table_sheet}'!{table_range}"
    )


def _eval_index_match(args_str, wb_formulas, wb_values, seen, origin_sheet):
    args = _split_args(args_str)
    if len(args) < 2:
        raise UnresolvedLookupError(f"INDEX with too few arguments: {args_str}")
    range_expr = args[0]
    match_expr = args[1].strip()

    m = _MATCH_RE.match(match_expr)
    if not m:
        raise UnresolvedLookupError(f"Unsupported INDEX(...) shape: INDEX({args_str})")

    match_args = _split_args(m.group("args"))
    lookup_value = _eval_value_expr(match_args[0], wb_formulas, wb_values, seen, origin_sheet)
    match_sheet, match_range = _parse_range(match_args[1], origin_sheet)

    ws_match_v = wb_values[match_sheet]
    min_col, min_row, max_col, max_row = _range_cells(ws_match_v, match_range)
    is_row_vector = min_row == max_row

    position = None
    if is_row_vector:
        for i, col in enumerate(range(min_col, max_col + 1)):
            if _values_equal(ws_match_v.cell(row=min_row, column=col).value, lookup_value):
                position = i
                break
    else:
        for i, row in enumerate(range(min_row, max_row + 1)):
            if _values_equal(ws_match_v.cell(row=row, column=min_col).value, lookup_value):
                position = i
                break

    if position is None:
        raise UnresolvedLookupError(
            f"MATCH: value {lookup_value!r} not found in '{match_sheet}'!{match_range}"
        )

    index_sheet, index_range = _parse_range(range_expr, origin_sheet)
    ws_index_f = wb_formulas[index_sheet]
    ws_index_v = wb_values[index_sheet]
    i_min_col, i_min_row, i_max_col, i_max_row = _range_cells(ws_index_v, index_range)

    if i_min_row == i_max_row:
        target_col, target_row = i_min_col + position, i_min_row
    else:
        target_col, target_row = i_min_col, i_min_row + position

    return resolve_cell(
        ws_index_f.cell(row=target_row, column=target_col),
        ws_index_v.cell(row=target_row, column=target_col),
        _seen=seen,
    )
