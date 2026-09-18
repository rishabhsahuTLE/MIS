"""Build the 19-column 'Working' rows from extracted source-sheet rows.

Each row is returned as a WorkingRow: the literal values for columns
A-L/O/P/S, the *computed* python values for the formula columns M/N/Q/R
(used to build Sheet1), and the *formula strings* to actually write into the
Working sheet's cells (so they stay live/editable in Excel, per the user's
choice).
"""

from dataclasses import dataclass

from .config import PeriodConfig, SourceSheet


def _to_text(value) -> str:
    """Reproduce Excel's default number->text conversion used by
    CONCATENATE, so computed strings match what the live formula would show.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip() if isinstance(value, str) else str(value)


@dataclass
class WorkingRow:
    row_number: int  # 1-based row in the Working sheet (data starts at 2)
    invoice_or_credit_note: str          # A
    travel_month: str                    # B
    name_of_passenger: str                # C
    location: str                        # D
    cost_code: str                        # E
    emp_id: str                           # F
    nature_of_transaction: str            # G
    nature_of_ticket: str                 # H
    amounts: float                        # I
    invoice_number: str                   # J
    invoice_date_raw: str                 # K (as found in the source sheet)
    gl: str                               # L
    gl_code: str                          # M (computed)
    gl_code_formula: str                  # M formula string
    narration: str                        # N (computed)
    narration_formula: str                # N formula string
    vendor_id: str                        # O
    submission_date: str                  # P
    gl_code_dup: str                      # Q (computed, == cost_code)
    gl_code_dup_formula: str              # Q formula string
    bu: str                               # R (computed, last 3 chars of Q)
    bu_formula: str                       # R formula string
    service_month: str                    # S


def build_working_rows(
    source_rows_by_sheet: dict[str, list[dict]],
    source_sheets: list[SourceSheet],
    period: PeriodConfig,
) -> list[WorkingRow]:
    rows: list[WorkingRow] = []
    row_number = 2  # Working row 1 is the header

    for sheet_config in source_sheets:
        for record in source_rows_by_sheet.get(sheet_config.name, []):
            cost_code = _to_text(record["cost_code"])
            name = _to_text(record["name_of_passenger"])
            emp_id = _to_text(record["emp_id"])
            location = _to_text(record["location"])
            nature_of_ticket = sheet_config.nature_of_ticket

            gl_code = _to_text(period.gl_prefix) + "-" + cost_code
            narration = (
                nature_of_ticket + " " + name + "-" + emp_id + "-"
                + period.travel_month + "-" + period.submission_date + "-"
                + location
            )
            gl_code_dup = cost_code
            bu = gl_code_dup[-3:]

            r = row_number
            rows.append(
                WorkingRow(
                    row_number=r,
                    invoice_or_credit_note=sheet_config.kind,
                    travel_month=period.travel_month,
                    name_of_passenger=name,
                    location=location,
                    cost_code=cost_code,
                    emp_id=emp_id,
                    nature_of_transaction=sheet_config.nature_of_transaction,
                    nature_of_ticket=nature_of_ticket,
                    amounts=record["amounts"],
                    invoice_number=_to_text(record["invoice_number"]),
                    invoice_date_raw=record["invoice_date"],
                    gl=period.gl_prefix,
                    gl_code=gl_code,
                    gl_code_formula=f'=CONCATENATE(L{r},"-",E{r})',
                    narration=narration,
                    narration_formula=(
                        f'=CONCATENATE(H{r}," ",C{r},"-",F{r},"-",B{r},"-",'
                        f'P{r},"-",D{r})'
                    ),
                    vendor_id=period.vendor_id,
                    submission_date=period.submission_date,
                    gl_code_dup=gl_code_dup,
                    gl_code_dup_formula=f"=E{r}",
                    bu=bu,
                    bu_formula=f"=RIGHT(Q{r},3)",
                    service_month=period.service_month,
                )
            )
            row_number += 1

    return rows
