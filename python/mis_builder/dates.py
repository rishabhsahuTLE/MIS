"""Shared invoice-date parsing, used by both 'Working' and 'Sheet1'.

Source 'Reference Date' cells are inconsistently either already-parsed
datetimes (if Excel stored them as a real date) or text in one of a few
formats -- some with a time-of-day component. This always returns a
date-only value (time-of-day dropped) so every downstream sheet shows a
plain DD-MM-YYYY date, or None (leave the cell blank) if nothing matches,
rather than guessing wrong.
"""

from datetime import date, datetime

_DATE_FORMATS = (
    "%d-%b-%Y",              # 16-AUG-2026
    "%d-%m-%Y %I:%M:%S %p",  # 17-08-2026 12:00:00 AM
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%Y-%m-%d",
)


def parse_invoice_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None
