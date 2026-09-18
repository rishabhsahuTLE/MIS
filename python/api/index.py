"""Flask app: serves the Sheet Injector front-end and wraps mis_builder's
pipeline as an HTTP endpoint. Deployable as-is as a Vercel Python function
(this whole module is auto-detected via its module-level `app`), or run
locally with `python api/index.py` for development.
"""

import io
import json
import sys
from pathlib import Path

# Make sure the repo root (parent of this api/ folder) is importable
# regardless of the working directory the runtime starts this file from.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from flask import Flask, Response, request, send_file
from openpyxl import load_workbook

from mis_builder.config import PeriodConfig
from mis_builder.pipeline import SheetLayoutError, process_workbook

app = Flask(__name__)

_PAGE_PATH = Path(__file__).with_name("index.html")


@app.get("/")
def index():
    return Response(_PAGE_PATH.read_text(encoding="utf-8"), mimetype="text/html")


@app.post("/api/process")
def process():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return {"error": "No file uploaded."}, 400
    if not upload.filename.lower().endswith((".xlsx", ".xlsm")):
        return {"error": "Only .xlsx/.xlsm files are supported."}, 400

    period = PeriodConfig(
        travel_month=request.form.get("travel_month") or PeriodConfig.travel_month,
        service_month=request.form.get("service_month") or PeriodConfig.service_month,
        submission_date=request.form.get("submission_date") or PeriodConfig.submission_date,
        gl_prefix=request.form.get("gl_prefix") or PeriodConfig.gl_prefix,
        vendor_id=request.form.get("vendor_id") or PeriodConfig.vendor_id,
    )

    try:
        wb = load_workbook(io.BytesIO(upload.read()), data_only=False)
    except Exception:
        return {"error": "Could not read this file — is it a valid .xlsx/.xlsm?"}, 400

    try:
        result = process_workbook(wb, period)
    except SheetLayoutError as exc:
        return {"error": str(exc)}, 422
    except Exception as exc:  # pragma: no cover - defensive: surface, don't 500 silently
        return {"error": f"Unexpected error while processing: {exc}"}, 500

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    report = {
        "rows_per_sheet": result.rows_per_sheet,
        "sheets_not_found": result.sheets_not_found,
        "working_row_count": result.working_row_count,
        "sheet1_invoice_count": result.sheet1_invoice_count,
        "sheet1_credit_note_count": result.sheet1_credit_note_count,
    }

    response = send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=upload.filename.rsplit(".", 1)[0] + "-built." + upload.filename.rsplit(".", 1)[1],
    )
    response.headers["X-Report"] = json.dumps(report)
    return response


if __name__ == "__main__":
    app.run(debug=True, port=5000)
