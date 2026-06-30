"""
Denovo Attendance Report — Web UI
"""

import os
import uuid
import tempfile
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file

from denovo_attendance import (
    load_clockings, load_employees, generate_report,
    get_weeks, week_start,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB

UPLOAD_DIR = Path(tempfile.gettempdir()) / "denovo_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _cleanup_old_files():
    """Remove upload files older than 1 hour."""
    import time
    cutoff = time.time() - 3600
    for f in UPLOAD_DIR.iterdir():
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
        except OSError:
            pass


def _weeks_from_clockings(clockings: dict) -> list:
    """
    Return list of week dicts sorted chronologically.
    Each dict: { week_id, monday (ISO), sunday (ISO), label, month, month_label, dates }
    """
    all_dates = {
        d
        for emp_data in clockings.values()
        for d in emp_data.keys()
    }
    if not all_dates:
        return []

    week_dates: dict = defaultdict(list)
    for d in all_dates:
        mon = week_start(d)
        week_dates[mon].append(d)

    weeks = []
    for mon in sorted(week_dates):
        sun = mon + timedelta(days=6)
        weeks.append({
            "week_id": mon.isoformat(),
            "monday": mon.isoformat(),
            "sunday": sun.isoformat(),
            "label": f"{mon.strftime('%d %b')} – {sun.strftime('%d %b %Y')}",
            "month": mon.strftime("%Y-%m"),
            "month_label": mon.strftime("%B %Y"),
        })
    return weeks


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/parse-dat", methods=["POST"])
def parse_dat():
    """Receive .dat file, store it, return available weeks."""
    _cleanup_old_files()

    f = request.files.get("dat_file")
    if not f:
        return jsonify(error="No file received"), 400

    file_id = str(uuid.uuid4())
    dat_path = UPLOAD_DIR / f"{file_id}.dat"
    f.save(str(dat_path))

    try:
        clockings = load_clockings(str(dat_path))
    except Exception as e:
        dat_path.unlink(missing_ok=True)
        return jsonify(error=f"Could not parse file: {e}"), 400

    if not clockings:
        dat_path.unlink(missing_ok=True)
        return jsonify(error="No clocking data found in file."), 400

    weeks = _weeks_from_clockings(clockings)
    emp_count = len(clockings)

    return jsonify(
        file_id=file_id,
        employee_count=emp_count,
        weeks=weeks,
    )


@app.route("/generate", methods=["POST"])
def generate():
    """Generate report and return .docx."""
    file_id = request.form.get("file_id", "").strip()
    selected_weeks = request.form.getlist("weeks")
    company = request.form.get("company", "Denovo Apparel Ltd").strip()

    if not file_id or not selected_weeks:
        return jsonify(error="Missing file ID or week selection."), 400

    dat_path = UPLOAD_DIR / f"{file_id}.dat"
    if not dat_path.exists():
        return jsonify(error="Upload session expired. Please re-upload the .dat file."), 400

    csv_file = request.files.get("csv_file")
    if not csv_file:
        return jsonify(error="No employee CSV file received."), 400

    csv_path = UPLOAD_DIR / f"{file_id}.csv"
    csv_file.save(str(csv_path))

    # Compute period from selected weeks
    mondays = sorted(date.fromisoformat(w) for w in selected_weeks)
    period_start = mondays[0]
    period_end = mondays[-1] + timedelta(days=6)

    output_path = UPLOAD_DIR / f"{file_id}_report.docx"

    try:
        generate_report(
            dat_path=str(dat_path),
            csv_path=str(csv_path),
            output_path=str(output_path),
            company=company or "Denovo Apparel Ltd",
            period_start=period_start,
            period_end=period_end,
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify(error=f"Report generation failed: {e}"), 500

    period_label = f"{period_start.strftime('%d%b')}_{period_end.strftime('%d%b%Y')}".replace(" ", "")
    download_name = f"Attendance_{period_label}.docx"

    return send_file(
        str(output_path),
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
