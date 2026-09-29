"""
convert_timeline.py
-------------------
Reads PhD_Timeline_ANU_3yr.xlsx and writes docs/data.json.

Run this script every time you update the Excel file, then commit
& push to GitHub. GitHub Pages will serve the updated Gantt chart.

Usage:
    python convert_timeline.py

Requirements:
    pip install openpyxl
"""

import json
import os
import sys
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

import openpyxl


def add_months(dt, n):
    """Replicate Excel EDATE: add n calendar months."""
    return dt + relativedelta(months=n)


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(script_dir, "PhD_Timeline_ANU_3yr.xlsx")
    out_path = os.path.join(script_dir, "docs", "data.json")

    if not os.path.exists(xlsx_path):
        print(f"ERROR: {xlsx_path} not found.")
        sys.exit(1)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Timeline"]

    # ---- Enrolment date (D3) ----
    enrol_val = ws["D3"].value
    if isinstance(enrol_val, datetime):
        enrolment_date = enrol_val
    elif isinstance(enrol_val, str):
        enrolment_date = datetime.fromisoformat(enrol_val)
    else:
        enrolment_date = datetime(2026, 9, 29)

    # ---- Key dates ----
    key_dates = [
        {"label": "Enrolment Start", "date": enrolment_date.isoformat()},
        {"label": "CoC Target (M9)", "date": add_months(enrolment_date, 9).isoformat()},
        {"label": "Final Seminar (M30)", "date": add_months(enrolment_date, 30).isoformat()},
        {"label": "Thesis Submission (M36)", "date": add_months(enrolment_date, 36).isoformat()},
        {"label": "Scholarship End (M42)", "date": (add_months(enrolment_date, 42) - timedelta(days=1)).isoformat()},
        {"label": "MEP (M48)", "date": (add_months(enrolment_date, 48) - timedelta(days=1)).isoformat()},
    ]

    # ---- Tasks (rows 14–57) ----
    tasks = []
    for r in range(14, 58):
        row_id = ws.cell(r, 1).value        # A: ID
        if not row_id:
            continue
        phase = ws.cell(r, 2).value or ""    # B: Phase
        task_name = ws.cell(r, 3).value or ""  # C: Task
        task_type = ws.cell(r, 4).value or "Task"  # D: Type
        start_m = ws.cell(r, 5).value        # E: Start M
        end_m = ws.cell(r, 6).value          # F: End M
        owner = ws.cell(r, 9).value or ""    # I: Owner
        status = ws.cell(r, 10).value or "Not started"  # J: Status
        notes = ws.cell(r, 11).value or ""   # K: Notes

        # Compute dates from month offsets (replicating the Excel formulas)
        start_date = add_months(enrolment_date, (start_m or 1) - 1)
        end_date = add_months(enrolment_date, (end_m or 1)) - timedelta(days=1)

        tasks.append({
            "id": row_id,
            "phase": phase,
            "task": task_name,
            "type": task_type,
            "startM": start_m or 1,
            "endM": end_m or 1,
            "startDate": start_date.isoformat(),
            "endDate": end_date.isoformat(),
            "owner": owner,
            "status": status,
            "notes": notes,
        })

    data = {
        "enrolmentDate": enrolment_date.isoformat(),
        "generatedAt": datetime.now().isoformat(),
        "keyDates": key_dates,
        "tasks": tasks,
    }

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=str)

    print(f"[OK] Wrote {len(tasks)} tasks to {out_path}")
    print(f"     Enrolment date: {enrolment_date.strftime('%d %b %Y')}")
    print(f"     Generated at:   {datetime.now().strftime('%d %b %Y %H:%M')}")


if __name__ == "__main__":
    main()
