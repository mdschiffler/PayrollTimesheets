# Project Snapshot

_Last updated: 2026-05-21_

## Mission

Turn payroll CSV reports (Notion contractor timesheets, Turno cleaning jobs, NGTecoTime punches) into a reviewable per-worker Excel workbook so the operator can approve payroll in one pass.

## Primary user

One non-technical end user (the project owner's wife) running the macOS app by double-clicking it from iCloud Drive. She does not run scripts, edit Python, or use the terminal. Everything she needs must work from the GUI, the rates CSV, and the generated workbook.

## Stack

- Language: Python 3 (CPython, system or bundled).
- Processing: `pandas` + `xlsxwriter`.
- Time zones: `zoneinfo` (`America/Puerto_Rico`).
- GUI: Tkinter (`ttk`).
- Packaging: PyInstaller producing a macOS `.app`.
- Persistence: a CSV rate table and generated `.xlsx` workbooks. No database.

## Source of truth

When repo and brief conflict, the repo wins. Within the repo, the canonical artifacts are:

| Concern | Source of truth |
|---|---|
| Processing logic | [_dev/export-timesheet.py](../_dev/export-timesheet.py) |
| GUI behavior | [_dev/payroll_app.py](../_dev/payroll_app.py) |
| Employee rates, start dates, extras, withholding details | [timesheet-rates.csv](../timesheet-rates.csv) |
| Build steps | [_dev/build_app.sh](../_dev/build_app.sh) |
| Wife-facing usage | [README.md](../README.md) |
| Agent rules and index | [AGENTS.md](../AGENTS.md) |

## Key entry points

- **End user:** `Optihome Payroll Processing.app` at repo root (alias to `dist/Optihome Payroll Processing.app`).
- **GUI dev:** `python3 _dev/payroll_app.py`.
- **CLI:** `python3 _dev/export-timesheet.py --output <path>.xlsx [--notion …] [--turno …] [--time …] [--rates …]`.
- **Library API:** `process_timesheet(csv_file, output_excel, turno_csv=None, rates_csv=None, notion_csv=None)` in `_dev/export-timesheet.py:556`.
- **Rebuild app:** `bash _dev/build_app.sh`.

## What this project is NOT

- Not a payroll _calculator of taxes_ — it surfaces a 10% withholding line for manual review and a no-withholding allowance field; it does not file or remit anything.
- Not a time-tracking tool. It only consumes exports from Notion, Turno, and NGTecoTime.
- Not multi-tenant, not cloud-hosted, not networked. Files live in iCloud Drive on the user's Mac.
