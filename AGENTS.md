# Optihome Payroll Processing - Agent Guide

## Purpose

This project converts payroll report CSVs into a reviewable Excel workbook for manual payroll approval. The user-facing entry point is the macOS app alias at the repo root; the source lives in `_dev/`.

## Architecture

- `_dev/export-timesheet.py` contains the processing logic and CLI.
- `_dev/payroll_app.py` is the Tkinter desktop wrapper.
- `_dev/build_app.sh` rebuilds the standalone macOS app with PyInstaller.
- `timesheet-rates.csv` is the editable employee lookup table.

The exporter accepts any combination of:

- Notion contractor timesheet CSV (`*_notion.csv`)
- Turno cleaning job CSV (`*_turno.csv`)
- NGTecoTime timeclock CSV (`*_time.csv`)

At least one input report is required.

## Processing Rules

- Notion rows are hourly work. `Person` identifies the paid worker; if blank, `Team Member` is used. Timestamps ending in `Z` are converted from UTC to `America/Puerto_Rico`. All rows with valid positive hours are included regardless of `Status`.
- Timeclock rows are hourly work. The script groups by person and punch date, then uses the earliest and latest punch as start/end.
- Turno rows are cleaning jobs. Cleaning prices come from Turno and are split when multiple teammates are assigned to the same property on the same date.
- Hourly pay is calculated from `timesheet-rates.csv` `RATE`.
- Recurring extras and withholding use `START`, `EXTRA`, and `DETAILS` from `timesheet-rates.csv`.
- If a Notion file is present, the period is 14 days ending on the `MM-DD-YYYY` filename date. Otherwise the period is 7 days.
- Expense reimbursement import is intentionally out of scope until a real Notion expense export exists.

## CLI

Preferred flag form:

```bash
python3 _dev/export-timesheet.py \
  --output /tmp/04-22-2026.xlsx \
  --notion Raw/2026/04-22-2026_notion.csv \
  --turno Raw/2026/04-22-2026_turno.csv \
  --rates timesheet-rates.csv
```

Legacy positional form is still supported:

```bash
python3 _dev/export-timesheet.py Timesheets/04-22-2026.xlsx Raw/2026/04-22-2026_time.csv Raw/2026/04-22-2026_turno.csv
```

## Development Checks

Run syntax checks after edits:

```bash
python3 -m py_compile _dev/export-timesheet.py _dev/payroll_app.py
```

Useful smoke tests:

```bash
python3 _dev/export-timesheet.py --output /tmp/04-22-2026_notion.xlsx --notion Raw/2026/04-22-2026_notion.csv --rates timesheet-rates.csv
python3 _dev/export-timesheet.py --output /tmp/04-22-2026_turno.xlsx --turno Raw/2026/04-22-2026_turno.csv --rates timesheet-rates.csv
python3 _dev/export-timesheet.py --output /tmp/01-21-2026_time.xlsx --time Raw/2026/01-21-2026_time.csv --rates timesheet-rates.csv
```

Inspect generated workbooks in Excel or by unzipping the `.xlsx` and checking `xl/sharedStrings.xml` / worksheet XML.

## GUI Notes

The GUI persists preferences in `~/.optihome_payroll_config.json`, including the period end day, last-used folders, and visible report selectors. Default visible reports are Notion and Turno; Timeclock is hidden until enabled in Advanced Settings.

When source changes need to be reflected in the double-clickable app, rebuild:

```bash
bash _dev/build_app.sh
```

The build script creates `dist/Optihome Payroll Processing.app` and a Finder alias named `Optihome Payroll Processing.app` at the project root.

## Maintenance Guidance

- Keep `README.md` focused on the human workflow.
- Keep this file as the canonical technical guide for future agents.
- Avoid duplicating long technical details in `_dev/CLAUDE.md`; point it here instead.
- Preserve local raw CSVs, generated workbooks, and existing unrelated worktree changes.
