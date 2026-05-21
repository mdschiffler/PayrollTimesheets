# Testing & Verification

_Last updated: 2026-05-21_

There is no automated test suite. The project ships with a CLI you can point at real or sample CSVs and inspect the output workbook by hand. Treat verification as scenario-driven, not coverage-driven.

## Required checks after editing source files

These are cheap and should run after every edit to `_dev/export-timesheet.py` or `_dev/payroll_app.py`:

```bash
python3 -m py_compile _dev/export-timesheet.py _dev/payroll_app.py
```

(Also enshrined in user-level memory: build after edits.)

## Smoke scenarios

Run at least one of these against the most recent `Raw/<year>/` inputs after a non-trivial change. Each writes a throwaway workbook to `/tmp/`.

```bash
# Notion only (14-day period)
python3 _dev/export-timesheet.py \
  --output /tmp/notion_only.xlsx \
  --notion Raw/2026/04-22-2026_notion.csv \
  --rates timesheet-rates.csv

# Turno only (7-day period)
python3 _dev/export-timesheet.py \
  --output /tmp/turno_only.xlsx \
  --turno Raw/2026/04-22-2026_turno.csv \
  --rates timesheet-rates.csv

# Timeclock only
python3 _dev/export-timesheet.py \
  --output /tmp/time_only.xlsx \
  --time Raw/2026/01-21-2026_time.csv \
  --rates timesheet-rates.csv

# Notion + Turno together (the typical fortnightly run)
python3 _dev/export-timesheet.py \
  --output /tmp/notion_turno.xlsx \
  --notion Raw/2026/04-22-2026_notion.csv \
  --turno Raw/2026/04-22-2026_turno.csv \
  --rates timesheet-rates.csv
```

## What to look at in the workbook

For each smoke scenario, open the `.xlsx` and confirm:

- **Summary sheet:** column order is `Person | Role | Period | Total Days | Total Hours | Total Cleans | Total $ | Withheld $ | Pay/Hour | Pay/Job | Reviewed`, with an "All sheets total" row.
- **Per-worker sheet:** sections appear in this order: Hourly Work → Mango Villas → Casa Damisela → MARU → Other → Summary block.
- **Hourly Work:** rates match `timesheet-rates.csv`. Notion rows show start/end in Puerto Rico time.
- **Turno sections:** when two teammates appear on the same `Property Alias` + date, the rate is split evenly between them.
- **Allowance row:** shaded red and pre-filled with $500 only for workers whose `START` is within the last 28 days or when the month is January.
- **Withheld:** equals `ROUNDDOWN(MAX(Subtotal − Allowance, 0) * 0.10, 2)`.
- **Total $:** equals `Subtotal − Withheld`.
- **Period:** reflects 14 days for workers with Notion rows, 7 days otherwise.

You can inspect a workbook without Excel by unzipping it and reading `xl/sharedStrings.xml` and `xl/worksheets/sheetN.xml`.

## Match verification depth to risk

| Change | Required verification |
|---|---|
| Comment / docstring / doc-only | `py_compile` and skip the rest. |
| GUI label / copy | `py_compile` + open the GUI and click through. |
| Parser logic (Notion/Turno/Timeclock) | Run all four smoke scenarios above and open each workbook. |
| Workbook structure / section order / formulas | Run all four scenarios + diff a previous workbook to confirm only intended cells changed. |
| Rate / withholding / allowance rules | Run all four scenarios + add a hand-computed expected total for at least one worker and compare. |
| Build script / packaging | `bash _dev/build_app.sh`, then double-click the alias and run a real export. |

## When you cannot verify

If real `Raw/` files for the period are not available (e.g. you are working from a clean checkout), say so explicitly. Do not claim a scenario passed when it was not run.

Anything that depends on the macOS Tkinter event loop, code signing, or the iCloud Drive path cannot be verified from CI or from another OS.

## Manual GUI checklist

When changing the GUI:

- Open the app from the project root with `python3 _dev/payroll_app.py`.
- Confirm `~/.optihome_payroll_config.json` is created on first save and re-read on next launch.
- Toggle Advanced Settings → Timeclock visibility and confirm the picker shows/hides.
- Pick a date whose `MM-DD-YYYY` matches files in `Raw/<year>/` and confirm auto-fill.
- Run an export and confirm warnings appear in the Output Log before the success line.

## What does NOT need verification

- Cosmetic edits to `docs/`, `README.md`, `AGENTS.md`.
- Edits to `timesheet-rates.csv` (it is operator-owned data; the exporter reads it on every run).
- Edits inside `Raw/`, `Timesheets/`, or `Quarterly Cleaning Expenses/` (also operator-owned data).
