# Data Pipeline & Operations

_Last updated: 2026-05-21_

## Recurring workflow (operator)

1. Export new CSVs for the period:
   - Notion contractor timesheet → save as `Raw/<year>/<MM-DD-YYYY>_notion.csv`.
   - Turno cleaning report → save as `Raw/<year>/<MM-DD-YYYY>_turno.csv`.
   - Notion expenses report → save as `Raw/<year>/<MM-DD-YYYY>_expenses.csv`.
   - (Optional) NGTecoTime punch export → save as `Raw/<year>/<MM-DD-YYYY>_time.csv`.
   The `MM-DD-YYYY` portion is the period-end date and drives auto-fill in the GUI.
2. Open the app, confirm period end day, confirm auto-filled inputs, confirm output path under `Timesheets/`.
3. Click **Run Export** and read the Output Log for warnings.
4. Open the generated workbook, review each worker sheet, set `Reviewed = y`, then approve payroll.

The wife-facing version of this lives in [README.md](../README.md).

## Manual operations

- **Quarterly expense PDFs:** `Quarterly Cleaning Expenses/` is still maintained by hand. The payroll exporter imports period expense CSV rows into the generated workbook; it does not generate quarterly PDF reports.
- **Quarterly cleanup:** old `Raw/` files are moved into `Raw/z-ARCHIVE/` by hand. There is no automation.
- **Rate table edits:** the operator edits `timesheet-rates.csv` directly in a spreadsheet app or text editor.

## Imports, integrations, jobs

- **No scheduled jobs.** The app runs only when the user clicks Run Export.
- **No external API calls.** Everything is local-file processing.
- **No credentials** are stored, requested, or required by this project.
- **No network access** is required at runtime. Notion / Turno / NGTecoTime exports are downloaded by the user out-of-band.

## Risky-operation defaults

This project has no destructive operations in the current scope. If a future change adds any, the conservative defaults are:

- Dry-run by default. Require an explicit `--apply` / confirmation flag for writes that mutate inputs or external systems.
- Never overwrite a previously generated workbook silently — today the exporter does overwrite the target `.xlsx` because PyInstaller-bundled `pd.ExcelWriter` opens it for writing. This is acceptable because outputs are derived; if you change the semantics so outputs include reviewer edits, add an overwrite guard first.
- Never modify input CSVs.

## External data ownership

- Notion, Turno, and NGTecoTime own their source data. The exporter is a read-only consumer of their exports.
- Expense rows from Notion are listed on each worker sheet. Only rows marked `Reimbursable = Yes` are added after withholding.
- The generated workbook becomes the operator's reviewed artifact and the system of record for what was paid. Treat it as user-owned data — do not overwrite reviewed cells in code, do not auto-mutate prior `Timesheets/*.xlsx` files.

## What to watch for

- **Filename dates drive the period.** If a CSV is exported with a different filename (e.g. missing the `MM-DD-YYYY` suffix), the exporter falls back to 7 days and the per-sheet `Period` line can be blank. Rename before running.
- **Time zone correctness.** Notion timestamps are assumed UTC. If Notion changes export format, every hourly row could shift by 4 hours. Spot-check after any Notion-side change.
- **Name drift.** If a worker's name changes in Notion / Turno but not in `timesheet-rates.csv`, name-fallback matching will silently miss and the row will be paid at $0 with a warning. Update the rates CSV when names change.
- **iCloud sync.** Files live in `iCloud Drive/AA-MDS-REAL ESTATE/aa-Payroll/`. If iCloud has not finished syncing a new CSV, the GUI auto-fill may not see it. Wait for the cloud icon to clear.
