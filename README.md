# Optihome Payroll Processing

Converts turno and timeclock CSVs into per-employee Excel workbooks with hourly totals, cleaning job details, and a summary sheet.

## How to use

1. Double-click **Optihome Payroll Processing.app** in this folder.
2. Verify the **Period end day** and date are correct.
3. The **Turno Report** is auto-filled if a matching file exists in `Raw/`. If not, click **Browse** to select it.
4. Click **Run Export**.
5. The Excel file is created in `Timesheets/` and opens automatically.

## Input files

- **Turno CSV** (`_turno.csv`) — Cleaning job records exported from Turno. Place in `Raw/`.
- **Timeclock CSV** (`_time.csv`) — Punch-in/out records from the timeclock system (optional, under Advanced Settings). Place in `Raw/`.
- **Employee Rates** (`timesheet-rates.csv`) — Lookup table in this folder with columns: ID, NAME, RATE, START, EXTRA, DETAILS.

## Output

An Excel workbook with:
- One sheet per employee showing hours worked and cleaning jobs
- Location sections split into Mango Villas, Casa Damisela, and Other
- A Summary sheet with totals for all employees

## Notes

- Input CSV filenames should contain a date like `MM-DD-YYYY` (e.g., `04-01-2026_turno.csv`). This sets the week range shown on each sheet.
- The `_dev/` folder contains source code and build tools — you can ignore it.
