# Optihome Payroll Processing

Converts timeclock and turno CSVs into per-employee Excel workbooks with hourly totals, cleaning job details, and a summary sheet.

## How to use (Desktop App)

1. Double-click **Optihome Payroll Processing.app** (in the `dist/` folder).
2. Click **Browse** next to "Timeclock CSV" and select the `_time.csv` file from `Raw/`.
3. Click **Browse** next to "Turno CSV" and select the `_turno.csv` file from `Raw/`.
4. The output path is auto-filled. Change it with **Save As** if needed.
5. Verify the **Employee Rates CSV** path points to `timesheet-rates.csv`. Use **Open** to edit rates or **Browse** to pick a different file.
6. Click **Run Export**.
7. The Excel file is created in `Timesheets/` and opens automatically when done.

## How to use (Command Line)

```
python export-timesheet.py Timesheets/MM-DD-YYYY.xlsx Raw/MM-DD-YYYY_time.csv Raw/MM-DD-YYYY_turno.csv
```

## Input files

- **Timeclock CSV** (`_time.csv`) — punch-in/out records exported from the timeclock system. Place in `Raw/`.
- **Turno CSV** (`_turno.csv`) — cleaning job records exported from Turno. Place in `Raw/`.
- **Employee Rates CSV** (`timesheet-rates.csv`) — lookup table with columns: ID, RATE, START, EXTRA, DETAILS. Lives in the project root.

## Output

An Excel workbook with:
- One sheet per employee showing hours worked and cleaning jobs
- Location sections split into Mango Villas, Casa Damisela, and Other
- A Summary sheet with totals for all employees

## Notes

- Input CSV filenames should end with a date like `MM-DD-YYYY` (e.g., `01-28-2026_time.csv`). This sets the week range shown on each sheet.
- Generated spreadsheets and raw CSVs are excluded from git.

## Building the app

Requires Python 3.9+ with Tcl/Tk support and a virtual environment:

```bash
source venv/bin/activate
pip install pandas xlsxwriter pyinstaller
bash build_app.sh
```

Output: `dist/Optihome Payroll Processing.app`
