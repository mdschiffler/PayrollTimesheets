# Payroll Timesheet Exporter

Python script that converts raw timeclock CSVs into per-employee Excel workbooks with summary and location sections.

## Requirements
- Python 3.9+ with `pandas` and `xlsxwriter`
- `timesheet-rates.csv` (ID,RATE,START,EXTRA,DETAILS) placed alongside the script

## Usage
```
python export-timesheet.py Raw/<input.csv> Timesheets/<output.xlsx>
```

The script parses punch records, applies rates from `timesheet-rates.csv`, builds individual worksheets per person, splits the location section into “Mango Villas” and “Casa Damisela” with placeholder rows, and adds a Summary tab.

## Notes
- Input CSV filenames ending with a date like `...-MM-DD-YYYY.csv` set the week range shown on each sheet.
- Outputs and raw CSVs are ignored by git (see `.gitignore`).
