# Payroll Timesheet Exporter

Python script that converts timeclock CSVs into per-employee Excel workbooks with summary and location sections.

## Requirements
- Python 3.9+ with `pandas` and `xlsxwriter`
- `timesheet-rates.csv` (ID,RATE,START,EXTRA,DETAILS) placed alongside the script

## Usage
```
python export-timesheet.py Timesheets/<output.xlsx> Raw/<input.csv> Raw/<turno.csv>
```

The script parses punch records, applies rates from `timesheet-rates.csv`, builds individual worksheets per person, splits the location section into “Mango Villas” and “Casa Damisela” with placeholder rows, and adds a Summary tab.

## Notes
- Input CSV filenames ending with a date like `...-MM-DD-YYYY.csv` set the week range shown on each sheet.
- Outputs and timeclock CSVs are ignored by git (see `.gitignore`).

## Super simple step-by-step
1) Open the **Terminal** app (Applications → Utilities → Terminal).
2) Copy/paste this and press Enter:
```
cd "/Users/mds/Library/Mobile Documents/com~apple~CloudDocs/AA-MDS-REAL ESTATE/aa-Payroll"
```
3) Make sure the two input files are in the `Raw` folder:
   - The timeclock file (ends with `_time.csv`)
   - The turno file (ends with `_turno.csv`)
4) Copy/paste this command, replacing the filenames with the ones you want:
```
python export-timesheet.py Timesheets/mm-dd-2026.xlsx Raw/mm-dd-2026_time.csv Raw/mm-dd-2026_turno.csv
```
5) Press Enter. The Excel file will appear in the `Timesheets` folder.
