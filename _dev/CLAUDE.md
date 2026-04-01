# CLAUDE.md — AA Payroll Timesheet Exporter

## Project Purpose

Consolidates timesheet data from two CSV sources (timeclock punch records + cleaning job assignments) into a per-employee Excel workbook. The output spreadsheet is used to manually review and approve payroll before sending payments from the bank website.

## Architecture

Single Python script (`_dev/export-timesheet.py`) with no framework. Takes 3 CLI args:

```
python _dev/export-timesheet.py <output.xlsx> <time_input.csv> <turno_input.csv>
```

### Data Flow

```
Raw/*_time.csv  (timeclock punches)  ─┐
Raw/*_turno.csv (cleaning job data)  ─┼─► _dev/export-timesheet.py ─► Timesheets/*.xlsx
timesheet-rates.csv (employee rates) ─┘
```

### Input Sources

1. **Timeclock CSV** (`_time.csv`) — Exported from NGTecoTime fingerprint system. Columns: `Person ID`, `Person Name`, `Punch Date`, `Attendance record`, `Verify Type`, `TimeZone`, `Source`. One row per punch event; script groups by person+date, uses earliest/latest punch as clock-in/out.

2. **Turno CSV** (`_turno.csv`) — Exported from Booking Calendar. Columns: `Teammate`, `Start Date & Time`, `End Date & Time`, `Cleaning Price`, `Property Alias`, `Property Group`, and others. Each row is a completed cleaning job.

3. **Rates CSV** (`timesheet-rates.csv`) — Manual lookup table in project root. Columns: `ID`, `RATE`, `START`, `EXTRA`, `DETAILS`. Keyed by employee Person ID.

### Output Structure (Excel Workbook)

- **Summary sheet** (first tab): Per-employee totals with cross-sheet formulas (hours, cleans count, total $, reviewed status).
- **Per-employee sheets** (one tab each, named `{ID} - {Name}`):
  - Header with person info and pay period dates
  - Main timesheet table (Location, Date, Start, End, Hours, Details)
  - Rate and Total $ rows
  - Three location sections: **Mango Villas**, **Casa Damisela**, **Other** — each with turno job rows + placeholder rows + subtotal
  - Summary block: Extras, Subtotal, Annual withheld, 10% withheld, Final Total, Reviewed dropdown

### Key Business Logic

- **Pay period**: Parsed from input filename date (`MM-DD-YYYY`); period is that date minus 6 days.
- **Location mapping**: Property names containing "MANGO" → Mango Villas; "DAMISELA" → Casa Damisela; else → Other.
- **Name matching**: Unicode-normalized, uppercase, first+last token match between turno teammate names and timeclock person names.
- **Withholding**: If employee hired <28 days ago OR current month is January → $0 withheld (soft red highlight). Otherwise → $500 annual limit, 10% of subtotal withheld per period.
- **Rates lookup**: Loads `timesheet-rates.csv` from the parent directory of the input CSV file (i.e., project root when input is in `Raw/`).

## Dependencies

- Python 3.9+
- `pandas` — CSV parsing and data manipulation
- `xlsxwriter` — Excel workbook generation (via `pd.ExcelWriter`)
- Virtual environment in `_dev/venv/`

## File Structure

```
timesheet-rates.csv                # Employee rate/bonus lookup (ID,RATE,START,EXTRA,DETAILS)
README.md                          # User-facing usage instructions
Raw/                               # Input CSVs (gitignored)
Timesheets/                        # Output Excel files (gitignored)
dist/                              # Built .app bundle (gitignored)
_dev/
  export-timesheet.py              # Core processing logic (importable + CLI entry point)
  payroll_app.py                   # Tkinter GUI application (wraps export-timesheet.py)
  build_app.sh                     # Builds a macOS .app bundle via PyInstaller
  CLAUDE.md                        # This file (developer documentation)
  OLD_export_timesheet.py          # Legacy version (no turno support, reference only)
  venv/                            # Python virtual environment (gitignored)
  build/                           # PyInstaller build artifacts (gitignored)
```

## Development Commands

```bash
# Activate virtual environment
source _dev/venv/bin/activate

# Run the export (CLI)
python _dev/export-timesheet.py Timesheets/MM-DD-YYYY.xlsx Raw/MM-DD-YYYY_time.csv Raw/MM-DD-YYYY_turno.csv

# Run the GUI (development mode)
python _dev/payroll_app.py

# Install dependencies (if setting up fresh)
pip install pandas xlsxwriter

# Build the macOS .app bundle (requires pyinstaller)
bash _dev/build_app.sh
# Output: dist/Optihome Payroll Processing.app + Finder alias at project root
```

## GUI Application (`_dev/payroll_app.py`)

Tkinter desktop app for non-technical users. Provides:

- Green branded header with title and instructions
- File picker for the Turno CSV (`_turno.csv`) as primary input
- Save-as dialog for the output Excel file (auto-suggests filename from date)
- Collapsible "Advanced Settings" section containing:
  - File picker for the Timeclock CSV (`_time.csv`)
  - File picker for the Employee Rates CSV with Browse/Open
- "Run Export" button that calls `process_timesheet()` in a background thread
- Status area showing success, warnings, and errors with colour coding
- Default browse directories: `Raw/` for inputs, `Timesheets/` for output

The GUI imports `process_timesheet()` from `export-timesheet.py` via `importlib` (needed due to the hyphenated filename). The function raises exceptions on fatal errors and returns a `(message, warnings)` tuple on success, which the GUI displays in-app.

### Building the .app Bundle

Run `bash _dev/build_app.sh` to produce `dist/Optihome Payroll Processing.app`. This uses PyInstaller to bundle Python, dependencies, and the processing script into a standalone macOS app. The build also creates a Finder alias at the project root so the user can double-click the app without navigating into `dist/`.

The `timesheet-rates.csv` is bundled but also checked for at runtime next to the app, so it can be updated without rebuilding.

## Conventions

- Date format in filenames: `MM-DD-YYYY` (e.g., `01-21-2026_time.csv`)
- Time CSV columns use either `YYYY-MM-DD` or `MM/DD/YYYY` date formats (both supported)
- Turno CSV datetimes are `YYYY-MM-DD HH:MM:SS AM/PM`
- Employee IDs can be short numeric (e.g., `18353`) or long numeric (e.g., `3343896104`)
- Sheet names truncated to 31 chars (Excel limit)
- All monetary values use `$#,##0.00` format
- Generated Excel files and raw CSVs are gitignored; only source code and rates CSV are tracked

## Warnings & Error Handling

- Missing input files → exit with error
- Missing rates CSV → warning, defaults all rates to $0
- Unparseable datetimes → skipped with warning
- Turno rows with missing/unmatched teammate → printed warning, row skipped
- Ambiguous name matches (multiple employees) → printed warning, row skipped
