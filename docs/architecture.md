# Architecture

_Last updated: 2026-05-21_

## Module layout

```
aa-Payroll/
  Optihome Payroll Processing.app   Finder alias → dist/Optihome Payroll Processing.app
  README.md                         Wife-facing usage
  AGENTS.md                         Agent rules + index
  timesheet-rates.csv               Employee rate table (editable)
  docs/                             Agent guidance (this folder)
  Raw/                              Input CSVs (gitignored)
    2026/…
    z-ARCHIVE/…
  Timesheets/                       Generated workbooks (gitignored)
  Quarterly Cleaning Expenses/      Manual quarterly reports (gitignored)
  _dev/
    export-timesheet.py             Processing logic + CLI
    payroll_app.py                  Tkinter GUI
    build_app.sh                    PyInstaller build script
    AppIcon.icns                    App icon
    venv/                           Local virtualenv (gitignored)
    OLD_export_timesheet.py         Pre-refactor copy (candidate for removal)
  dist/                             PyInstaller output (gitignored, kept for the .app)
```

## Components

### Exporter — [_dev/export-timesheet.py](../_dev/export-timesheet.py)

Single-file pipeline. Public entry point is `process_timesheet(csv_file, output_excel, turno_csv=None, rates_csv=None, notion_csv=None, expenses_csv=None)`. CLI shim is `_run_cli`.

Pipeline stages, in order:

1. **Validate inputs.** At least one of Notion / Turno / Expenses / Timeclock must be present; each provided path must exist.
2. **Load rates.** `_load_rates` reads `timesheet-rates.csv`, walking up from the source file if no explicit path is given. Builds two maps: `rates_dict` keyed by normalized employee ID, and `rates_by_name` keyed by the first two normalized name tokens.
3. **Parse each source.** Stage-specific parsers populate shared dicts keyed by `(person_id, person_name)`:
   - `persons` — membership set.
   - `hourly_events` — list of hourly rows (Notion + Timeclock).
   - `turno_events` — dict of cleaning rows per location bucket (`LOCATION_BUCKETS = ["Mango Villas", "Casa Damisela", "MARU", "Other"]`).
   - `expense_events` — list of Notion expense rows keyed by `Expensed By`.
4. **Determine period.** `_find_date_in_paths` extracts an `MM-DD-YYYY` date from the output or input filename. `_person_period` picks 14 days if any Notion rows exist, else 7.
5. **Write the workbook.** A `Summary` sheet plus one sheet per person, built section by section: Hourly Work → location sections (when applicable) → Other → Expenses (when applicable) → per-sheet Summary block (totals, extras, allowance, 10% withheld, final total, reviewed flag).
6. **Emit warnings.** Missing rates, unparseable dates, ambiguous name matches, empty files, etc. are accumulated and returned alongside the success message.

Name matching uses `name_key` (`_dev/export-timesheet.py:30`): NFKD-normalized, uppercase, alpha-only, first two tokens. The same tokens are used for both rate lookup and de-duping people seen across sources.

### GUI — [_dev/payroll_app.py](../_dev/payroll_app.py)

Tkinter app that wraps the exporter. Notable behaviors:

- Loads `export-timesheet.py` by path because the hyphenated filename is not importable by name (`_import_export_module`, `_dev/payroll_app.py:60`).
- Resolves the project root by walking up from the script or `.app` bundle looking for `timesheet-rates.csv` (`_get_project_dir`, `_dev/payroll_app.py:35`).
- Persists settings in `~/.optihome_payroll_config.json`: period end day-of-week, last-used folders, and which source pickers are visible.
- Auto-fills source files in the configured `Raw/` folder matching the period-end date.
- Runs the export on a worker thread; appends warnings and the success line to the Output Log.

### Build — [_dev/build_app.sh](../_dev/build_app.sh)

PyInstaller invocation that bundles `payroll_app.py` plus `export-timesheet.py` as a data file, produces `dist/Optihome Payroll Processing.app`, and refreshes the alias at the repo root.

## State and data flow

```
Notion CSV   ┐
Turno CSV    ├─► _parse_* ──► persons / hourly_events / turno_events / expense_events ──► per-sheet writer ──► .xlsx
Expenses CSV │                       ▲
Timeclock    ┘                       │
                                   │
                  timesheet-rates.csv (rates_dict, rates_by_name)
```

The exporter holds the entire run in memory; there is no streaming, no checkpointing, no partial writes.

## Public interfaces (stable contracts)

- CLI: the `--output`, `--notion`, `--turno`, `--expenses`, `--time`, `--rates` flags and the legacy positional form (`<output> <timeclock> [turno]`).
- Library: `process_timesheet(...)` signature and return shape `(message: str, warnings: list[str])`.
- Workbook shape: `Summary` sheet column layout (Person, Role, Period, Total Days, Total Hours, Total Cleans, Total $, Withheld $, Pay/Hour, Pay/Job, Reviewed) and the per-sheet section order (Hourly Work → Mango Villas → Casa Damisela → MARU → Other → Expenses when applicable → Summary block).
- Rates CSV: column names `ID`, `NAME`, `RATE`, `START`, `EXTRA`, `DETAILS`.

Changing any of these is a contract change — update [data-model.md](data-model.md), [pipeline.md](pipeline.md), and [setup-commands.md](setup-commands.md) in the same PR.
