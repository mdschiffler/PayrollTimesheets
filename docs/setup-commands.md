# Setup & Commands

_Last updated: 2026-05-21_

## Prerequisites

- macOS (the app is built and signed as a `.app` bundle; the CLI is platform-agnostic Python).
- Python 3.9+ (for `zoneinfo`). Project was last built and exercised on the system Python that lives in `_dev/venv/`.
- Dependencies (installed into the venv):
  - `pandas`
  - `xlsxwriter`
  - `pyinstaller` (for app builds only)

No `requirements.txt` or `pyproject.toml` ships with the repo today. If you create one, list it in [snapshot.md](snapshot.md) under "Source of truth".

## One-time setup

```bash
cd "/Users/mds/Library/Mobile Documents/com~apple~CloudDocs/AA-MDS-REAL ESTATE/aa-Payroll"
python3 -m venv _dev/venv
source _dev/venv/bin/activate
pip install pandas xlsxwriter pyinstaller
```

The venv directory is gitignored.

## Day-to-day commands

Run all commands from the project root.

### Run the GUI (development)

```bash
source _dev/venv/bin/activate
python3 _dev/payroll_app.py
```

### Run the exporter CLI

Preferred flag form:

```bash
python3 _dev/export-timesheet.py \
  --output /tmp/04-22-2026.xlsx \
  --notion Raw/2026/04-22-2026_notion.csv \
  --turno Raw/2026/04-22-2026_turno.csv \
  --rates timesheet-rates.csv
```

Legacy positional form (still supported, kept for back-compat — do not rely on it for new scripts):

```bash
python3 _dev/export-timesheet.py Timesheets/04-22-2026.xlsx Raw/2026/04-22-2026_time.csv Raw/2026/04-22-2026_turno.csv
```

At least one of `--notion`, `--turno`, `--time` is required. `--rates` is optional — if omitted, the exporter walks up from the first input file looking for `timesheet-rates.csv`.

### Rebuild the macOS app

```bash
bash _dev/build_app.sh
```

This:

1. Activates `_dev/venv` if present.
2. Installs PyInstaller if missing.
3. Cleans `_dev/build/` and `dist/Optihome Payroll Processing.app`.
4. Runs PyInstaller bundling `_dev/payroll_app.py`, `_dev/export-timesheet.py` (as data), `timesheet-rates.csv` (as data), and Tkinter.
5. Re-signs the bundle (copies to `/tmp` first to dodge iCloud resource forks).
6. Refreshes the Finder alias `Optihome Payroll Processing.app` at the repo root.

## Environment variables

None. The project intentionally has no environment-based configuration. The GUI persists user preferences in `~/.optihome_payroll_config.json`.

## File-system conventions

- Inputs live under `Raw/<year>/` with the date pattern `MM-DD-YYYY` in the filename.
- Outputs live under `Timesheets/` with the same date pattern.
- Older inputs are moved under `Raw/z-ARCHIVE/` by hand.
- `Quarterly Cleaning Expenses/` is operator-maintained and outside the exporter's scope.

## What is gitignored

See [.gitignore](../.gitignore):

- `__pycache__/`, `*.pyc`
- `_dev/venv/`
- `.DS_Store`
- `*.xlsx`, `Raw/`, `Timesheets/`, `Quarterly Cleaning Expenses/`
- `_dev/build/`, `dist/`, `*.spec`

Note: `.DS_Store` is in `.gitignore`, but at least one was committed in the past. If `git status` shows a tracked `.DS_Store`, untrack it with `git rm --cached <path>`. Do not do this unprompted — see [review-queue.md](review-queue.md).
