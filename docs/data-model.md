# Data Model

_Last updated: 2026-05-21_

## Sources of truth

- **Employee identity, rate, recurring extras, withholding details:** [timesheet-rates.csv](../timesheet-rates.csv).
- **Hourly contractor work:** Notion CSV export (`*_notion.csv`).
- **Cleaning jobs:** Turno CSV export (`*_turno.csv`).
- **Time-clock punches (optional, currently hidden by default in GUI):** NGTecoTime CSV export (`*_time.csv`).
- **Generated workbook:** `Timesheets/<MM-DD-YYYY>.xlsx`.

All inputs are local files. There is no database and no remote write path. The repo's `.gitignore` excludes `Raw/`, `Timesheets/`, `Quarterly Cleaning Expenses/`, and `*.xlsx`; raw data and outputs are not version controlled.

## `timesheet-rates.csv`

| Column   | Type        | Required | Notes |
|----------|-------------|----------|-------|
| `ID`     | string/int  | Yes      | Employee ID. Matched against Notion `Person` and Timeclock `Person ID`. Trailing `.0` is stripped (pandas coercion artifact). |
| `NAME`   | string      | Yes      | Display name. Used for name-fallback matching via first two normalized tokens. |
| `RATE`   | number      | Yes      | Hourly rate in USD. Defaults to 0 if missing or unparseable. |
| `START`  | `MM-DD-YYYY` | Yes     | Employee start date. Drives the no-withholding allowance gate (see Withholding rules below). |
| `EXTRA`  | number      | Yes      | Recurring per-check dollar amount (supervisor stipend, gas, etc.). |
| `DETAILS`| string      | No       | Free-text description shown next to `EXTRA` on the worker sheet. |

Invariants:

- One row per employee. Duplicate IDs are not actively guarded — keep them unique.
- `NAME` should be the form Notion / Turno actually use, because name-based fallback matches on the first two normalized tokens (`name_key` in `_dev/export-timesheet.py:30`).
- `START` is parsed by pandas; ambiguous formats may silently become `NaT`, which then forces the $500 allowance default. Use `MM-DD-YYYY`.

## Notion CSV (`*_notion.csv`)

Required columns: `Start Time (UTC)`, `End Time (UTC)`, `Hours (calc)`, and at least one of `Person` / `Team Member`.

Optional columns surfaced into the workbook: `Date`, `Status`, `Category`, `Property`, `Notes`, `Time Log URL`.

Rules:

- `Person` is preferred; `Team Member` is the fallback when `Person` is blank. Fallback rows are counted and warned at the end of the run.
- `Start Time (UTC)` and `End Time (UTC)` are parsed as UTC and converted to `America/Puerto_Rico`.
- Rows with missing names, unparseable times, or non-positive hours are skipped with a warning.
- Period when any Notion file is present: 14 days ending on the date in the output/input filename.

## Turno CSV (`*_turno.csv`)

Required columns: `Teammate`, `Start Date & Time`, `End Date & Time`, `Cleaning Price`, `Property Alias`, `Property Group`.

Rules:

- Location bucket is derived from `Property Group` + `Property Alias`: `MANGO` → "Mango Villas", `DAMISELA` → "Casa Damisela", `MARU` → "MARU", else "Other".
- When multiple teammates appear on the same `Property Alias` + date, the cleaning price is split evenly across them (`_parse_turno`, `_dev/export-timesheet.py:414`).
- Hours computed from start/end. If outside `[0.25, 5]` they are clamped to a flat 2.0 (defensive default for malformed exports).
- Rows missing teammate, with invalid name tokens, or with missing start/end are skipped with a warning.

## NGTecoTime CSV (`*_time.csv`)

Required columns: `Person ID`, `Person Name`, `Punch Date`, `Attendance record`.

Rules:

- Per person, per `Punch Date`, the earliest punch is treated as check-in and the latest as check-out.
- Location is hard-coded to "Maru".
- Rows with unparseable dates or non-positive hours are skipped with a warning.
- The GUI defaults to hiding this source; it must be enabled in Advanced Settings.

## Period semantics

- The period-end date comes from the first `MM-DD-YYYY` match in the output path, then any input path (`_find_date_in_paths`).
- Period length per worker (`_person_period`): 14 days if the worker has any Notion rows, else 7. (For workers with both Notion and Turno rows, it's still 14.)
- The period string appears on each worker sheet and in the Summary `Period` column.

## Withholding and allowance

Per worker, the Summary block contains:

- `Extras $` — `EXTRA` from the rates CSV, with `DETAILS` shown alongside.
- `Subtotal $` — sum of all section totals + extras.
- `No-withholding allowance applied this check $` — defaults to **$500** when the employee's `START` is within the last 28 days OR the current month is January, otherwise **$0**. The cell is validated to `[0, 500]` and shaded red when defaulted to $500 to flag review.
- `10% withheld today $` — `ROUNDDOWN(MAX(Subtotal − Allowance, 0) * 0.10, 2)`.
- `Total $` — `Subtotal − Withheld`.
- `Reviewed` — manual dropdown (`y` / blank).

The $500 cap, the 10% rate, and the 28-day / January rule are baked into the exporter. Changing any of them is a contract change.

## Provenance and auditability

- Every generated workbook is named for its period end date and stored under `Timesheets/`.
- Source CSVs live under `Raw/<year>/`; older inputs are moved into `Raw/z-ARCHIVE/`.
- The exporter does not modify input files. It does not write logs to disk — warnings are returned to the caller (printed by the CLI, shown in the GUI Output Log).
- There is no per-row source URL captured in the workbook today except the Notion `Time Log URL`, which is concatenated into the `Details` column.

## Missing data, honestly represented

- Missing rate → row still appears, hourly pay shows `$0`, and the person is named in a trailing warning. Do not silently impute a rate.
- Ambiguous name match (same first two tokens map to multiple rate rows or to multiple already-seen people) → row is skipped with a warning. Resolve by fixing the source or by aligning `NAME` in the rates CSV.
- Empty workbook (no usable rows) → a warning is added and only the `Summary` sheet is written.
