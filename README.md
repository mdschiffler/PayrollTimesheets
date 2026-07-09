# Optihome Payroll Processing

Desktop app and command-line exporter for turning payroll CSV reports into a reviewable Excel workbook.

The workbook has a Summary sheet plus one sheet per worker. It supports hourly work from Notion or the timeclock, cleaning jobs from Turno, expense rows from Notion, recurring extras from `timesheet-rates.csv`, and withholding review fields.

## How to use the app

1. Double-click **Optihome Payroll Processing.app** in this folder.
2. Verify the **Period end day**.
3. Select all applicable input reports to process. Use **Advanced Settings** to show or hide report options.
4. Confirm the **Output File**.
5. Click **Run Export**.
6. Review warnings in **Output Log**, then review the generated Excel workbook.

## Input files

- **Notion Report** (`MM-DD-YYYY_notion.csv`) - Bi-weekly contractor hourly timesheet export from Notion. The app reads `Person`, falling back to `Team Member` when `Person` is blank, and converts `Start Time (UTC)` / `End Time (UTC)` to Puerto Rico time.
- **Turno Report** (`MM-DD-YYYY_turno.csv`) - Cleaning job report exported from Turno. Jobs are grouped into Mango Villas, Casa Damisela, MARU, and Other. Rooms named ROOM ONE through ROOM FIVE count as MARU.
- **Expenses Report** (`MM-DD-YYYY_expenses.csv`) - Notion expense export. Rows appear in an **Expenses** section on the worker sheet. Only rows marked `Reimbursable = Yes` add to the final payroll total.
- **Timeclock File** (`MM-DD-YYYY_time.csv`) - Optional NGTecoTime punch export. The app uses the first and last punch per person per day.
- **Employee Rates** (`timesheet-rates.csv`) - Lookup table in this folder with `ID`, `NAME`, `RATE`, `START`, `EXTRA`, and `DETAILS`.

Put report files in `Raw/` or a year folder such as `Raw/2026/`. The app auto-fills files matching the selected period end date.

## Output

Generated workbooks are normally saved in `Timesheets/`.

Each worker sheet includes:

- **Hourly Work** - Notion and timeclock rows paid as `Hours * RATE`.
- **Mango Villas**, **Casa Damisela**, **MARU**, and **Other** - Turno cleaning jobs paid from the Turno cleaning price.
- **Expenses** - Expense rows grouped by `Expensed By`; reimbursable rows are added after withholding.
- **Summary** - totals, recurring extras, withholding, final total, and a reviewed dropdown.

Sections with no data are left off the sheet, so each worker only sees the sections they worked.

The Summary tab rolls up hours, clean counts, totals, withholding, pay/hour, pay/job, and review status.

## Notes

- Notion files use a 14-day period ending on the date in the filename.
- Non-Notion exports keep the existing 7-day period behavior.
- Name matching uses the first two normalized name tokens. A stable employee ID in Notion would make matching safer.

## For maintainers

See [AGENTS.md](AGENTS.md) for the agent workflow rules and an index into [docs/](docs/) (architecture, data model, pipeline, setup, testing, roadmap, and the review queue).

_Last updated: 2026-07-08_
