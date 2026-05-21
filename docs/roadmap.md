# Roadmap & Boundaries

_Last updated: 2026-05-21_

## In scope (current)

- Convert Notion / Turno / NGTecoTime CSVs into a per-worker Excel workbook with a Summary sheet.
- Hourly pay, cleaning pay (with split-when-shared logic), recurring extras, no-withholding allowance, and 10% withholding line for manual review.
- Local-only operation: macOS app, iCloud Drive storage, no network calls.
- Single non-technical operator running the bundled `.app`.
- Tkinter GUI with persisted preferences in `~/.optihome_payroll_config.json`.

## Explicit non-goals

These have been considered and intentionally _excluded_. Do not introduce them without a direct request:

- **Tax filing or remittance.** The 10% withholding line is a review aid, not a remittance.
- **Multi-user editing or collaboration.** One operator, one Mac.
- **Cloud hosting / web UI / mobile app.** iCloud Drive sync is sufficient.
- **Authentication, accounts, roles.**
- **Scheduled jobs, cron, background sync.**
- **External API calls or paid services** (Notion API, Turno API, Stripe, SendGrid, etc.).
- **Database persistence.** CSVs and `.xlsx` files are the system of record.
- **Expense reimbursement import.** Out of scope until a real Notion expense export exists. Do not infer the schema.
- **Automatic email / Slack / SMS notifications.**
- **Auto-mutating prior generated workbooks.** Once a workbook is written and the operator has reviewed it, it is treated as the reviewed artifact.
- **Replacing Tkinter with a different UI toolkit.**
- **Cross-platform packaging** (Windows / Linux installers). The CLI works anywhere Python runs; the bundled app is macOS-only by design.

## Open ideas (not committed)

Captured here so they are not lost. None of these are scheduled and none should be implemented without an explicit go-ahead.

- **Stable employee ID in Notion.** Today name-fallback matching uses the first two normalized name tokens. A real ID column in the Notion export would remove the ambiguity warnings. (Mentioned in [README.md](../README.md) "Notes".)
- **Expense reimbursement import.** Wait for a real Notion expense export sample before designing the schema.
- **Automatic archival of old `Raw/` files** into `Raw/z-ARCHIVE/`.
- **Sanity-check report** comparing a new workbook to the previous period's totals (per-worker delta, headcount delta) before approval.
- **Bundled requirements file** (`requirements.txt` or `pyproject.toml`) so the venv is reproducible without remembering the dependency list.

## Constraints to respect

- The operator is non-technical. Any change that requires terminal interaction or knowledge of Python is a regression for the primary user.
- iCloud Drive is the storage layer. Resource forks, `.DS_Store` files, and sync delays are facts of life; the build script already works around them.
- The repo is open from `~/Library/Mobile Documents/com~apple~CloudDocs/AA-MDS-REAL ESTATE/aa-Payroll/`. Path-dependent assumptions belong in setup docs, not in the code.

## How to propose a change to scope

1. Add an entry under "Open ideas" with the problem and the smallest plausible shape of a solution.
2. Update this file's `Last updated` date.
3. Do not implement until the operator (the project owner) confirms.
