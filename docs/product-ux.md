# Product & UX

_Last updated: 2026-05-21_

## Who uses it

- **Primary user:** one non-technical operator (the project owner's wife), running the macOS app from iCloud Drive on a personal Mac. She is the only routine user. Optimize for her flow.
- **Secondary user:** the project owner, occasionally running the CLI or editing `timesheet-rates.csv`.
- **No other audiences.** No web users, no multi-tenant, no external collaborators.

## Workflows the GUI must serve

1. **Fortnightly payroll run.** Open the app, accept the auto-filled period end and source files, click Run Export, scan warnings, open the workbook.
2. **Mid-period spot-check.** Same flow, but maybe with only one source file present. The exporter must accept any non-empty combination of Notion / Turno / Expenses / Timeclock.
3. **Onboarding a new worker.** The operator edits `timesheet-rates.csv` (in Numbers / Excel / a text editor) and re-runs. New `START` dates within 28 days trigger the $500 no-withholding allowance default (shown shaded red for review).

## Copy

There is no copy dictionary, CMS, locale, or message catalog. User-facing strings live directly in `_dev/payroll_app.py` (GUI labels, dialogs, button text, log messages) and in `_dev/export-timesheet.py` (warnings, summary headers, section titles). Project is English-only by design.

When changing user-visible strings:

- Keep wording plain. The operator is non-technical.
- Match the existing tone in [README.md](../README.md) — short, imperative, no jargon.
- If a warning identifies a person, include the name (and ID when available) so the operator can act without opening source files.

## Visible states the app must handle

- **Empty:** no source file selected → Run Export is disabled / produces a clear error. (Current behavior: `process_timesheet` raises `ValueError` if no input is provided. The GUI should surface that in the Output Log.)
- **Partial inputs:** any subset of the four sources present → exporter must still produce a workbook covering only that data.
- **No usable rows:** valid CSVs but every row was skipped → workbook contains only the Summary sheet and a warning explains why.
- **Loading / running:** export runs on a worker thread; the button should not look frozen. The Output Log appends progress.
- **Error:** missing file, malformed CSV, missing columns → user-facing message names the file and the missing column, not a stack trace.
- **Success with warnings:** the workbook still opens; warnings are listed in the Output Log so the operator can decide whether to fix and re-run.

## Layout & responsiveness

- The app runs in a fixed desktop window on macOS. There is no mobile or web target.
- Window must remain usable on small laptops (no off-screen controls).
- Tkinter `ttk` widgets only; do not introduce a second UI toolkit.

## Accessibility

- Keep keyboard navigation working: Tab through file pickers and the Run button.
- Use semantic labels on `ttk.Label` + paired controls.
- Do not rely on color alone to convey state (e.g. the red "$500 allowance" cell in the workbook is paired with the cell's value and the row label).
- Maintain readable font sizes; do not shrink the Output Log font below the default.

## Things to NOT add without an explicit ask

- Authentication, accounts, or roles.
- Cloud sync (iCloud already handles file sync).
- Email / Slack / SMS notifications.
- Multi-user editing of the workbook.
- A web UI.
- Auto-upload of generated workbooks anywhere.

See [roadmap.md](roadmap.md) for the non-goals list.
