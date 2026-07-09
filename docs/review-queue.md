# Code Review Issue Queue

_Last updated: 2026-07-08_

Prioritized list of known issues. Severity bands:

- **P0** — security exposure, data loss, broken production behavior.
- **P1** — correctness, reliability, scalability, or major UX problems.
- **P2** — cleanup, fragility, minor inconsistencies, low-risk improvements.

Each entry: area, what is wrong, why it matters, suggested fix.

When fixed, move the entry to **Resolved** with a date and short note.

---

## P0

_None known._

## P1

_None known._

## P2

_None tracked yet. Add new entries here as they are discovered. Do not act on a P2 without an explicit go-ahead from the operator._

---

## Resolved

- **2026-07-08 — Turno parser:** standalone `ROOM ONE`–`ROOM FIVE` cleanings (blank `Property Group`) landed in "Other"; they now map to the MARU section, so they also count in Summary "Total Cleans" / "Pay/Job".
- **2026-07-08 — Worker sheets:** removed the "Apt X" / "Details here" placeholder rows from the cleaning sections; empty sections keep one blank row for manual entries.
- **2026-07-08 — GUI:** `payroll_app.py` used `importlib.util` while importing only `importlib` (worked via side-effect imports); now imports `importlib.util` explicitly. Turno tooltip copy updated to include the MARU section.
