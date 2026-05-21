# Optihome Payroll Processing — Agent Guide

_Last updated: 2026-05-21_

This is the root guide for any coding agent working in this repo. Read it first.

The repo is the source of truth. If this guide or any `docs/` file disagrees with the code, fix the doc in the same change that touches the code.

## What this project is

A local-only macOS desktop app + Python CLI that turns Notion / Turno / NGTecoTime CSV exports into a per-worker Excel workbook for manual payroll review. Single non-technical operator. No network calls, no database. Full mission, stack, and entry points: [docs/snapshot.md](docs/snapshot.md).

## First principles

- **Understand before changing.** Read the relevant `docs/` file and the affected source file before editing. Only ask the operator when the answer cannot be discovered from the repo and the ambiguity would materially change the implementation.
- **Smallest coherent change.** Prefer the minimum diff that satisfies the request. No drive-by refactors, formatting sweeps, dependency churn, or speculative abstractions.
- **Preserve the worktree.** Unrelated edits may be present. Never revert, overwrite, or reformat work you did not author unless explicitly asked.
- **Follow existing conventions.** The project uses pandas + xlsxwriter, Tkinter `ttk`, single-file modules, and snake_case helpers. Match that.
- **Be honest about uncertainty.** If a check could not run because of missing data, credentials, or services, say so. Do not claim unrun checks passed.

## Workflow rules for any non-trivial change

1. Inspect the relevant source file(s) and the matching `docs/` page.
2. State (or note in the conversation) the smallest safe approach.
3. Make focused changes using existing patterns.
4. Run the required checks for the change's risk band — see [docs/testing.md](docs/testing.md).
5. Update any `docs/` page whose content drifted from the new behavior, and bump that page's `Last updated` date.
6. Report what changed, what was verified, what was skipped, and any residual risk.

For ambiguous work, write a short plan first (goal, success criteria, in/out of scope, public-interface impact, edge cases, verification). Confirm with the operator before implementing.

For code review, prioritize: security → data loss → correctness → performance → UX regressions → missing tests → cleanup. Track ongoing items in [docs/review-queue.md](docs/review-queue.md).

## Hard rules — do not break

- **No new external dependencies, services, scheduled jobs, network calls, or paid APIs** without an explicit request. See the non-goals list in [docs/roadmap.md](docs/roadmap.md).
- **No destructive operations** without an explicit `--apply` / confirmation flag. The exporter overwrites the target `.xlsx`; that is the only destructive write today and it operates on a derived artifact.
- **Do not fabricate missing data.** Skip rows with a warning rather than inventing values. Missing rates produce $0 hourly pay and a warning, never an invented rate.
- **Do not modify input CSVs** under `Raw/` or operator-owned data under `Timesheets/` and `Quarterly Cleaning Expenses/`.
- **Do not introduce auth, billing, accounts, multi-user features, cloud sync, web UIs, mobile UIs, or background jobs.**
- **Do not commit anything under `Raw/`, `Timesheets/`, `Quarterly Cleaning Expenses/`, `*.xlsx`, `dist/`, `_dev/build/`, `_dev/venv/`, or `.DS_Store`** — see [.gitignore](.gitignore).

## Required check after edits

After editing `_dev/export-timesheet.py` or `_dev/payroll_app.py`:

```bash
python3 -m py_compile _dev/export-timesheet.py _dev/payroll_app.py
```

For risk-matched verification beyond `py_compile`, see [docs/testing.md](docs/testing.md).

## Documentation index

Each file below carries its own `Last updated` date. If you change behavior that a file describes, update the file and its date in the same change.

| File | When to read it |
|---|---|
| [docs/snapshot.md](docs/snapshot.md) | Mission, primary user, stack, source-of-truth map, entry points. |
| [docs/architecture.md](docs/architecture.md) | Module layout, pipeline stages, state flow, public contracts. |
| [docs/data-model.md](docs/data-model.md) | CSV schemas, period semantics, withholding rules, missing-data handling. |
| [docs/pipeline.md](docs/pipeline.md) | Operator workflow, file conventions, what to watch for at run time. |
| [docs/product-ux.md](docs/product-ux.md) | User, workflows, copy rules, visible states, accessibility. |
| [docs/setup-commands.md](docs/setup-commands.md) | Prerequisites, install, dev commands, CLI reference, build script. |
| [docs/testing.md](docs/testing.md) | Required checks per risk band, smoke scenarios, manual GUI checklist. |
| [docs/roadmap.md](docs/roadmap.md) | In-scope, explicit non-goals, open ideas, constraints. |
| [docs/review-queue.md](docs/review-queue.md) | Prioritized known issues (P0 / P1 / P2) and resolved log. |
| [README.md](README.md) | Wife-facing usage. Keep its tone plain and non-technical. |

## Maintenance guidance

- Keep [README.md](README.md) focused on the operator workflow. Do not push agent-targeted detail into it.
- Keep the canonical technical guidance in `docs/`. The root `AGENTS.md` (this file) is workflow + rules + index — do not duplicate long technical content here.
- `_dev/CLAUDE.md` is a stub pointer; do not move content into it.
- Preserve local raw CSVs, generated workbooks, the rates CSV, and any unrelated worktree changes.
