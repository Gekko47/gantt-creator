---
mode: agent
description: Start a new work item from docs/work-items/TEMPLATE.md
---

# New work item

You are starting work item **${input:ID:ID from docs/03-ROADMAP.md, e.g. R0.1}**.

## Before you write a line of code

1. Read the work item file: `docs/work-items/${input:ID}.md`. If the file does not exist, copy `docs/work-items/TEMPLATE.md` and fill the sections.
2. Read only the relevant sections of `docs/02-ARCHITECTURE.md`, `docs/04-TEST-STRATEGY.md`, and `docs/07-GANTT-ENTITY-GUIDE.md` (if visual).
3. Restate to the user, in one paragraph: outcome, files likely to change, exclusions, acceptance tests, and uncertainties. Do not edit until the user confirms.
4. If an API, Office behaviour, or package capability is uncertain, verify it in installed metadata, a minimal spike, or primary documentation. Mark unverified statements as `unknown`.

## While you edit

- One work item in progress. Do not opportunistically refactor neighbours.
- Add or change tests in the same commit as behaviour.
- Run the narrowest affected test after each meaningful edit.
- Never disable analyzers, loosen coverage, delete tests, catch-and-ignore, or add arbitrary sleeps to force green output.
- Never call Excel or PowerPoint from `GanttCreator.Core`.
- Save and restore Excel application state in `try/finally` whenever you change it.

## Before you hand back

1. `pwsh ./scripts/verify-quick.ps1` — must exit 0.
2. `pwsh ./scripts/verify.ps1` — must exit 0 if the work item's gate is full.
3. `git diff --check`, `git status --short`, `git diff --stat`, `git diff` — review the diff yourself.
4. Update `docs/STATUS.md` with: ID, branch, started date, outcome, evidence path, next safe action.
5. Report exactly: changed behaviour, important files, commands run and results, residual risks, and the next roadmap item. Do not claim a command ran unless its output was observed.

## Stop conditions

Stop and ask the human if:

- acceptance criteria are missing or contradictory,
- a change would violate a product invariant,
- a new production dependency is required,
- the same failure persists after two evidence-led fixes,
- the diff touches a public schema, supported platform, or export contract.

Do not commit until the human reviews the diff.
