# ADR-0051 — Retain the unprotected worksheet; add-in-owned protection is not adopted

- **Status**: Accepted (owner decision recorded 2026-11-01)
- **Date**: 2026-11-01
- **Supersedes**: none
- **Relates to**: ADR-0008 D4 (guard-first refusal), R2.7a, R4.7C column
  classification, R5.11 refresh/write behaviour, R5.15

## Context

R5.15 asked whether the add-in could protect the single visible Gantt
worksheet so that only approved authoring cells and user actions remain
editable, while the add-in safely maintains engine columns, derived values,
table structure, outlines, validation, shapes, and chart layout.

The workbook contract already classifies every column (`Authoring`,
`ReadOnlyVisible`, `EngineHidden`) and writes `Locked` flags, but those flags
have no effect until the sheet is protected. ADR-0008 D4 requires every
mutating adapter to refuse a protected target before any COM mutation, and
R2.7a makes that refusal typed and test-pinned.

## Decision criteria (stated before the evidence was read)

Adoption would be permitted only if **all** of the following held:

1. The add-in can maintain engine columns, derived values, shapes, and chart
   layout under protection **without bypassing the common mutation guard**
   (ADR-0008 D4).
2. Reopen/reapplication of UI-only protection is **deterministic**.
3. Adoption requires **no** stop condition: no schema change, no second
   helper worksheet, no security promise, no Office-version support change.

Any failed criterion retains the current unprotected workflow.

## Evidence

Live COM probe on the desktop Excel host (Microsoft.Office.Interop.Excel
16.0.0 PIA, `Worksheet.Protect(..., UserInterfaceOnly:=$true)`):

| Operation | Result |
| --- | --- |
| `Worksheet.Protect(..., UserInterfaceOnly:=$true)` | `ProtectContents = True` |
| COM write to a locked cell (`Range.Value2`) | **Blocked** — `COMException`: "The cell or chart you're trying to change is on a protected sheet." |
| `UserInterfaceOnly` persistence across close/reopen | **Unproven** — known Excel limitation; not demonstrated to survive reopen |

`UserInterfaceOnly = true` does not permit COM interop writes to locked cells
from outside Excel's object model. An Excel-DNA add-in runs inside Excel, but
its own shape/table/cell writes are COM writes and are subject to the same
protection, so criterion 1 fails unless the guard is bypassed (prohibited) or
an unprotect/reprotect window is introduced (a stop condition, and unsafe).
Criterion 2 is unproven, and the remaining proof plan steps (table
insert/delete/resize, sort/filter, outline, shape writes, failure modes) all
depend on criterion 1 holding.

## Decision

**Retain the unprotected worksheet.** No production protection code is added.
The `GanttTableColumns` `Locked` classification remains inert documentation of
the intended state if a user manually protects the sheet. ADR-0008 D4 and
R2.7a are unchanged: a protected target is refused with a typed, no-mutation
result. Preventive controls — Ribbon gating, validation, warnings, and
confirm-required repair per ADR-0011 — remain the corruption guards.

## Consequences

- The add-in never unprotects, reprotects, or removes a user's protection.
- A user who protects the sheet keeps a working read-only experience for
  reads, while every mutation refuses with the existing typed message.
- If a future Excel-DNA/Office build makes UI-only protection deterministic
  for in-process COM writes, this ADR is the decision to revisit; the R5.15
  proof plan and its criteria remain the template.
- Worksheet protection is not a security or confidentiality boundary in any
  case; no password is treated as a control.
