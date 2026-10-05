# ADR-0036 — Every insert shifts the sheet: a worksheet row, never `ListRows.Add(position)`

- **Status:** Accepted
- **Date:** 2026-10-03 (D5/D6 added the same day after a second live probe)
- **Relates to:** ADR-0035 D3; ADR-0026 D3; ADR-0019; R2.8; R4.7D
- **Decided by:** implementation, from live-host measurement. No product-contract change.
- **Supersedes:** nothing. **Amends:** ADR-0035 (D3's mechanism, and two recorded claims withdrawn as below).

## Context

A live defect was reported: adding an activity moved the table's row data but left the cells and the Gantt shapes behind, and the padding row was visibly interfered with — a body row carrying a description, its text clipped and colliding with the row beneath it.

ADR-0035 D3 described the append branch as `ListRows.Add()` **followed by** an explicit `Worksheet.Rows[lastRow + 1].Insert(xlShiftDown, xlFormatFromLeftOrAbove)`. That order is the defect.

`ListRows.Add()` with **no** position does not shift the sheet. It **claims** the row already sitting below the table — the reserved 6pt bottom padding row — and converts it into a body row. The compensating insert that followed could not repair it: it created a *new* blank row *below the table*, while the already-claimed row stayed a 6pt body row forever.

ADR-0035 also recorded two host claims that shaped the code and had never been re-tested:

1. That writing `Shape.Placement = xlMoveAndSize` was unnecessary because a live probe found `AddTextbox`'s default already `xlMoveAndSize`.
2. That Excel **refuses** a worksheet row inserted inside a `ListObject` range, "which the live probe hit directly".

Claim 1 was generalised from **one** shape family to four. Claim 2 conflicts with what a direct measurement now shows. Neither claim was re-examined when the defect was reported, because both were recorded in an accepted ADR and read as settled.

## Evidence

Measured 2026-10-03 on Excel 16.0 x64 by `scripts/probe-rowinsert-anchoring.ps1` (a disposable spike; its output is recorded here and the script is not a committed test).

**Q1 — `Shape.Placement` immediately after creation, all four families the renderer actually uses:**

| Family | Creation call | Placement |
| --- | --- | --- |
| Bar (rectangle) | `AddShape` | `xlMoveAndSize` |
| Delineator / grid line | `AddLine` | `xlMoveAndSize` |
| Label | `AddTextbox` | `xlMoveAndSize` |
| Milestone (diamond) | `AddShape` | `xlMoveAndSize` |

**Q2 — a real worksheet row insert, at a row *inside* the `ListObject` range, SUCCEEDED.** Every shape moved down exactly one body-row height and `TopLeftCell` followed: bar 60→78, line 80→98, label 60→78, diamond 60→78 (delta 18 on all four); `D4→D5`, `D6→D7`, `F4→F5`, `H4→H5`.

**Q3 — worksheet row FIRST, then `ListRows.Add()`:** new body row **18pt** (correct), row below the table **6pt** (correct), with no height write at all.

**Q4 — the shipped order, `ListRows.Add()` FIRST then the compensating insert:** new body row **6pt**, and **still 6pt** after the compensating insert; padding row also 6pt.

So: claim 1 above is **confirmed** for all four families, not one. Claim 2 is **withdrawn as wrong** — Excel does not refuse the insert; the earlier probe hit a parameter-count error, which was a call-signature problem misread as a host policy. The defect is Q4, and it is purely an ordering error.

## Decision

- **D1 — the append branch inserts the genuine worksheet row BEFORE `ListRows.Add()`.** The padding row is displaced downward carrying its own 6pt, and `ListRows.Add()` then claims the *fresh* row, which inherited 18pt from the body row above it. One insert, one add, both rows correct.
- **D2 — no `Shape.Placement` write, and no shape repositioning.** All four families are natively `xlMoveAndSize`, so the chart tracks the sheet with no add-in code on the insert path. This honours the invariant that normal worksheet changes never render. `ExcelShapeWriter` is unchanged by this decision; ADR-0035's rejection of that write stands, now on four-family evidence rather than one.
- **D3 — the new body row's height is written explicitly** from the `GanttRowHeightPt` token. Q3 shows Excel's inheritance already produces the right value, so this is not required for correctness today. It is written anyway because the inheritance is a host behaviour rather than a stated invariant, and the defect it guards is invisible to a test that does not assert it.
- **D4 — the positional branch is unchanged.** `ListRows.Add(position)` shifts the sheet by itself, so it must **not** also insert a row from below; doing so would displace everything twice for one activity. A test now pins that asymmetry.

### D5/D6 — added 2026-10-03 after the first fix was reported still broken

D4 above was **wrong**, and the tests that encoded it were wrong with it. A live report after D1–D4 shipped showed a mid-table Add Activity still leaving the cells and Gantt shapes behind.

A second probe (`scripts/probe-positional-insert.ps1`) measured the branch the first probe never tested:

**Q1 — `ListRows.Add(position)` shifts NOTHING.** Adding at position 2 moved all three probe shapes by **delta=0** (`TopLeftCell` unchanged) and pushed the row below the table from 6pt to 15pt. It rearranges rows *inside* the table and consumes the padding row. This is precisely the reported symptom: the table's row data moves, the cells and shapes do not. **ADR-0035 D3's claim that the positional branch "already inserts a real worksheet row" is false**, and D4's reasoning — that a second insert would double-shift — was built on it.

**Q2 — a worksheet row inserted INSIDE the table's range does everything needed.** `Rows(3).Insert(xlShiftDown)` moved all three shapes by **delta=18** with `TopLeftCell` following, **and** auto-expanded the `ListObject` (ListRows 3 → 4).

**Q3 — the auto-expanded row is a writable data row.** Its cell accepted a written value.

**Q4 — a row inserted BELOW the table does not auto-expand it** (ListRows stayed 3). That is why the append branch still pairs its insert with `ListRows.Add()`.

- **D5 — the positional branch inserts a genuine worksheet row INSIDE the table's range and reads the row back** (`GetListRowAt`) instead of calling `ListRows.Add(position)`. The `ListObject` absorbs the row, so no `ListRows` call is needed. The dead `AddRowAtPosition` seam is deleted.
- **D6 — `InsertWorksheetRow(table, row)` is extracted as the shared insert**, used by both branches: inside the table for a positional insert, one past the last row for an append.

So **both** branches now shift the sheet exactly once, and the sheet is shifted by the one mechanism that is measured to move shapes.

### The test gap, restated

The pre-existing theory `An_in_table_active_cell_inserts_below_it_and_an_outside_one_appends` asserted that a positional insert must **not** shift the sheet, because `ListRows.Add(position)` "already inserts a real worksheet row". That test **encoded the defect as a requirement** and passed green against a visibly broken sheet. It now asserts one worksheet-row insert on either branch. A source-scan test that pinned the presence of `rows.Add(position)` was inverted to pin its absence.

## Consequences

- **A test asserts the ORDER, not the presence of both calls.** The old code made both calls and was still wrong, so a presence check passes straight through the defect. Mutating the order back fails `The_append_branch_inserts_the_worksheet_row_before_claiming_it_with_ListRows` with `order was: AddRow -> InsertWorksheetRowBelowTable`.
- **The live integration gate now asserts the BODY row's height**, which it never did. Every existing assertion — including the padding row's 6pt — passed while the sheet was visibly broken. That gap is why the defect reached production, and it is now closed.
- **`InsertWorksheetRowBelowTable` no longer writes a height.** Writing `ChartPaddingRowHeightPt` to the row it creates would set the height of the row that is about to become a *body* row, reproducing the defect inside the method meant to prevent it.
- **Non-vacuity proven by mutation.** Reverting the order fails the ordering test and the refused-padding-report test; removing the height write fails both height tests with `The collection was empty`. Each failure is the exact defect predicted.
- **No product contract, schema, or dependency change.** ADR-0029's schema version is untouched: nothing about what the workbook stores changed.
- **Method note.** ADR-0035 generalised a one-family probe to four families and recorded a host refusal that a two-line probe disproves. "The probe covered one family" was the visible tell, and it is why this ADR records the family list per row rather than a single conclusion.

## Alternatives considered

- **Write `Shape.Placement = xlMoveAndSize` explicitly** (rejected: measured unnecessary for all four families — D2. Adding it would be a no-change presented as a fix, which is the failure mode ADR-0035 already identified and correctly avoided).
- **Reposition every shape after the insert** (rejected: duplicates work Excel already does, and would require a render on a row change, breaking the "never render on change" invariant).
- **Keep the compensating insert but move the height write to the claimed row** (rejected: it repairs the symptom and leaves the padding row's identity ambiguous — the padding row would be consumed by the table and re-created, so the chart's bottom margin becomes a different row each time. D1 avoids consuming it at all).
- **Pass `AlwaysInsert` to `ListRows.Add`** (rejected: unreachable. The PIA exposes exactly `ListRows.Add(System.Object)`, verified by reflection on `Microsoft.Office.Interop.Excel` 16.0.0 — the second parameter does not exist in the interop signature).