# ADR-0062 — The diagnostic fill belongs to the report: cleared on insert, reconciled by colour

- **Status:** Accepted (owner-confirmed fix 2026-11-16)
- **Related:** ADR-0008 D4 (a host refusal is reported, never thrown), ADR-0060 D6 (the insert-time
  presentation repair), R5.11 (coalesced automatic refresh)

## Context

Owner report (2026-11-16), immediately after the Start|Finish grid rule was restored: "the cell
border is now present but the warning fill is incorrect, e.g. milestones are showing start and
finish as errors."

A milestone reads `Start` only (`EntityDateMode.StartOnly`), so a pastel `Finish` cell on a
milestone row is never a validation finding. The validator can only produce `StartRequiredForPoint`
(Start, Error) for a dateless milestone, and `NotUsedByType` (Finish, Warning) if the Finish cell
holds a value — both pinned by `GanttRowValidatorTests`. The pink Finish was therefore not a finding
at all.

**Measured.** `scripts/probe-insert-border-inheritance.ps1` (phase B) shows the row an Add creates
comes out carrying the previous row's pastel Start/Finish fill, because a visible cell fill travels
with the cell format and `ListRows.Add` formats the new row from the existing body. Three defects in
the fill lifecycle follow:

1. **The new row is born filled.** Nothing removed it: the report had not written it.
2. **The reconcile could only clear what it wrote.** `ReconcileHighlights` compared *addresses*, and
   a fill inherited from another row is not in the tracker, so it was permanent.
3. **The colour never followed the severity.** Because the address was already tracked, a cell whose
   finding changed from Warning to Error (or back) was "unchanged" and skipped — the amber fill of a
   warning outlived the error it became.

## Decision

### D1 — The row-insert adapter clears the new row's fill

`ExcelGanttRowInserter.ClearRowFill` writes `xlColorIndexNone` over the new row after its values are
written. One COM call, on a row that has just been created, so nothing user-authored can be lost. A
host refusal is traced (`gantt.add.fill-clear refused`) and never fails an insert that landed
(ADR-0008 D4).

### D2 — A row whose findings changed is adopted, not trusted

`ReconcileHighlights` clears a reported row's fill (one call) and repaints its cell targets whenever
the row's recorded fills do not already match the finding set. This makes the SHEET, not the tracker,
the authority for the rows the report is speaking about, and it is the arm that removes an inherited
or legacy fill.

### D3 — The tracker records effectiveness, not just addresses

The record becomes `(row, column) -> severity`. A cell is skipped only when the severity it was
painted at still matches. A persisted entry carries no severity (the name format is `row|column`),
so a reopened workbook reports "colour unverified" and the first reconcile re-asserts each tracked
cell once.

### D4 — One cell with two findings takes the worse colour

Two issues can name the same cell — a Start-only type's `Finish` can be both "not used" and
"unsupported value". The desired state is collapsed to the highest severity before painting, so
colour is a rule rather than argument order.

## Consequences

- The unchanged-row path still costs no COM write; the extra writes are one row clear plus that row's
  cells, only when its finding set or severity changed.
- The `IRowWarningHighlighter.ReconcileHighlights` contract text changes: "addresses in both sets are
  skipped" becomes "skipped only when the colour still matches".
- `GanttRowInserterTests` gains two facts (the clear runs after the row is created; a refused clear is
  traced and the Add still succeeds). `ExcelRowWarningHighlighterTests` gains four (the adoption
  clear, the severity recolour, the worst-severity collapse in both orders, and a converged cell row
  costing nothing). Two existing expectations moved with the behaviour, each with its reason in
  place.
- The live fact `AddRowIntegrationTests.An_appended_row_does_not_clone_the_error_fill_of_the_row_above`
  now asserts the reported symptom itself: a milestone added under a pink activity has a clean
  `Finish` cell. It fails without D1 — measured, the cell reads the pastel red — and passes with it.
- A fill on a row the report never mentions — a fully valid row whose fill was inherited before this
  change, or a user's own cell colour — is still not cleared by a Refresh. Initialise and Repair own
  presentation; the next Add clears the row it creates.

## Unknown

Why the host copies a *fill* into the created row when the copy origin is
`xlFormatFromRightOrBelow` — the value the 2026-11-09 change relied on to stop exactly this — is not
established. The measurement is that it does, for the row `ListRows.Add` formats. The clear does not
depend on that mechanism.
