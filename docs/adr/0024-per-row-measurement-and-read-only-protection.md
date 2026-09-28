# ADR-0024 — Measure each row; a read-only measurement does not refuse a protected sheet

- **Status:** Accepted
- **Date:** 2026-09-28
- **Relates to:** entity guide §3, §4; ADR-0008 D4; ADR-0023
- **Depends on:** ADR-0023
- **Plan:** [`REMEDIATION-PLAN-C.md`](../REMEDIATION-PLAN-C.md)

## Context

ADR-0023 made the Core contract able to describe a body whose rows differ in height.
`ExcelPanelGridMeasurement` could not yet supply one, so the capability was unusable
end to end. Three separate things were in the way.

1. **A mixed body was refused.** `ReadRowHeight` read `DataBodyRange.RowHeight` and
   `HasUniformRowHeight` then demanded every body row equal the aggregate. The
   `Range.RowHeight` reference states that a range of differing heights "might return
   the height of the first row or might return Null". `Null`/`DBNull` was catchable —
   that is why the earlier fix existed — but **the "first row's height" case is a
   plain number that passed every check** while describing a body most of which is
   taller.
2. **The header height was the body height.** ADR-0023 added a separate
   `HeaderHeightPt`; the adapter filled it from the body measurement.
3. **A read-only adapter refused a protected worksheet.** The remediation brief called
   this `P2-1` and framed it as command policy leaking into a read-only port.

## Decision

- **D1 — The body aggregate is never consulted.** The adapter walks the body rows and
  reads each row's own height. `ConfirmUniformRowHeight` and `HasUniformRowHeight` are
  **deleted**, not relaxed. The aggregate is unreliable in *both* documented directions
  for a non-uniform body, so detecting the bad case is strictly worse than never
  reading the value.
- **D2 — The header row is measured separately**, from `ListObject.HeaderRowRange`, on
  its own seam.
- **D3 — One unreadable row refuses the whole measurement.** The panel's row positions
  are cumulative, so a single defaulted height would displace every row below it and
  the result would look plausible. `InvalidMeasurement` is the refusal.
- **D4 — The protection guard is removed from this adapter, and the constructor
  parameter with it.** ADR-0008 D4 requires the guard "first check in every *mutating*
  adapter"; this one mutates nothing, and Excel's worksheet protection blocks writes
  rather than reads — column widths and row heights are readable on a protected sheet.
  Keeping the consultation made a read-only measurement, the input to a read-only
  diagnostic or export, unavailable on any protected target: a capability restriction
  with no product behind it. **Enforcement moves to the Refresh/write boundary (R4.9)**,
  which is the architectural layer where it belongs; that row is not built yet, so
  nothing regresses in the meantime and no protected target can be written, because
  nothing renders yet either.
- **D5 — `NoActiveWorkbook` is unaffected.** The adapter reads `ActiveWorkbook`
  directly rather than inferring it from a protection query, so the refusal is still
  produced and still typed. A positive test pins this.
- **D6 — The guard parameter was removed rather than kept and ignored.** An
  accepted-but-unused parameter is either dead state (`CA1823`) or a held field that
  reads as "this adapter checks protection" — the exact misconception being removed.
  Two call sites and the test harness were updated instead.

### One brief claim verified rather than trusted

The remediation brief asserted that this change "would require updating
`ProtectionGuardFirstTests` accordingly". **It does not.** That test classifies every
Office source by *data mutation*, and `ExcelPanelGridMeasurement` is already listed
under `NoDataMutationOfficeFiles`; removing a guard call does not change its read-only
classification. This was checked by running the architecture suite (84/84 pass,
unchanged), not by assertion.

## Consequences

- A worksheet with mixed row heights is now measurable, and the live integration test
  asserts the exact per-row values rather than a success flag.
- The "first row's height" ambiguity is removed structurally, not by a heuristic. That
  was a live defect: a 15pt/45pt body reporting 15pt previously produced a grid laid
  out entirely at 15pt.
- A read-only measurement now works on a protected workbook. If a future Office build
  ever refuses those reads, the new live test fails and the policy must be revisited
  **with that host evidence** rather than by assumption.
- The R4.6 blocker is discharged: the style work no longer waits on a panel contract
  that could not represent a real worksheet.
