# Remediation implementation plan — Commit C (live per-row measurement)

Follows [`REMEDIATION-PLAN-B.md`](REMEDIATION-PLAN-B.md) (Commit B, `11df4e1`), which
made the Core contract able to express a mixed-height body. **This plan is written
before any code change**, per the traceability requirement.

## Problem

`ExcelPanelGridMeasurement` still cannot supply what Commit B defined. Two distinct
gaps, plus one policy question the brief raised.

1. **It refuses a mixed-height body.** `ReadRowHeight` reads
   `DataBodyRange.RowHeight`, then `ConfirmUniformRowHeight` walks every body row and
   demands each one equals the aggregate. A body whose rows differ is reported as
   `InvalidMeasurement`. The Core contract can now express such a body; the adapter
   cannot yet measure one, so the capability Commit B bought is unusable end to end.
2. **The header height is the body height.** Commit B added
   `PanelCellGrid.HeaderHeightPt` and `ExcelPanelGridMeasurement` fills it with the
   same confirmed body height. §4 says the header "follows live header-cell bounds",
   and the header row is its own Excel row with its own height.
3. **A read-only adapter refuses a protected worksheet.** `Measure` consults
   `IWorksheetProtectionGuard` and returns `TargetProtected`, even though it writes
   nothing. This is the brief's `P2-1`.

## Decisions (recorded before implementation)

- **D-C1 — Measure each body row individually; the aggregate is not consulted at all.**
  The per-row confirmation loop is deleted, not relaxed. The `Range.RowHeight`
  reference documents that a mixed range "might return the height of the first row or
  might return Null", so the aggregate is *unreliable in both directions* for a
  non-uniform body: `DBNull` is caught today, but the "first row's height" case is a
  plain number that passes every check while describing a body most of which is taller.
  Reading rows individually removes the ambiguity instead of detecting it.
- **D-C2 — The header row is measured separately** from `ListObject.HeaderRowRange`,
  on its own seam. This is what §4 asks for and what Commit B's field is for.
- **D-C3 — A row the host cannot report is refused, not defaulted.** One absent row
  height must not become one guessed height, because the panel's row positions are
  cumulative: a single wrong height displaces every row below it. `InvalidMeasurement`
  stays the refusal.
- **D-C4 — The read-only protection consultation is removed from this adapter.**
  Excel's worksheet protection blocks *writes*; reading column widths and row heights
  is permitted. Keeping the guard meant a read-only diagnostic and a read-only export
  measurement could not run on a protected sheet — a capability restriction with no
  product behind it. Enforcement belongs at the write/Refresh boundary, which is R4.9;
  that row is not built yet, so removing it here opens no window in which a protected
  target is written.
  - **`NoActiveWorkbook` is unaffected**: the adapter already checks
    `application.ActiveWorkbook is null` itself, so that refusal reason is still
    reachable and still typed.
  - **The remediation brief's claim that this needs a `ProtectionGuardFirstTests`
    update is wrong.** That test classifies every Office source by *data mutation*, and
    `ExcelPanelGridMeasurement` is already listed under `NoDataMutationOfficeFiles`;
    removing a guard call does not change its read-only classification. The claim is
    verified by running the architecture suite rather than by assertion.
- **D-C5 — The live integration test is updated, not deleted.** It currently proves a
  uniform body measures and then that a mixed body is *refused*. The second half
  inverts: it must now prove a mixed body measures with three distinct heights. It
  remains tagged `OfficeIntegration` and **is not run** here.

## Steps

1. Write this plan. **Done.**
2. `ExcelPanelGridMeasurement`: replace the aggregate read with a per-row walk over
   `GetBodyRowAt`/`GetBodyRowCount`; add a `HeaderRowRange` seam and read the header
   height from it; assemble `RowHeightsPt` in body order plus `HeaderHeightPt`.
3. Remove the protection consultation and its comment, keeping the `NoActiveWorkbook`
   path intact via the adapter's own check.
4. Update `ExcelPanelGridMeasurementTests` (contract): a uniform body still yields one
   height per row; a **mixed** body yields the exact distinct heights in order; the
   header height is read from the header row and is independent of the body; a row
   reporting `DBNull` refuses; a body with no rows refuses; **measurement succeeds on a
   protected target** (the new positive test for D-C4).
5. Update `PanelGridMeasurementIntegrationTests` (live, not run here) so the mixed case
   asserts success with three heights instead of a refusal.
6. Run the architecture suite to confirm D-C4's claim about
   `ProtectionGuardFirstTests` rather than asserting it.
7. Docs: ADR-0024, entity guide §3/§4 if the reading needs restating, REPO-MAP, STATUS.

## Non-vacuity

The mixed-height contract test asserts the **exact** ordered list `[12, 24, 15]`. A
tolerance or a "succeeded" assertion would pass against the old uniform-only code
path if the heights happened to coincide, so the values themselves are the assertion.
The protected-target test is a positive test: it drives a target reporting protected
and asserts a grid comes back, which fails while the guard is still in place.

## Explicitly not in this commit

No Core change — the contract is already correct, and changing it again would reopen
Commit B. No `IShapeWritePort` or reconciliation change. No Refresh-command protection
policy, which is R4.9's and is named here as the place the rule now belongs.
