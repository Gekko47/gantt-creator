# Shape locking by default, with a ribbon unlock

Owner ruling 2026-11-04: a hand-dragged bar staying wrong until an explicit Refresh is
**acceptable**, and the shapes should be **locked by default**, with a ribbon button that
unlocks the selected shape for deliberate manual editing.

That ruling also unblocks proposal **F** in
[auto-refresh-streamlining.md](auto-refresh-streamlining.md) (layout-gated verification), and it
makes F *safe*: once shapes cannot be dragged, the only case F stops repairing is an edit the
product has decided is out of scope.

## The constraint, verified

From Microsoft's `_Worksheet.Protect` / `Worksheet.Protect` reference:

- `Shape.Locked` is inert unless the **worksheet is protected**. The add-in currently refuses a
  protected worksheet outright (ADR-0008 D4: `ProtectContents == true` ⇒ `TargetProtected`), so
  locking without touching protection would be self-defeating.
- `DrawingObjects` **defaults to False**. Shape protection requires
  `Protect(..., DrawingObjects:=true, Contents:=true, UserInterfaceOnly:=true)` explicitly.
- **`UserInterfaceOnly` is not persisted.** After save and reopen the sheet is *fully* protected
  until `Protect` is called again. Code (including this add-in) is blocked in that window.
- `Worksheet.ProtectDrawingObjects` is a read-only flag we can use to detect a shape-protected
  sheet.
- With no `Password` argument, `Unprotect()` succeeds without one.

**The consequence that simplifies everything:** the re-arm does not need a workbook-open hook. The
first refresh after a reopen can unprotect and re-protect with `UserInterfaceOnly:=true` before it
writes anything, because we set no password. Between reopen and that refresh the sheet is fully
protected — which is exactly the state the user wants (cells in unlocked columns still accept
typing; shapes cannot be dragged). A `WorkbookActivate` hook is a nice-to-have, not a prerequisite.

## Decision summary

1. **The add-in owns protection of the Gantt worksheet.** `Protect` with
   `DrawingObjects:=true`, `Contents:=true`, `UserInterfaceOnly:=true`, and **no password**
   (accidental-edit prevention, not security; documented as such).
2. **Ownership is recorded in the workbook** so the guard can distinguish "protected by us" from
   "protected by the user". Two candidate homes, to be settled in the ADR:
   - the VeryHidden config settings table (the sanctioned home for configuration metadata) — but a
     new setting key is a schema change: `GanttSchemaVersion.CurrentSchemaVersion` is **21**, and
     ADR-0029 D6 has **no migration** (the remedy is re-Initialise), so every existing workbook
     would need re-initialising;
   - a workbook custom document property (`CustomDocumentProperties`), outside the sheet schema and
     surviving save, at the cost of a second configuration store outside the VeryHidden sheet.
   *Recommendation: the config setting plus a schema bump*, because the VeryHidden sheet is the
   stated single home for configuration metadata and the integrity checker already repairs a
   missing key. The re-Initialise cost must be accepted explicitly.
3. **The guard authorises our own protection.** `ExcelWorksheetProtectionGuard` returns
   `NotProtected` when the sheet is protected *and* the workbook records that we own it *and*
   `ProtectDrawingObjects` is true; every other protected state keeps returning `SheetProtected`.
   Keeping the existing enum member means no call-site audit and no behaviour change anywhere else,
   at the cost of the outcome meaning "writes are authorised" rather than "the sheet is open" —
   which is the contract the 24 call sites actually rely on. Documented on the enum.
4. **Owned shapes are locked.** `Shape.Locked = msoTrue` is set with the rest of the ownership stamp
   in [`ExcelShapeWriter.ApplyOwnership`](../src/GanttCreator.Office/ExcelShapeWriter.cs), and
   `ListOwned`'s existing walk sets it on shapes created before this change (one read, one
   conditional write per shape, only until they are all locked).
5. **Refresh re-arms protection.** At the start of `Refresh()` — before `RunRefresh`, outside the
   write path — if the sheet is protected and ownership is recorded, unprotect and re-protect with
   the three flags. This is the reopen recovery and it makes every refresh self-healing.
6. **"Unlock shape" acts on the selected shape.** Enabled only when the current `Selection` is a
   `Shape` whose name and tag prove ownership. It unprotects the sheet and clears that shape's
   `Locked`. Per the ruling, the next refresh re-protects and re-locks — with the open question
   below.

## The one open question for review

If "unlock → edit → next refresh re-locks and re-asserts geometry" is literal, the manual edit is
undone within one edit-triggered refresh, which makes the button nearly useless. The cheaper,
more useful reading is: unlocking **exempts that shape from reconciliation** until it is re-locked
or the workbook closes, so the user's edit sticks and the shape stops being engine-owned. That is
an in-memory set of exempt primitive ids honoured by `ShapeReconciler` (no schema change), and it
is the recommendation. Confirm before implementing; the literal reading is a one-line difference in
the reconcile classification.

## Work items

1. ADR-0055 (add-in-owned protection, shape locking, guard semantics, no-password rationale).
2. Config: ownership setting (or document property, per the ADR) + schema bump to 22 if the config
   route is taken; integrity-checker expectation and repair path.
3. `IWorksheetProtectionGuard` / `ExcelWorksheetProtectionGuard`: ownership-aware `Query` and
   `QueryTarget`; new seam for the ownership proof so it is testable without Excel.
4. A small `IWorksheetProtectionPort` (unprotect / protect-with-UI-only / read
   `ProtectContents` + `ProtectDrawingObjects`) so the refresh re-arm and the unlock command share
   one implementation and all three COM calls are replaceable in tests.
5. `ExcelShapeWriter`: `Locked = msoTrue` in the ownership stamp; `ListOwned` back-fill for
   pre-existing shapes; `Shape.Locked` as a virtual seam.
6. Orchestrator: re-arm protection as the first step of `Refresh()`.
7. AddIn: unlock command + ribbon button + enablement via the existing selection tracking; re-lock
   path.
8. Guard/architecture tests: protected-by-us authorises writes; protected-by-user refuses;
   `ProtectDrawingObjects == false` refuses; the re-arm runs once and is idempotent.
9. Contract tests: lock on create, back-fill on an unlocked pre-existing shape, unlock clears only
   the selected owned shape and never a user shape, re-lock restores.
10. Live gate (`verify-office.ps1` plus a scripted Excel run): shapes cannot be dragged or deleted
    while protected; the add-in still creates, moves, deletes, z-orders and writes cells; reopen
    recovers; the unlock button works on a selected shape and is disabled otherwise.
11. Then proposal **F** from the streamlining plan, now that a hand-drag is out of scope.

## Risks to record in the ADR

- A user who unprotects manually can edit shapes; the next refresh re-protects and repairs. Intended.
- Any other add-in or VBA can still write (UI-only protects the interface, not macros).
- Losing `UserInterfaceOnly` on reopen is silent; the recovery is the next refresh, so a workbook
  opened and never refreshed sits fully protected. Accepted: that is the user-visible desired state.
- Back-filling `Locked` on existing shapes is a one-time write per shape; it must not run on every
  refresh once locked.
- If the ownership record is lost (a user copies the sheet, or re-initialises), the guard refuses
  as today and the remedy is a refresh that re-establishes ownership — or Initialise.
