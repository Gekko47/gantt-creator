# ADR-0055 — Lock the chart by default, with add-in-owned protection

## Context

Two owner rulings (2026-11-04) set the direction:

1. A hand-dragged shape staying wrong until an explicit Refresh is **acceptable** — which also
   unblocks proposal F in [`plans/auto-refresh-streamlining.md`](../../plans/auto-refresh-streamlining.md).
2. The chart shapes should be **locked by default**, with a ribbon button that unlocks the
   selected shape for deliberate manual editing.

`Shape.Locked` alone cannot deliver (2). Microsoft's `Worksheet.Protect` / `Shape.Locked`
reference establishes four facts that the design has to obey:

- `Shape.Locked` is **inert unless the worksheet is protected**. The add-in refuses a protected
  worksheet outright (ADR-0008 D4: `ProtectContents == true` ⇒ `TargetProtected`), so locking
  without touching protection would make the two features mutually exclusive.
- `DrawingObjects` **defaults to False**. Shape protection requires
  `Protect(..., DrawingObjects:=true, Contents:=true, UserInterfaceOnly:=true)` explicitly; omit it
  and nothing about the shapes changes.
- **`UserInterfaceOnly` is not persisted.** After save and reopen the sheet is *fully* protected
  until `Protect` is called again, so code — including this add-in — is blocked in that window.
- With no `Password` argument, `Unprotect()` succeeds without one.

## Decision

1. **The add-in owns protection of the visible Gantt worksheet.** It calls `Protect` with
   `DrawingObjects:=true`, `Contents:=true`, `UserInterfaceOnly:=true` and **no password**. The
   feature prevents accidental edits; it is not a security control, and a password we do not store
   could never be re-entered. A user who unprotects by hand is not fought with: the next refresh
   re-protects.
2. **Ownership is recorded in the workbook** as the `ProtectGanttSheet` setting (schema version
   **22**). The key is both the ownership record and the user's switch: `FALSE` means the add-in
   neither protects nor re-protects, and the sheet behaves exactly as it did before.
3. **The guard authorises our own protection.** `ExcelWorksheetProtectionGuard` returns
   `NotProtected` when the sheet is protected **and** the workbook records that we own it **and**
   `ProtectDrawingObjects` is true; every other protected state keeps returning `SheetProtected`.
   The existing enum member is reused deliberately: `NotProtected` is what all 24 call sites
   actually rely on ("writes are authorised"), and adding a member would demand an audit of every
   one of them for no behavioural difference. The enum's documentation states the widened meaning.
4. **Owned shapes are locked.** `Locked = msoTrue` joins the ADR-0019 ownership stamp on create,
   and the existing `ListOwned` walk back-fills it once on shapes created before this change.
5. **Refresh re-arms protection before it writes.** Because `UserInterfaceOnly` does not survive a
   reopen, the first refresh after a reopen unprotects and re-protects with the three flags. This is
   why no workbook-open hook is required: between the reopen and that refresh the sheet is fully
   protected, which is the state the user asked for (unlocked data columns still accept typing;
   shapes cannot be dragged).
6. **Unlock is transient at the reconciler and suspended at the scheduler.** "Unlock shape" acts
   only on a selected, provably-owned shape: it unprotects the sheet and clears that shape's
   `Locked`. Automatic refresh is **suspended** for as long as anything is unlocked, so a manual
   edit is not lost to an unrelated cell edit; an explicit Refresh, or a "Lock shapes" command,
   re-protects, re-locks and re-asserts the shape's geometry — reverting the manual edit, exactly
   as ruled.

## Why not an exemption set

The alternative — exempting an unlocked shape from reconciliation until it is re-locked — was
considered and rejected. It makes ownership partial: "every owned shape equals the scene" stops
being an invariant, `ShapeReconciler` grows a special case, and the next manual Refresh must either
honour the exemption (the chart is permanently inconsistent with the table) or discard it (the
exemption was a lie). Suspending the *scheduler* instead keeps the invariant total and costs one
flag.

## Consequences

- A workbook written before schema 22 reports a version/hash mismatch and needs Initialise
  (ADR-0029 D6, no migration). Accepted, because the VeryHidden config sheet is the single home for
  configuration metadata and the integrity checker already repairs a missing key.
- The add-in now writes protection state to the user's worksheet. That is a new class of mutation,
  and it is confined to one port (`IWorksheetProtectionPort`) so it can be tested and audited.
- A user who protects the sheet themselves with "protect shapes" but without our marker still
  gets `SheetProtected`: we only authorise protection we can prove we own.
- Another add-in or VBA can still write (UI-only protects the interface, not macros).
- Losing `UserInterfaceOnly` on reopen is silent; the recovery is the next refresh, so a workbook
  opened and never refreshed stays fully protected. Accepted: that is the desired user-visible
  state.
- The `Shape.Locked` back-fill is a one-time write per shape and must not repeat once locked.

## Validation

Contract tests: the guard authorises ours and refuses the user's; `ProtectDrawingObjects == false`
refuses; the re-arm is idempotent and runs once per refresh; `Locked` is set on create; the back-fill
touches only unlocked shapes; unlock refuses a user shape and clears only the selected owned one;
the scheduler does not schedule while unlocked.

Live gate (required — the feature is almost entirely host behaviour): shapes cannot be selected,
dragged or deleted while protected; the add-in can still create, move, resize, delete, z-order and
write cells; reopen then refresh recovers writing; the ribbon button enables only on an owned shape.
