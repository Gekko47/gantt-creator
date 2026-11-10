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

   **The visible table is exempt (owner ruling 2026-11-09, clause 7).** `Contents:=true` protects
   cells whose `Locked` is `true`, and Excel's default is `true` on every cell — so the original
   form of this clause made the whole schedule read-only and blocked copy, paste and formatting.
   Every protect now unlocks the table's data cells first and grants the formatting permissions.
   The protection's remaining job is the **chart**: `DrawingObjects` plus the `Shape.Locked` stamp.

   **The native outline symbols stay usable.** `Protect` greys out the +/- group buttons on a
   protected sheet unless `Worksheet.EnableOutlining` is `true` — Microsoft documents it as "True if
   outlining symbols are enabled when user-interface-only protection is turned on" — and it "is not
   saved with the worksheet or session". The port therefore sets it immediately before every
   `Protect`, skipping the set when the sheet is already protected (the host refuses the property
   there, and a second Lock shapes click must stay a no-op). Without this the lock removed the
   collapse/expand control that the outline pass exists to build (ADR-0026 D5), which is a
   user-facing regression rather than a lock.
2. **Ownership is recorded in the workbook** as the `ProtectGanttSheet` setting (schema version
   **22**). The key is both the ownership record and the user's switch: `FALSE` means the add-in
   neither protects nor re-protects, and the sheet behaves exactly as it did before.
3. **The guard authorises our own protection.** `ExcelWorksheetProtectionGuard` returns
   `NotProtected` when the sheet is protected **and** the workbook records that we own it; every
   other protected state keeps returning `SheetProtected`. The existing enum member is reused
   deliberately: `NotProtected` is what all 24 call sites actually rely on ("writes are
   authorised"), and adding a member would demand an audit of every one of them for no behavioural
   difference. The enum's documentation states the widened meaning.

   Ownership is claimed only for protection **this session applied**. `ProtectionArmingState`
   carries that fact, and the refresh sets the ownership flag to `protects && armed`. A sheet the
   *user* protected is therefore still theirs: arming it either succeeds (the add-in then owns it)
   or is refused — typically by a password we do not have — and in the second case the existing,
   actionable `TargetProtected` refusal is kept rather than replaced by a host error halfway
   through a write.

   `ProtectDrawingObjects` is deliberately **not** part of the authorisation test. It describes the
   lock's behaviour, not whose protection this is, and requiring it would refuse writes to a sheet
   we own whose shape flag had been lost — exactly the state a refresh must repair. The ownership
   fact reaches the guard through `IProtectionOwnershipSource`, set from the configuration the
   refresh has already read, so the decision costs **no host call**; a guard that read the
   configuration itself would add a second config read to every command and every refresh.
4. **Owned shapes are locked.** `Locked = true` joins the ADR-0019 ownership stamp on create, and
   `IShapeWritePort.LockOwnedShapes()` is the one-time back-fill for shapes created before this
   change. It is a separate operation rather than something `ListOwned` does, because `ListOwned`
   runs on every refresh and reading `Locked` once per shape per pass would cost more than the write
   it avoids. `Shape.Locked` is a `Boolean` on this PIA, not an `MsoTriState`.
5. **Every write phase runs inside a release window (owner ruling, 2026-11-04).** The live measurement
   below shows a protected sheet refuses shape creation, shape movement and table-row insertion, so
   the feature cannot simply "protect and let code through". The shape is:

   - **Arm once per session** (`WorksheetProtectionArming`): read the state → release if protected →
     lock the owned shapes → record the arming. It ends with the sheet **released**, because
     protecting here would stop the first write the arming exists to enable.
   - **Write inside the window** (`WorksheetProtectionWindow`): `Open()` releases once, the outermost
     scope's `Dispose()` re-applies `Protect(DrawingObjects, Contents, UserInterfaceOnly)` once, and
     nesting is by depth so a command that triggers a render still releases exactly once. The scope is
     returned even when nothing is armed, because arming can happen inside it; when nothing is armed
     at either end the cost is a bool read and no host call.
   - **Every refresh and every structural command opens one.** Refresh (manual and scheduled), Add
     activity, Add child, Promote, Demote, Move up, Move down, Remove, and outline collapse. Cell-only
     commands do not need it: cells are the one write the lock still allows.

   The window costs **two host calls per write phase**, which the owner accepted in place of the
   original "no extra step" wording. The sheet is unlocked for the few hundred milliseconds a refresh
   takes and locked the rest of the time, which is the state the ruling asked for.
6. **The two toggles replaced the per-shape unlock (owner ruling 2026-11-09).** The original form of
   this clause released exactly one selected, provably-owned shape and remembered which, as a
   transient exemption, behind "Unlock shape" and "Lock shapes" buttons. The ruling replaced that
   pair with **one toggle for the CHART** and **one for the SHEET CELLS** (clause 7), each visibly
   reflecting the state it is in. The one-shot commands, the selection reader that authorised them
   and the per-shape unlock set that suspended the scheduler were retired with it, so there is
   exactly one statement of "what is locked" rather than two that could disagree.

   The properties the original clause existed for are kept. Automatic refresh is **suspended** for
   as long as the chart is unlocked, so a manual drag is not lost to an unrelated cell edit, and an
   explicit Refresh re-protects, re-locks and re-asserts the shape's geometry — reverting the manual
   edit, exactly as ruled. Suspension is **chart-only**: unlocking the sheet's cells is the user
   choosing to type, and the automatic refresh is what consumes that edit, so the cell toggle never
   suppresses it. The decision is one read of `ProtectionPreferences.ChartLocked`, and enabling the
   two buttons costs no host call — the state comes from the same object the protect path reads.
7. **The visible table stays editable while the chart is locked (owner ruling, 2026-11-09).** The
   requirement, in the owner's words: "the visible table area should be fully unlocked at all times
   for the user to copy paste change formatting etc". Three facts make that two changes rather than
   one. `Locked` is a **per-cell format** property and Excel defaults it to `true`, so `Contents`
   protects every cell until it is cleared. **Formatting is a sheet permission, not a per-cell one**,
   so unlocking the cells alone still leaves Format Cells and Paste Special refused. And `Locked`
   cannot be written on a protected sheet, so the unlock has to happen inside the release the protect
   path already performs.

   The port therefore **unlocks the table's data cells and then grants `AllowFormattingCells`,
   `AllowFormattingColumns` and `AllowFormattingRows`** on every protect. The data body is unlocked
   rather than the whole table on purpose: unlocking the header would let a user rename a schema
   column, which silently breaks the name-keyed column map every adapter builds. A refused unlock is
   recorded (`cellsUnlocked=refused`) and does not abandon the protect — a sheet protected with a
   password we do not have is still worth protecting.

   **What still refuses, and why that is the point.** Shapes still cannot be created or moved on the
   protected sheet (measured below), so the chart cannot be damaged by a stray drag; the table row
   insert is still refused, which is why the structural commands keep their release window. What
   changed is only who may edit the *data*: the user, freely, at any time.

## Measured host behaviour (2026-11-04, live Excel 16.0.20527.20112 x64)

`ProtectionArmingIntegrationTests` protects a real sheet with `DrawingObjects`, `Contents` and
`UserInterfaceOnly` and then performs the writes a refresh makes. Everything below is observed, not
inferred:

| operation | sheet unprotected | sheet protected (UI-only) |
|---|---|---|
| write a cell (`Range.Value2`) | ok | **ok** |
| `Shapes.AddShape` (create a bar) | ok | **COMException, "The specified value is out of range"** |
| `Shape.Left = …` (move a shape) | ok | **ArgumentException, same message** |
| insert a table row (`ListRows.Add`) | ok | **refused: `RowWriteRefused`** |
| `Shape.Locked = true` on an owned shape | ok | **ok** |
| `Unprotect()` then create + insert | — | **ok** |
| `Protect()` again, then write a cell | — | **ok** |

So `UserInterfaceOnly` exempts **cells only** in this host: the drawing-object lock blocks every
shape write and the contents lock blocks the row insert. The consequence is that a refresh cannot run
with the sheet protected, and neither can Add activity, Add child, Promote, Demote, Move or Delete —
which is why clause 5 makes every one of those run inside the release window, and why the window was
accepted despite costing two host calls.

Two further facts from the same run: a workbook protected by an earlier session refuses a command
that runs before the first refresh (`TargetProtected`), because ownership is only established by a
refresh; and the "out of range" message names no cause, so the log line matters more than the text.

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
- The arming memory is **process state, not workbook state**: a second Excel instance, or the same
  workbook reached by another add-in process, re-arms on its first refresh. That is the honest
  reading of "we applied this protection", and it is why the memory is invalidated rather than
  trusted across a workbook change.

## Validation

Contract tests: the guard authorises ours and refuses the user's; `ProtectDrawingObjects == false`
refuses; a refresh locks before it protects and arms once per session, not once per pass; an
already-protected sheet is released before it is protected again; a refused protection write keeps
the chart and leaves the arming unmarked; an opt-out workbook is never protected and an opt-out after
arming releases it; `Locked` is set on create; the back-fill touches only tagged shapes; each toggle
**sets** the state it is given instead of flipping it, re-applies protection, records the arming and
leaves the other toggle alone; the scheduler queues while only the cells are unlocked and stops while
the chart is unlocked. Live: a protect leaves `EnableOutlining` true (and it is false before the
add-in sets it), so the +/- group buttons survive the lock.

Live gate (required — the feature is almost entirely host behaviour): shapes cannot be selected,
dragged or deleted while protected; the add-in can still create, move, resize, delete, z-order and
write cells; reopen then refresh recovers writing; the chart toggle decides whether a shape drag
lands; the cell toggle decides whether a cell outside the table is writable; the table's own cells
stay editable in both states.
