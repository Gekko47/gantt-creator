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
6. **Unlock is transient and suspends the scheduler, never the reconciler.** "Unlock shape" acts
   only on a selected, provably-owned shape: it unprotects the sheet and clears that shape's
   `Locked`. Automatic refresh is **suspended** for as long as anything is unlocked, so a manual
   edit is not lost to an unrelated cell edit; an explicit Refresh, or a "Lock shapes" command,
   re-protects, re-locks and re-asserts the shape's geometry — reverting the manual edit, exactly
   as ruled. The two controls are answered from the unlock set and from one `Selection` read rather
   than from the ribbon snapshot, so enabling the button costs no host call during a refresh.

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
arming releases it; `Locked` is set on create; the back-fill touches only tagged shapes; unlock
refuses a user shape and clears only the selected owned one; the scheduler does not schedule while
unlocked. Live: a protect leaves `EnableOutlining` true (and it is false before the add-in sets
it), so the +/- group buttons survive the lock.

Live gate (required — the feature is almost entirely host behaviour): shapes cannot be selected,
dragged or deleted while protected; the add-in can still create, move, resize, delete, z-order and
write cells; reopen then refresh recovers writing; the ribbon button enables only on an owned shape.
