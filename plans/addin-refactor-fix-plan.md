# Add-in refactor fix plan — Code Mode implementation

Status: proposed. Produced by a read-only review. No product code was changed.
Date: 2026-11-26.
Source review: ranked findings on `src/GanttCreator.AddIn/` (55 hand-written files, 20,531 physical lines).

This plan is a work order for Code Mode. Every slice is behaviour-preserving. Every slice is
independently revertable. Every slice names the tests that prove it.

---

## 1. Ground rules for Code Mode

1. **One slice per commit.** Conventional Commit prefix `refactor(addin):` unless the slice only
   adds a log record, which is `fix(addin):`.
2. **No message text changes.** Every user-facing sentence stays byte-identical. Tests assert them.
3. **No public or internal signature changes**, except the new members named in this plan.
4. **No parameter reordering** on `RibbonState`. Adding argument *names* is the fix. Reordering is a
   different change with a wider blast radius and is out of scope.
5. **Tests in the same commit** as the change, per `AGENTS.md`.
6. **Do not delete or weaken a test** to get green. If a test fails, stop and report.
7. **Do not add a new production dependency.** Nothing here needs one.
8. **Stop and ask** if a slice needs a Core/Office change, a `RibbonState` shape change, or a new
   interface that crosses project boundaries.

### Gate commands

```text
pwsh -NoProfile -File scripts/verify-quick.ps1          # every commit
pwsh -NoProfile -File scripts/verify.ps1                # branch-final, clean tree
pwsh -NoProfile -File scripts/check-status.ps1          # STATUS accuracy
```

Targeted loop for every slice:

```text
dotnet build src/GanttCreator.AddIn/GanttCreator.AddIn.csproj
dotnet test  tests/GanttCreator.AddIn.Tests/GanttCreator.AddIn.Tests.csproj
```

`Directory.Build.props` sets `TreatWarningsAsErrors`, `AnalysisMode=All`,
`EnforceCodeStyleInBuild`, and `AnalysisLevel=latest`. The build is the analyzer gate. No slice may
end with a new suppression.

### Baseline before slice A

Run the targeted test project once and record the count. Every later slice compares against it.
Do not start if the baseline is red; report the red test as unrelated and out of scope.

---

## 2. Evidence this plan rests on

| Fact | Measurement |
|---|---|
| File sizes | `RibbonStateService.cs` 4,416 lines; `ThemeSettingsForm.cs` 3,153; `GanttRibbon.cs` 2,815; `RibbonState.cs` 398 |
| `RibbonState` shape | 34 parameters, [`RibbonState.cs:166`](../src/GanttCreator.AddIn/RibbonState.cs:166) |
| Unnamed snapshot builds | [`RibbonState.cs:212`](../src/GanttCreator.AddIn/RibbonState.cs:212) (16 args), [`RibbonStateService.cs:3500`](../src/GanttCreator.AddIn/RibbonStateService.cs:3500) (28 positional args) |
| Duplicated presenter helper | 15 files, **7 distinct bodies** by hash |
| Duplicated selection gate | 4 files |
| Duplicated dependency ladder | 2 files, 10 arms each |
| Pragma balance | every file pairs disable/restore except [`GanttRefreshScheduler.cs:1`](../src/GanttCreator.AddIn/GanttRefreshScheduler.cs:1) |
| Silent failure paths | [`RibbonStateService.cs:1479`](../src/GanttCreator.AddIn/RibbonStateService.cs:1479), [`RibbonStateService.cs:3538`](../src/GanttCreator.AddIn/RibbonStateService.cs:3538), [`AddInHost.cs:183`](../src/GanttCreator.AddIn/AddInHost.cs:183) |

The duplicated presenter helper lives in: `AddRowCommand.cs`, `CollapseGroupsCommand.cs`,
`CopyEditableCommand.cs`, `DemoteToChildCommand.cs`, `MoveDownCommand.cs`, `MoveUpCommand.cs`,
`PromoteToParentCommand.cs`, `RefreshSheetCommand.cs`, `RemoveCommand.cs`, `RepairConfigCommand.cs`,
`SendToPowerPointCommand.cs`, `SetDateLabelCommand.cs`, `SetDateLabelPlacementCommand.cs`,
`SetDescriptionLabelCommand.cs`, `SetStyleKeyCommand.cs`.

Body groups: `75-0D-8E` × 8 (CollapseGroups, DemoteToChild, PromoteToParent, Remove, SetDateLabel,
SetDateLabelPlacement, SetDescriptionLabel, SetStyleKey); `D4-9F-EB` × 3 (AddRow, MoveUp, MoveDown);
singletons in CopyEditable, RefreshSheet, RepairConfig, SendToPowerPoint.

---

## 3. Slices, in implementation order

Order is by risk and by dependency: the cheap mechanical work first, the wide refactors last.

| Slice | Finding | Size | Risk | Files touched |
|---|---|---|---|---|
| A | 1 — high-parameter snapshot | S | Very low | 2 |
| B | 3 — duplicated presenter policy | M | Low | 16 |
| C | 4, 5, 6 — silent failure paths | S | Low | 3 |
| D | 2 — god class | L | Medium | 1 → 6 |
| E | 7 — 328-line OK handler | M | Medium | 1 → 3 |
| F | 8 — file-wide pragma | S | Low (build-gated) | 1 |
| G | 9 — duplicated refresh postlude | S | Low | 1 |
| H | 10 — ribbon getter boilerplate | M | Low | 2 |

Slices F, G and H are independent of A–E. They can ship in any order after the baseline.

---

## Slice A — Name every argument in the two snapshot builds

**Goal.** Remove the silent-transposition hazard in the 34-parameter record. No type changes.

**Why.** `Initial` passes 16 unnamed values, including four `string.Empty` in a row and three `null`
in a row. `CaptureSnapshot` passes 28 unnamed values. Two swapped `bool` arguments compile. The
whole ribbon surface reads from this record, so a swap is a silent, whole-surface misbehaviour.

### A1 — `RibbonState.Initial`

File: [`RibbonState.cs`](../src/GanttCreator.AddIn/RibbonState.cs:212)
Operation: add `Name:` to every one of the 16 arguments. Values, order and line wrapping stay the
same otherwise.

Acceptance:
- `git diff -w --word-diff` shows only `Name:` insertions.
- `RibbonValueContractTests` and `RibbonStateServiceTests` green.

### A2 — `CaptureSnapshot`

File: [`RibbonStateService.cs`](../src/GanttCreator.AddIn/RibbonStateService.cs:3500)
Operation: add `Name:` to the 28 unnamed arguments. The 8 trailing arguments are already named; leave
them.

Acceptance:
- Diff is names only.
- `RibbonStateServiceTests`, `RibbonDropdownContractTests`, `GanttRibbonTests` green.

### A3 — Deferred: group the capture facts

Do **not** do this in Slice A. If it is wanted later, it is its own change:
introduce `readonly record struct SnapshotFacts(bool? SheetInitialised, bool? ConfigurationPresent,
bool? GanttSheetPresent, bool? CriticalIntervalParentAvailable, bool? ActiveSheetIsGanttSheet,
bool? HasGroups)` and change `CaptureSnapshot`'s signature. That touches the call site at
[`RibbonStateService.cs:2220`](../src/GanttCreator.AddIn/RibbonStateService.cs:2220) and every
partial call (for example `_state, _state.SheetInitialised`). It is behaviour-preserving but larger,
so it needs its own review.

**Rollback.** Revert the file. Nothing else refers to these call sites by shape.

---

## Slice B — One presenter-failure policy

**Goal.** Replace 15 copies of `Present(Action<string>, string)` with one tested helper.

**Why.** The policy "a presenter failure degrades to no dialog" is written 7 different ways today.
Any change is a 7-site edit, and the drift proves the copies will keep diverging.

### B1 — Add the helper

File: [`CommandBoundary.cs`](../src/GanttCreator.AddIn/CommandBoundary.cs:268)
Operation: add one member next to `WriteTechnicalRecord`:

```csharp
/// <summary>
/// Invokes a user-facing presenter, degrading to no presentation when the host
/// refuses it. The single authority for the "a dialog failure must not escape
/// into Excel" policy; commands call this instead of carrying their own copy.
/// </summary>
internal static void SafePresent(Action<string> presenter, string message)
{
    ArgumentNullException.ThrowIfNull(presenter);

    // CA1031: a presenter failure is outside command behaviour; the work has
    // already happened and must not be turned into an escaping exception.
#pragma warning disable CA1031
    try
    {
        presenter(message);
    }
    catch
    {
        // Intentionally empty: the dialog degrades to nothing.
    }
#pragma warning restore CA1031
}
```

Tests, in [`CommandBoundaryTests.cs`](../tests/GanttCreator.AddIn.Tests/CommandBoundaryTests.cs):

- `SafePresent_InvokesPresenterOnceWithTheMessage`
- `SafePresent_SwallowsAPresenterFailure`
- `SafePresent_ThrowsWhenPresenterIsNull`

Acceptance: three tests green; no other file changed in this commit.

### B2 — Replace the 15 copies

Operation, one commit:
- Delete each private `Present` method.
- Replace each call `Present(presenter, message)` with
  `CommandBoundary.SafePresent(presenter, message)`.

Order inside the commit is irrelevant, but run the tests per group to localise a failure:
1. The eight `75-0D-8E` files.
2. The three `D4-9F-EB` files.
3. `CopyEditableCommand.cs`, `RefreshSheetCommand.cs`, `RepairConfigCommand.cs`,
   `SendToPowerPointCommand.cs`.

Watch for behaviour that the private copies did *not* share:
- `ExportPngCommand.cs` and several commands have no `Present` helper; they call the presenter
  directly elsewhere. Leave those call sites alone unless they also catch.
- `RefreshSheetCommand.Present` is called with `RibbonStatusBar.Warn` as well as
  `CommandErrorDialog.Show`. The helper takes the delegate, so both keep working.

Acceptance:
- `grep -rn "private static void Present" src/GanttCreator.AddIn` returns nothing.
- These suites green: `CommandBoundaryTests`, `AddRowCommandTests`, `CollapseGroupsCommandTests`,
  `DemoteToChild`/`PromoteToParent` coverage via `ParentChildCommandTests`, `MoveCommandTests`,
  `RemoveCommandTests`, `RepairConfigCommandTests`, `LabelCommandsTests`, `CopyEditableCommandTests`,
  `SendToPowerPointCommandTests`, `RefreshSheetCommandTests`.

### B3 — Optional: one selection-gate helper

Only if B2 lands clean. Files: `SetDescriptionLabelCommand.cs:43`, `SetStyleKeyCommand.cs:42`,
`SetDateLabelCommand.cs`, `SetDateLabelPlacementCommand.cs`.
Operation: add

```csharp
internal static bool TryRequireSingleBodyRow(ISelectionIdentifier selection, out SelectionFacts facts)
```

to a new `CommandPreconditions.cs`, returning `false` when not a single body row. Each command keeps
its own translator, so the messages stay per-command.

Acceptance: `LabelCommandsTests`, `DateLabelRefusalTranslatorTests`,
`DescriptionLabelRefusalTranslatorTests` green. Do not merge the four translators.

**Rollback.** Revert B2 alone; B1 can stay because it is additive.

---

## Slice C — Give the three silent failures a witness

**Goal.** Keep the safe degradation. Add one technical record so the degradation is diagnosable.

**Rule.** No new dialog. No state change. One new log line per failure.

### C1 — Settings write with no writer

File: [`RibbonStateService.cs:1479`](../src/GanttCreator.AddIn/RibbonStateService.cs:1479)

```csharp
if (writer is null)
{
    // The in-memory state is the whole truth here. Record it so a half-armed
    // session is diagnosable instead of silently not persisting or refreshing.
    _commandBoundary.WriteTechnicalRecord("gantt.settings persist skipped=no-writer");
    return true;
}
```

Test (in `RibbonStateServiceTests`): install a `CommandBoundary` whose log is a capturing fake, call
`SetTimeScale` (or any setter that reaches `TryPersistPlotLayout`) with no catalogue writer injected,
assert exactly one `TechnicalRecord:` line containing `gantt.settings persist skipped=no-writer`, and
assert the returned state still changed.

The capturing `IRollingLog` fake: reuse the one in `CommandBoundaryTests.cs`. If it is file-private,
promote it to a shared test helper in the same test project. That is a test-only change.

### C2 — Snapshot build failure

File: [`RibbonStateService.cs:3538`](../src/GanttCreator.AddIn/RibbonStateService.cs:3538)

```csharp
catch (Exception exception)
{
    CommandBoundary.Instance.WriteTechnicalRecord(
        $"gantt.ribbon snapshot kept-previous type={exception.GetType().FullName}");
    return previous;
}
```

Test: inject `SetLogAvailabilitySource(() => throw new InvalidOperationException("probe"))`, call
`Refresh()`, assert one record named `gantt.ribbon snapshot kept-previous`, and assert the published
snapshot is unchanged. This keeps getter purity: the record is written from `Refresh`, never from a
getter.

### C3 — Add-in arming failure

File: [`AddInHost.cs:183`](../src/GanttCreator.AddIn/AddInHost.cs:183)

```csharp
catch (Exception exception)
{
    // Armed nothing: every gated command will stay disabled. Record it once.
    CommandBoundary.Instance.WriteTechnicalRecord(
        $"gantt.addin arm failed type={exception.GetType().FullName}");
}
```

Test (in `AddInHostTests`): construct `AddInHost` with an application-adapter source that throws,
attach a capturing log to `CommandBoundary.Instance`, call `AutoOpen`, assert it does not throw and
assert one record named `gantt.addin arm failed`. Reset the singletons afterwards using the same
pattern as `RibbonTestCollectionGuardTests`.

If the throw cannot be injected without changing production code, **do not fake the test**. Report
and leave C3 as a code-only change with its reason stated in the commit message.

**Acceptance for the slice:** three positive tests, no new dialog, no message text changed, build
with analyzers green.

**Rollback.** Revert per finding; C1, C2 and C3 are independent.

---

## Slice D — Split `RibbonStateService` (do in this order)

**Goal.** Cut the 4,416-line class along responsibilities that already have their own tests.

**Method.** For each step: create the new type, move the members, leave a one-line delegate on
`RibbonStateService` with the **same** signature, run tests, commit. Callers must never be edited,
because caller edits would hide a changed behaviour.

### D1 — `RibbonDropdownProjection` (lowest risk, pure reads)

Move (from `RibbonStateService.cs`):
- `GetDescriptionLabelCount` 2753, `GetDescriptionLabelItemID` 3011, `GetDescriptionLabelItemLabel` 3030,
  `GetDescriptionLabelSelectedIndex` 3047
- `GetCustomStyleCount` 3133, `GetCustomStyleItemID` 3143, `GetCustomStyleItemLabel` 3157,
  `GetCustomStyleSelectedIndex` 3169
- `GetDateLabelCount` 3213, `GetDateLabelItemID` 3228, `GetDateLabelItemLabel` 3265,
  `GetDateLabelSelectedIndex` 3284
- `LabelPositionDisplayText` 3076, `DateLabelDisplayText` 3329, `Constants` 3117

Shape: `internal static class RibbonDropdownProjection` with methods taking
`RibbonState state, int index` and returning the same values. No ports, no Excel, no logging.

Keeps on the service: `SetDescriptionLabelPosition` 3102, `SetCustomStyleKey` 3195,
`SetDateLabelPosition` 3355 — they run commands and call `Refresh()`.

Tests: `RibbonDropdownContractTests`, `RibbonValueContractTests`, `CustomStyleRibbonTests`,
`DateLabelRibbonTests`, `DescriptionLabelRibbonTests`, `ColumnVisibilityRibbonTests`.
Expected shrink: about 300 lines.

### D2 — `EntityAddMenuModel`

Move: `FamilyOf` 2441, `DefaultTypeFor` 2469, `PrimaryTypeFor` 2490, `RememberAddedType` 2505,
`PrimaryControlIdFor` 2519, `FamilyOfPrimary` 2530, `AddEntityPrimaryLabel` 2545,
`AddEntityPrimaryGlyph` 2568, `AddEntityPrimaryImage` 2611, and the `AddEntityFamily` enum.
Shape: static class, `RibbonState` passed in for the two label/image selectors.
Tests: `AddRowRibbonTests`, `GanttRibbonTests`, `RibbonValueContractTests`.
Expected shrink: about 350 lines.

### D3 — `PlotRangeSettings`

Move: `SetPlotStartAuto` 488, `SetPlotFinishAuto` 590, `SetPlotStartDate` 767,
`SetPlotFinishDate` 1007, `TryAcceptExplicitStart` 837, `TryAcceptExplicitFinish` 876,
`TryPersistPlotRange` 1413, `LoadPlotSettingsFromWorkbook` 1820, `RefreshEffectivePlotDates` 1994,
`NormaliseStoredPlotDate` 2090, `ReadRangePaddingDays` 2106, `ReadPeriodLabelFormat` 1919,
`ReadTimeScale` 1932, `ReadMargin` 1943, `ReadMarginCm` 1956, `ReadPreset` 1968,
`CanonicalPeriodLabelFormat` 1904, `InvalidatePlotDateRow` 1376.

Do **not** move `ShowPlotDatePicker` 674 or `PickPlotDate` 696: they own the dialog seam and an
active virtual for tests. The setters stay on the service and delegate.

Shape note: the new type needs the catalogue reader/writer, the warning presenter, the command
boundary (for `WriteTechnicalRecord`) and the refresh scheduling callback. Pass them as constructor
delegates, not singleton lookups, so `RibbonStateServiceTests` keeps working.

Tests: `PlotDatePickerTests`, `PlotRangeWarningTests`, `RibbonStateServiceTests`.
Expected shrink: about 600 lines.

### D4 — `RibbonSnapshotFactory` (highest risk, do last)

Move: `CaptureSnapshot` 3449, `CaptureRowActionContext` 3554, `RowActionContext` 3591,
`ResolveMode` 3607, `HasEligibleParent` 3629, `CaptureDescriptionLabelContext` 3689,
`UnavailableDescriptionLabelContext` 3770, `CaptureCustomStyleContext` 3782,
`UnavailableCustomStyleContext` 3875, `CaptureDateLabelContext` 3894,
`UnavailableDateLabelContext` 3958, `CaptureDateLabelPlacementContext` 3974,
`UnavailableDateLabelPlacementContext` 4039.

Inputs it needs: application adapter, log-availability source, selection identifier, table rows,
and the four `RibbonState`-projected contexts. Pass a small `SnapshotSources` record.

Keep the sticky policy and the catch exactly. If C2 has landed, keep the record.

Tests: `RibbonStateServiceTests`, `RibbonDropdownContractTests`, `RibbonTestCollectionGuardTests`,
`GanttRibbonTests`.
Expected shrink: about 600 lines.

### D5 — Deferred: `RibbonScheduling`

Activation, subscriptions, `Refresh`, `Teardown` (2031-2159 and 4046-4415) are event-lifetime code.
Do not move them in this plan. They are the highest regression risk for the least line saving, and
`AddInHostTests` plus the live Office gate are the only real proof. Revisit as its own work item.

**Stop and ask during D** if:
- a moved method needs a type from Core or Office that the AddIn does not already reference;
- a test must change its assertions (that means behaviour changed);
- line counts shrink but a caller needs an edit.

**Rollback.** Each of D1–D4 is one commit and reverts alone.

---

## Slice E — Take the OK handler apart

**Goal.** `OkButton_Click` is 328 lines ([`ThemeSettingsForm.cs:2748`](../src/GanttCreator.AddIn/ThemeSettingsForm.cs:2748)).
Split it into three named steps without touching ordering.

### E1 — Extract the request builder

Move the diffing block (2878–3007) into
`ThemeEditRequestBuilder.Build(styleEdits, metricEdits, customStyleEdits, chromeColourEdits)`.
Keep `ThemeEditRequest.TryValidate` where it is; the builder only assembles.

Signature must take plain data: `IReadOnlyDictionary<string, CustomStyleState>`,
`IReadOnlyList<CustomStyleDef> originalCustomDefs`, `IReadOnlyDictionary<string, string> pendingRenames`,
`IReadOnlyDictionary<string, EntityStyleColours>`.
The rename-first, delete-second, create/update-third order is the contract. Preserve it exactly.

Tests: add `ThemeEditRequestBuilderTests` cases for
`Rename_IsNotAlsoReportedAsDelete`, `UnchangedStyle_ProducesNoEdit`,
`ChangedLabelPositions_ProduceAnUpdate`.

### E2 — Extract colour validation

Move the four validation loops (2754–2820) into `ThemeColourValidation.TryValidate(...)` returning a
typed failure: `(string Message, int TabIndex, Func<Control?> Focus)`. The strings stay identical.

Tests: `ThemeSettingsFormTests` cases for a bad entity colour, a bad custom colour and a bad chrome
colour. WinForms control creation stays on the existing STA test path.

### E3 — Extract the write/warn/close sequence

Move 3030–3074 into `ApplyWriteThenWarn(ThemeEditRequest request)`. Preserve:
- write happens before the warning;
- `alreadyShown` skips the second write;
- a refused write sets `DialogResult = OK` so the caller reports it;
- a fresh warning keeps the dialog open.

Tests: `ThemeSettingsFormContrastWarningTests` — the suite already pins the two-OK behaviour.

Acceptance for the slice: the three test files green, messages byte-identical, no change to
`DialogResult` outcomes.

**Rollback.** One commit per extraction.

---

## Slice F — Remove the file-wide pragma from the scheduler

File: [`GanttRefreshScheduler.cs:1`](../src/GanttCreator.AddIn/GanttRefreshScheduler.cs:1)
Operation:
1. Read the timer construction (the file is 452 lines; the expression is around 260–452).
2. Remove the ambiguity. Preferred shape: assign the callback to a typed local before the
   constructor call.

```csharp
TimerCallback callback = OnTimerElapsed;
_timer = new Timer(callback, state: null, Timeout.Infinite, Timeout.Infinite);
```

or, if the real expression differs, add the parentheses the analyzer asks for.
3. Delete line 1 only after the build is clean.

Acceptance:
- `dotnet build src/GanttCreator.AddIn/GanttCreator.AddIn.csproj` green with
  `TreatWarningsAsErrors` (the build *is* the analyzer proof).
- `GanttRefreshSchedulerTests` green.
- `grep -c "#pragma warning disable" GanttRefreshScheduler.cs` equals the number of restores.

**Do not** replace the pragma with `NoWarn` in the csproj. That moves the hole, not closes it.

---

## Slice G — One refresh postlude

Files: [`RefreshSheetCommand.cs:182`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs:182) (manual)
and [`RefreshSheetCommand.cs:252`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs:252) (silent).

Operation: extract

```csharp
private static void FinishRefresh(
    GanttRefreshOutcome outcome,
    bool automatic,
    Action<string> failurePresenter,
    bool invalidateSelectionGates)
```

covering: validation issues, warnings/mark-current/status-bar on success, the failure presentation,
the eight `InvalidateControl` calls when the flag is set, the tab restore, and the three trace calls.

Call it from both paths with `invalidateSelectionGates: false` (manual) and `true` (silent).
Keep the arguments `automatic: false/true` so the trace text is unchanged.

Acceptance:
- `RefreshSheetCommandTests`, `GanttRefreshSchedulerTests`, `ExportPngCommandTests`,
  `AutoRefreshRibbonTests` green.
- The trace strings stay `source=manual` / `source=automatic`.
- Diff removes one copy and leaves one.

**Stop and ask** if the manual path's behaviour turns out to depend on `RibbonStateService.Refresh()`
running before the tab restore. The current order is: command → `NotifyRibbonStateChanged` → tab
restore. The extraction must not move the tab restore inside `FinishRefresh` if that would reorder it.
If it would, keep the tab restore at the call site and share only the rest.

---

## Slice H — One helper for ribbon getters

File: [`GanttRibbon.cs`](../src/GanttCreator.AddIn/GanttRibbon.cs:1032).

Reality check before starting: the *command* callbacks already share `OnAddRowClick`
([`GanttRibbon.cs:879`](../src/GanttCreator.AddIn/GanttRibbon.cs:879)). The duplication that remains is
in the **getters**, where each internal overload carries its own try/catch because the failure value
differs.

### H1 — Add `RibbonGetter`

```csharp
internal static class RibbonGetter
{
    internal static bool Bool(RibbonStateService state, Func<RibbonStateService, bool> read, bool onFailure)
        => Read(state, read, onFailure);

    internal static string String(RibbonStateService state, Func<RibbonStateService, string> read, string onFailure)
        => Read(state, read, onFailure);

    internal static int Int(RibbonStateService state, Func<RibbonStateService, int> read, int onFailure)
        => Read(state, read, onFailure);

    internal static object? Object(RibbonStateService state, Func<RibbonStateService, object?> read)
        => Read<object?>(state, read, null);

    private static T Read<T>(RibbonStateService state, Func<RibbonStateService, T> read, T onFailure)
    {
        ArgumentNullException.ThrowIfNull(state);
        // CA1031: a probe failure must fail to the getter's own safe value, never
        // into Excel's getter dispatch.
#pragma warning disable CA1031
        try
        {
            return read(state);
        }
        catch
        {
            return onFailure;
        }
#pragma warning restore CA1031
    }
}
```

### H2 — Convert one family per commit

Families, in this order: export toggles (417–440), plot dates (1103–1290), toggle states/labels/images
(912–1000), label dropdowns (1534–2050), size and margin (2059–2600).

For each converted getter: keep the public wrapper, keep the `_ = control;` line where it exists
(IDE0060 is enforced), keep `ArgumentNullException.ThrowIfNull(stateService)` behaviour (it moves into
the helper), and pass the exact failure value the old `catch` returned.

Acceptance:
- `RibbonCallbackSignatureTests`, `RibbonValueContractTests`, `RibbonDropdownContractTests`,
  `GanttRibbonTests` green.
- The per-getter failure values are unchanged. Prove it by reading each old `catch` value before the
  edit and passing it explicitly.
- Diff removes try/catch blocks only.

**Do not** merge the public and internal methods. The internal overloads are the test seam.

---

## 4. Test matrix

| Slice | Must run |
|---|---|
| A | `RibbonValueContractTests`, `RibbonStateServiceTests`, `RibbonDropdownContractTests`, `GanttRibbonTests` |
| B | `CommandBoundaryTests`, `AddRowCommandTests`, `CollapseGroupsCommandTests`, `ParentChildCommandTests`, `MoveCommandTests`, `RemoveCommandTests`, `RepairConfigCommandTests`, `LabelCommandsTests`, `CopyEditableCommandTests`, `SendToPowerPointCommandTests`, `RefreshSheetCommandTests` |
| C | `CommandBoundaryTests`, `RibbonStateServiceTests`, `AddInHostTests`, `RibbonTestCollectionGuardTests` |
| D1 | `RibbonDropdownContractTests`, `CustomStyleRibbonTests`, `DateLabelRibbonTests`, `DescriptionLabelRibbonTests`, `ColumnVisibilityRibbonTests` |
| D2 | `AddRowRibbonTests`, `GanttRibbonTests`, `RibbonValueContractTests` |
| D3 | `PlotDatePickerTests`, `PlotRangeWarningTests`, `RibbonStateServiceTests` |
| D4 | `RibbonStateServiceTests`, `RibbonDropdownContractTests`, `RibbonTestCollectionGuardTests` |
| E | `ThemeSettingsFormTests`, `ThemeSettingsFormContrastWarningTests`, `CustomStyleRenameTests` |
| F | `GanttRefreshSchedulerTests` + build |
| G | `RefreshSheetCommandTests`, `GanttRefreshSchedulerTests`, `ExportPngCommandTests`, `AutoRefreshRibbonTests` |
| H | `RibbonCallbackSignatureTests`, `RibbonValueContractTests`, `RibbonDropdownContractTests`, `GanttRibbonTests` |

Every commit also runs `verify-quick.ps1`. The branch ends with `verify.ps1` on a clean tree.

---

## 5. Risk register

| Risk | Mitigation |
|---|---|
| A silenced argument is accidentally swapped while adding names | Diff check: only `Name:` insertions; no value reordering allowed |
| A presenter helper change alters one command's behaviour | Per-group test runs in B2; the helper takes the delegate, so wording and channel are unchanged |
| `RibbonStateServiceTests` asserts private behaviour | Extract with delegates, not singleton lookups; a test edit means behaviour changed — stop |
| D3 moves the dialog seam and blocks tests | `ShowPlotDatePicker` and `PickPlotDate` stay on the service |
| D4 changes the sticky-snapshot policy | Keep the catch and the previous-value return; only add the record from C2 |
| F fails the build for an unrelated rule | Fix only the timer expression; do not widen the suppression |
| G reorders the tab restore | Keep the tab restore at the call site if the extraction would move it |
| H changes a getter's failure value | Read the old `catch` value, pass it explicitly, one family per commit |
| Singleton state leaks between tests | Use the existing collection guard; reset `CommandBoundary`, `RibbonStateService` and `DiagnosticsService` |

---

## 6. Stop-and-ask conditions

Stop and report, do not improvise, when:

- a slice needs a Core or Office change;
- a test assertion has to change;
- a public or internal signature must change beyond this plan;
- the `RibbonState` constructor shape must change;
- two attempts on the same failure both fail (record both, then stop);
- the build fails on an analyzer rule unrelated to the slice.

---

## 7. Definition of done

Per slice:
- build green with `TreatWarningsAsErrors`;
- the slice's tests green;
- no new suppression, no new dependency, no message change;
- `docs/STATUS.md` updated at branch level, not per commit;
- `git diff --check` clean;
- commit message names the slice and the preserved behaviour, for example:
  `refactor(addin): name RibbonState snapshot arguments (no behaviour change)`.

Branch end:
- `pwsh -NoProfile -File scripts/verify.ps1` green on a clean tree;
- `pwsh -NoProfile -File scripts/check-status.ps1` green;
- handoff as `Outcome / Changed / Proof / Office / Risk / Deferred / Next`, with any gate not run
  reported as `Not run` and why.

---

## 8. Expected outcome

| Metric | Before | After |
|---|---|---|
| `RibbonStateService.cs` lines | 4,416 | about 2,950 after D1–D4 |
| Presenter-policy copies | 15 files / 7 bodies | 1 helper + 1 test |
| Silent failure paths named in this plan | 3 | 0 (each records once) |
| File-wide analyzer suppressions | 1 | 0 |
| Refresh postlude copies | 2 | 1 |
| Unnamed snapshot arguments | 44 | 0 |

The plan does not reduce total line count by much. Its value is that each remaining line has one
reason to change, and every failure leaves a trace.
