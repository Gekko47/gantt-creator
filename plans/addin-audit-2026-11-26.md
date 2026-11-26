# Add-in audit — 2026-11-26

Scope: `src/GanttCreator.AddIn` (64 files, 18,967 lines) and `tests/GanttCreator.AddIn.Tests`
(55 files, 18,456 lines). Load assumed: one analyst, one workbook, one Excel main thread.

VERDICT: **Healthy.** No crash, data-loss or invariant break found. One medium test gap and
three duplication items are worth doing; nothing blocks a release.

Counts: 0 blockers, 0 high, 4 medium, 4 low.

Evidence base: measured line counts, merged Cobertura coverage from the six reports that
`verify.ps1` produced, and `grep` over both projects. Every claim below was read in the source.

---

## Implementation record (2026-11-26)

All nine findings were implemented in seven commits, each green before the next:

| Finding | Commit | Outcome |
|---|---|---|
| 6 — `SafePresent` copy | `6daee742` | deleted; four sites call the shared boundary |
| 2 — duplicated export wording | `6c4cecf2` | parity guard added; Send's four merged catalogue sentences made specific again |
| 1 — postlude never executed | `761546e1` | `FinishRefresh` takes the service; two effect facts replace the source-text pin |
| 4 — theme request projection | `ca3aa9eb` | `ThemeEditRequestFactory` with four facts; service 3,054 → 2,955 lines |
| 5 — unconstructed dialogs | `efff18f9` | four surface facts across the three dialogs |
| 9 — TODO pin | `ee92d0f2` | deleted; R5.10 stays tracked by the roadmap and the XML note |
| 3 — seven fakes | `c7c3260d` | one shared adapter and one shared selection double; 567 lines removed |
| 7 — coverage under target | (follows) | AddIn merged **78.82% line / 71.67% branch**, now over both targets |
| 8 — test class size | planned | Slice 8 below: four topic classes each, with a name-set proof |

Three corrections to the findings above, found while implementing:

1. **Finding 2's premise was too broad.** The read, config and measurement translators really were
   the same skeleton with a different voice. The bounds and shape-write translators are not: they
   differ in grammar ("create the copy" versus "prepare the transfer"), so templating them would
   hide the sentences rather than share them. Only the first three are guarded by the parity test.
2. **Finding 4's "test-only second write" was wrong.** `ShowThemeSettingsDialog_returns_true_when_theme_written_successfully`
   drives that path, so it is supported behaviour and stays. Only the projection moved.
3. **The CS8509 route for new enum members does not work.** A switch expression over a public enum
   must also cover values no member defines, so the default arm cannot be removed. The guard is a
   test instead: no defined member may fall through to that arm.

Deferred, with the reason: `AddInHostTests.RecordingAdapter` stays hand-written. Its purpose is an
interleaved `subscribe`/`detach` log read against another object's log in the teardown-order
contract; folding that into the shared double would couple the double to one test's log.

Verification at branch head: AddIn Debug **983/983**, Release **974/974**, Core 2302, Office
Contract 1767, Office Integration 39, Architecture 142, Raster 127, Release build with
`-warnaserror` **0 warnings**, `verify.ps1` **PASS in 285.5 s**. `verify-office.ps1` **Not run** —
human Excel host; findings 1 and 4 touch live-host paths.

---

## Should fix

### 1. The refresh postlude never runs in a test — its only guard is a text search (P2)

- Where: `src/GanttCreator.AddIn/RefreshSheetCommand.cs:275`, body 284–335; pin at
  `tests/GanttCreator.AddIn.Tests/RefreshSheetCommandTests.cs:223`.
- What it is: after the pipeline returns, one method publishes warnings, marks the chart
  current, clears the status bar, re-asks the selection gates and restores the add-in tab.
- Problem: merged coverage shows **zero hits for every line 284–335**; the class is 26/188
  lines (14%). The only guard reads the file and greps for `FinishRefresh(` and
  `MarkChartRenderCurrent`. Reorder the steps, or drop the tab restore, and that test still
  passes because the strings remain.
- Fix: pass the state service in as a parameter (the method already takes the outcome, both
  timestamps and the presenter) and assert effects with a fake: success marks current and
  repaints the AUTO dates; failure presents exactly once and does **not** mark current; the tab
  restore runs on both paths. Then delete the source-text fact.
- If skipped: a regression that refreshes the chart but leaves Export PNG convinced the chart
  is stale, or leaves the user on the Table Design tab, ships green.

### 2. The export refusal sentences exist twice, and the default arm hides a new reason (P2)

- Where: `CopyEditableCommand.cs:439/455/479/495/598` vs
  `SendToPowerPointCommand.cs:641/659/675/696/712`.
- What it is: five refusal-to-sentence maps per command (read, config, measurement, bounds,
  shape write).
- Problem: the copies are identical apart from the verb ("then copy again" / "then send
  again"). Both end with `_ => ConstructionMessage`, so one new `GanttTableReadRefusalReason`
  member compiles silently and one button shows the generic sentence while the other shows a
  real one.
- Fix: one translator taking the command name; about 200 lines removed; wording stays
  byte-identical.
- If skipped: every new refusal reason is a two-file edit with no compiler help.

### 3. Seven hand-written fakes of the same Excel port (P2)

- Where: `AddInHostTests.cs:471`, `CustomStyleRibbonTests.cs:61`, `DateLabelRibbonTests.cs:109`,
  `DescriptionLabelRibbonTests.cs:105`, `PlotRangeWarningTests.cs:184`,
  `RibbonEventWiringTests.cs:95`, `RibbonStateServiceTests.cs:225`.
- What it is: `IExcelApplicationAdapter` has eight members and each test file re-implements it.
- Problem: one new non-default member breaks seven files. The same holds four times over for
  `ISelectionIdentifier`, `IGanttTableReader`, `IConfigCatalogueReader`, and six times for
  `IPanelGridMeasurementPort`.
- Fix: one shared `TestExcelApplicationAdapter` with optional counters and throw flags; the repo
  already does this well for label commands (`LabelCommandsTests.cs:214`).
- If skipped: every port change stays a seven-file chore and the copies diverge.

### 4. `ShowThemeSettingsDialog` does four jobs and keeps two copies of the write policy (P2)

- Where: `RibbonStateService.cs:1426`, 187 lines.
- What it is: reads the workbook, projects the stored theme into a dialog request (1450–1544),
  injects a write callback (1561–1575), then repeats the write-and-report block after the
  dialog closes (1591–1601).
- Problem: the projection is pure logic over `ConfigReadOutcome` with no direct test, so "open
  the dialog showing what the chart currently draws" is proven only through the form's tests.
  The second write block is reachable only when no writer was injected (a unit-test seam), so
  production carries a duplicate of the write rule for test convenience.
- Fix: move 1450–1544 to a static factory with three facts; then remove the fallback write.
- If skipped: the dialog can open with the wrong palette after a reader format change.

### 5. Three dialogs are never constructed by a test (P3)

- Where: `ChartTitleForm.cs` (0/88 lines), `PlotDatePickerForm.cs` (0/84),
  `WarningsPanelForm.cs` (0/91) in merged coverage.
- What it is: the title dialog, the plot-date picker, the warnings panel.
- Problem: the decision logic is tested in Core (`ChartTitleEditRequestTests`, 5 facts), so
  runtime risk is small — but the wiring (which request is stored, which `DialogResult` is set,
  whether the panel lists current issues) is proven only live. `ChartTitleForm:79` already
  exposes `GetValidatedRequest`, so the seam exists.
- Fix: one construction plus one accept fact each, on the existing STA pattern
  (`ThemeSettingsFormTests.cs:38`). `PlotDatePickerForm` needs an internal accessor for the
  chosen date.
- If skipped: a wrong default or an empty panel ships silently.

---

## Nice to have

### 6. The last copy of `SafePresent` (P3)

`ValidateSheetCommand.cs:118`, called at `:104` and `:108`, is byte-identical to
`CommandBoundary.SafePresent` (`CommandBoundary.cs:291`). Deleting it removes ~20 lines and one
`CA1031` guard; behaviour identical.

### 7. Coverage sits under the AddIn target (P3, tracking only)

Merged line 74.17%, branch 68.63% against 75% / 65%. Branch passes, line does not; enforcement
is deferred per `docs/04-TEST-STRATEGY.md`. Finding 1 is most of the gap.

### 8. Two test classes hold 4,604 lines (P3)

`GanttRibbonTests.cs` 2,625 and `RibbonStateServiceTests.cs` 1,979. No behaviour risk; finding a
fact is slow. Split by feature on the next touch — by job, not by line count.

### 9. A test pins a TODO in place (P3)

`ParentChildRibbonTests.cs:251` asserts the ribbon XML still contains `TODO(R5.10)`. Finishing
that owed work and removing the comment turns a green test red.

---

## What is healthy and should not be changed

- No `Thread.Sleep` anywhere; the scheduler tests wait on named signals with deadlines
  (`GanttRefreshSchedulerTests.cs:60`).
- No skipped tests.
- 971 executing tests against 18,456 test lines, for 18,967 production lines.
- Each of the 57 `CA1031` guards in the AddIn states its own reason.
- The two source-text facts about COM wiring (`RefreshSheetCommandTests.cs:150`) guard a hazard
  no unit test can observe, and are documented as such.

Lean: about −220 production lines (200 duplicated translators, 20 `SafePresent`) and about −200
test lines (fake consolidation). No dependency to remove.

Not checked: Core, Office and Raster beyond the calls the AddIn makes; the live Excel gate; the
runtime behaviour of the three dialogs.

---

## Implementation plan for Code Mode

Order is risk-first: pure deletions first, then the seam work, then the behavioural test.

| # | Slice | Files | Acceptance |
|---|---|---|---|
| 1 | Delete the `SafePresent` copy (finding 6) | `ValidateSheetCommand.cs` | behaviour identical; validate/ribbon suites green |
| 2 | One export refusal translator (finding 2) | new `ExportRefusalMessages.cs`, `CopyEditableCommand.cs`, `SendToPowerPointCommand.cs`, tests | every existing message byte-identical, proven by the command tests |
| 3 | Shared Excel adapter fake (finding 3) | new `TestExcelApplicationAdapter.cs`, seven test files | no test asserts change; suite green |
| 4 | Postlude seam and behavioural facts (finding 1) | `RefreshSheetCommand.cs`, `RefreshSheetCommandTests.cs` | the source-text pin is replaced by effect facts |
| 5 | Theme request factory (finding 4) | new `ThemeEditRequestFactory.cs`, `RibbonStateService.cs`, tests | three projection facts; no duplicate write |
| 6 | Dialog accept facts (finding 5) | `ChartTitleForm.cs` (no change), `PlotDatePickerForm.cs`, `WarningsPanelForm.cs`, tests | one accept and one cancel fact each |
| 7 | TODO pin (finding 9) | `ParentChildRibbonTests.cs` | the XML comment is asserted as owed work without failing a completed change |

Finding 7 closed with slice 4: AddIn merged coverage went 74.17 → 78.82 % line and
68.63 → 71.67 % branch, over both targets. Finding 8 is Slice 8.

### Slice 8 — split the two oversized test classes (finding 8)

**Goal.** `GanttRibbonTests.cs` (2,625 lines, ~94 facts) and
[`RibbonStateServiceTests.cs`](../tests/GanttCreator.AddIn.Tests/RibbonStateServiceTests.cs)
(1,979 lines, ~95 facts) become topic classes each, in the same xUnit collection. No test method
changes, no assertion changes, no production change. The ribbon half is finished: the file is now
[`GanttRibbonXmlTests.cs`](../tests/GanttCreator.AddIn.Tests/GanttRibbonXmlTests.cs),
[`GanttRibbonChartTests.cs`](../tests/GanttCreator.AddIn.Tests/GanttRibbonChartTests.cs) and
[`GanttRibbonCommandTests.cs`](../tests/GanttCreator.AddIn.Tests/GanttRibbonCommandTests.cs).

**The constraint that shapes the plan.** `RibbonTestCollectionGuardTests` reads every top-level
`*.cs` and fails a file that names `RibbonStateService`, `GanttRibbon.`, `DiagnosticsService.Instance`
or `CommandBoundary.Instance` without the literal `diagnostics-service` in that same file. It does
not distinguish a test class from a support type, so:

- `partial` parts cannot work — a part that names the singleton trips the rule, and a duplicate
  `[Collection]` on one type is wrong;
- a helper file that builds a `RibbonStateService` trips it too.

**Decision.** Separate topic classes, each carrying `[Collection("diagnostics-service")]`, plus one
refinement to the guard: it considers only files that declare a test class, with a new positive
control for the support case. The invariant is unchanged — a file that never runs a test cannot
interleave a record into another test's log.

**Status 2026-11-26: steps 1–3 and two of the four ribbon moves have landed.**

1. Done (`9661f9e6`). `RibbonTestSupport` holds the ribbon instance, the callback vocabulary, the two
   XML validators, the invalidation reader, the callback resolver and the plot-bounds row builder.
   `using static` keeps every call site unchanged; the class is 2,625 → 2,448 lines.
2. Done (`c1d73122`). The tripwire now applies only to a file that declares a fact or a theory, with
   a new shared-fixture control beside the hazard and compliant cases.
3. Done (`ae6fae42`). `RibbonStateTestSupport` takes `Arm`, `ArmedWithTableRead`,
   `CountInvalidations`, `NewRowId` and `Row`; the class is 1,866 → 1,813 lines. **An abstract base
   was tried and rejected**: it forced the derived test class and the base to `public`, which dragged
   the analyzers' public-surface rules (CA1062, CA1816, CA1001) into a test fixture. Each class keeps
   its own constructor and `Dispose` instead, so the singleton reset stays visible where it applies.

Ribbon moves, one commit each:

4a. Done (`e0ce22f0`). Chart and plot controls — 34 facts from
    `Chart_group_declares_start_finish_and_scale_rows` through
    `GetPlotStartDate_throws_for_a_null_state_service` — in `GanttRibbonChartTests`; the class is
    2,448 → 1,592 lines and the new file is 840.
4b. Done. Command routing — 16 facts: the four callbacks that cross `CommandBoundary`, the two
    theme-dialog clicks, the null-dependency refusals — in `GanttRibbonCommandTests`; the class is
    1,592 → 1,210 lines. The theme-dialog double became `StubThemeDialogStateService` in
    `RibbonTestSupport`, because the routing facts and the state-service topic both need it and
    neither can reach the other's private nested copy. Proof: suite 983, the distinct `--list-tests`
    name set identical to the pre-split baseline apart from that baseline's own `csproj`/`dll`
    artefacts, and the attribute arithmetic holds — HEAD's 60 facts = 44 left + 16 moved.
4c. Done. XML and surface contracts — 32 facts — in `GanttRibbonXmlTests`; the remaining 12
    routing and getter facts (`OnDiagnosticsClick`, `OnRepairConfigClick`, `ResolveCommandName`,
    `OnLoad`, `GetEnabled`, `GetVisible` fail-closed) joined `GanttRibbonCommandTests`, so
    `GanttRibbonTests.cs` is deleted rather than renamed. The three topic classes hold all 94 facts
    — 34 chart, 32 XML, 28 command — and the XML file carries the file-scoped `CA2000` suppression
    the original had. Proof: suite 983, distinct name set identical, and `docs/REPO-MAP.md` plus
    the four historical `STATUS.md` links now name the topic classes instead of the deleted file.
5. Move `RibbonStateServiceTests` topic by topic: snapshot and modes (~600), gates and activation
   (~350), capture and invalidation (~300), preferences and dialogs (~600). The three nested stubs
   (`StubThemeDialogRibbonStateService`, `HostApplicationRibbonStateService`,
   `FakePanelGridMeasurementPort`) move with the topic that uses them, or into the support type when
   two topics use one.
6. The ribbon leftover needed no rename: it was deleted once its last fact moved, which is the
   equivalent. The same rule still applies to `RibbonStateServiceTests`.

**Tooling caution.** A name-driven mover was written and abandoned: on its first real run it reported
a method missing that was present, and the working tree afterwards matched HEAD exactly, so the
mechanism is not understood well enough to trust. Move by hand, or write a mover that asserts the
before and after name sets itself — do not script a 4,600-line move on an unverified parser.

**Proof after every move.** `dotnet test --list-tests`, reduced to the sorted simple method names,
must be identical before and after (fully-qualified names change with the class). The suite totals
must hold at 983 Debug / 974 Release, and the collection guard must stay green. A compile error
catches a broken move; only the name-set diff catches a dropped fact.

**Risks.**

- A lost fact in a move is the whole risk; only the name-set diff sees it.
- A new class that forgets the `Reset()` lifecycle leaks the singleton between tests; the shared
  base is what prevents it.
- `_nextPlotBoundsId` is the only mutable static and is safe only while every class stays in the
  serialised collection.
- It pays back in navigation and merge-conflict surface, not correctness; if these two files are
  rarely edited, schedule it for when they are next touched rather than immediately.

Per-slice rules: one commit each; `dotnet build` with analyzers; the named suites; Debug and
Release; `verify-quick.ps1`; `verify.ps1` branch-final; `verify-office.ps1` Not run with the
reason stated (human host), because findings 1 and 4 touch live-host paths.
