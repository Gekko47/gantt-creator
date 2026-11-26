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
| 8 — test class size | deferred | a split is a next-touch change, not a mechanical one |

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

Findings 7 and 8 are tracking notes, not code changes: coverage follows from slice 4, and the
test-class split is deliberately deferred to the next touch of those files.

Per-slice rules: one commit each; `dotnet build` with analyzers; the named suites; Debug and
Release; `verify-quick.ps1`; `verify.ps1` branch-final; `verify-office.ps1` Not run with the
reason stated (human host), because findings 1 and 4 touch live-host paths.
