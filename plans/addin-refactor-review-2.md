# Add-in refactor — post-implementation review

Status: review complete. Produced in read-only `code-review` mode; no product code was changed by the review.
Date: 2026-11-26.
Scope: commits `a37394a9..4fad4096` (nine commits, 33 changed-file entries) measured against
[`plans/addin-refactor-fix-plan.md`](addin-refactor-fix-plan.md).

Verdict: **Ship with fixes.** No correctness defect and every gate is green. Two plan-named
hazards survive and one new failure path has no test.

Counts: 0 blockers, 0 high, 3 medium, 5 low.

Top risk: the six-`bool?` probe list still travels positionally at two call sites, so the exact
swap hazard that opened this refactor can still compile.

---

## 1. Plan against code

| Slice | Status | Note |
|---|---|---|
| A — name snapshot arguments | Done | factory call named; the two probe lists are not (finding 1) |
| B — one presenter boundary | Done | 15 copies gone; `SafePresent` has 3 tests |
| C — three failure records | Done | each has a positive test |
| D1–D4 — collaborators | Done | 4 new types, +1,576 lines |
| D5 — scheduling cohort | Deferred by the plan | still inside the service |
| E — theme handler split | Done in-file | separate builder/validator types not created |
| F — pragma removal | Done | `TreatWarningsAsErrors` build is the proof |
| G — refresh postlude | Done, pin strengthened | asserts delegation and one marker |
| H — getter boundary | Done | 26 of 36; 15 pragmas remain in the file |
| A3 — `SnapshotFacts` struct | **Not done** | deferred, and now the main gap |

Measured effect:

| File | Before | After | Pragma before → after |
|---|---|---|---|
| `RibbonStateService.cs` | 4,416 | 3,251 | 26 → 19 |
| `GanttRibbon.cs` | 2,815 | 2,574 | 41 → 15 |
| `ThemeSettingsForm.cs` | 3,154 | 3,206 | 3 → 3 |
| Presenter helper copies | 15 files / 7 bodies | 1 helper | — |
| AddIn test facts (Debug) | 952 | 958 | — |

---

## 2. Findings

| # | Severity | Where | Issue |
|---|---|---|---|
| 1 | P2 | `RibbonStateService.cs:1945`, `RibbonStateService.cs:2866` | Six adjacent `bool?` values passed positionally, twice. |
| 2 | P2 | `RibbonGetter.cs:35` | The getter `catch` has no positive test; no test references the type. |
| 3 | P2 | `RibbonStateService.cs:1903` | 3,251 lines and still about eight reasons to change. |
| 4 | P3 | four label/style commands | The single-row selection gate is still copied four times. |
| 5 | P3 | `CopyEditableCommand.cs`, `SendToPowerPointCommand.cs` | The 10-arm dependency ladder is still written twice. |
| 6 | P3 | `ThemeSettingsForm.cs` | 3,206 lines; the request builder is still one long method. |
| 7 | P3 | `GanttRibbon.cs:1175` | Ten getter guards remain, none marked as deliberate. |
| 8 | P3 | `RefreshSheetCommand.cs` | Static mutable session state still outlives `AutoClose`. |

### Finding 1 — the probe list still travels positionally (P2, must-fix)

`Refresh` builds six nullable bools and passes them in order to `CaptureSnapshot`, which passes the
same six in order to `RibbonSnapshotFactory.Capture`. Four names differ by one word and share a
type. Swapping `configurationPresent` and `ganttSheetPresent` compiles and changes the Ribbon mode
with no failing test.

Fix: the deferred A3. One `SnapshotProbeFacts` record struct, constructed with named arguments at
each call site, replaces the six parameters. Behaviour is unchanged.

### Finding 2 — the getter boundary's failure arm is untested (P2, must-fix)

`RibbonGetter.Read` returns the caller's safe value on a throw. No test names `RibbonGetter`; the
getter tests assert only the null-argument throw. Dropping the `catch` would keep every test green
while a COM probe failure started surfacing in Excel. This breaks the repository rule that a
failure path ships with a test that fires it.

Fix: two facts — the safe value is returned when the read throws, and the null guards hold.

### Finding 3 — the service still carries the scheduling cohort (P2, nice-to-have, own change)

Activation, the three subscriptions, the scheduler lifetime, teardown, the theme and chart-title
dialogs, the PNG preferences, the warnings panel, the status bar and the `GetEnabled` truth table
remain. The file is still the edit point for six unrelated features.

Fix: the plan's D5 as its own change, guarded by `AddInHostTests` plus the Office gate. Do not fold
it into a bug fix.

### Findings 4–8 (P3, payback on next touch)

- Selection gate copied four times: `SetDescriptionLabelCommand`, `SetStyleKeyCommand`, and the two
  date-label commands.
- Dependency ladder copied twice: `CopyEditableCommand`, `SendToPowerPointCommand`; one `Require`
  helper with the refusal factory removes a 40-line copy.
- `BuildThemeEditRequest` is still one method with four diffing loops; plan E1 needs
  `CustomStyleState` and `EntityStyleColours` promoted before a separate type can exist.
- Ten getter guards remain (for example `GetPlotStartAuto:1175`, `GetTimeScale:1334`,
  `GetMargin:1890`); each should state why the single-read helper does not fit.
- Static session state in `RefreshSheetCommand` (protection window, shape cache, arming) is not
  reset by `AutoClose`; low harm today because the Excel application outlives the add-in.

---

## 3. What is genuinely better

One authority each for presenter failures, getter failures, dropdown projections, the add-entity
model, plot-range derivation and snapshot construction. The three silent paths leave a record and
each has a test. The file-wide analyzer suppression is removed rather than relocated. The refresh
pin asserts a fact the code cannot satisfy by accident.

---

## 4. Implementation order for Code Mode

1. **Finding 1** — add `SnapshotProbeFacts`, change `Capture` and `CaptureSnapshot`, update the
   three call sites. Gate: `RibbonValueContractTests`, `RibbonStateServiceTests`,
   `RibbonDropdownContractTests`, `GanttRibbonTests`.
2. **Finding 2** — two `RibbonGetter` facts. Gate: `RibbonValueContractTests`.
3. **Finding 5** — one `Require` helper for the two export ladders (optional in the same session).
4. **Finding 4** — one selection-gate helper, keeping each command's own translator.
5. Findings 3, 6, 7, 8 — separate changes; 3 needs the Office gate.

Each step: build with analyzers, run the named suites, commit with a Conventional Commit message.
The suite baseline is AddIn 958/958 (Debug) and the Release non-Office suite is green.
