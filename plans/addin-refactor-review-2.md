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

### Finding 3 — the service still carries the scheduling cohort (P2, resolved 2026-11-26)

Activation, the three subscriptions, the scheduler lifetime, teardown, the theme and chart-title
dialogs, the PNG preferences, the warnings panel, the status bar and the `GetEnabled` truth table
remain. The file is still the edit point for six unrelated features.

What landed: `RibbonEventWiring` owns the four Excel connections and the coalescing scheduler —
attach, replace, detach once, and a refused subscribe degrades to an inert connection instead of
failing the load. The handlers stay on the service, because each transitions state the service
owns, so `Activate` and `Teardown` delegate to one wiring object and nothing else moved. The
service is 3,251 → 3,054 lines with 19 → 12 local `CA1031` guards, five new facts pin the seam in
`RibbonEventWiringTests`, and no existing test changed.

Still open: the theme and chart-title dialogs, the PNG preferences, the warnings panel, the status
bar and the `GetEnabled` truth table remain on the service. Each owns its own state, so each is a
separate seam and a separate work item.

`verify-office.ps1` was **Not run** — the plan makes the live host gate this slice's acceptance and
that gate needs the human machine.

### Findings 4–8 — all resolved or assessed 2026-11-26

- Selection gate copied four times → `CommandPreconditions.TryRequireSingleBodyRow`, with each
  command keeping its own translator and five facts pinning the shared half.
- Dependency ladder copied twice → **assessed and deliberately not implemented.** The two ladders
  differ in `Refuse` arity (four arguments versus five, because the PowerPoint outcome carries
  `presentationCreated`) and every step names a distinct refusal enum member, so a shared `Require`
  needs either an `out` chain per step (four lines either way) or a bool parameter that couples the
  two commands. It would hide the ladder without shortening it; the explicit ladder is the more
  greppable form and its duplication is a compile-time reminder when a request member is added.
- `BuildThemeEditRequest` is still one method with four diffing loops → **done.** `CustomStyleState`
  and `EntityStyleColours` were already `internal` at file scope, so the four diffs are now named
  private steps on the form, and the custom-style half is renames, deletes, then creates/updates in
  the writer's order. A dead local was removed.
- Ten getter guards remain → **done.** `GetPlotStartAuto` and `GetMargin` convert to
  `RibbonGetter.Read`; the other eight state why the single-read helper does not fit (a COM read of
  Excel's own control, a mutating command, a two-value projection, or a multi-statement search).
  `GanttRibbon` holds 8 `CA1031` guards against 41 before the work.
- Static session state in `RefreshSheetCommand` is not reset by `AutoClose` → **done.**
  `ResetSessionState` drops the window (and its port over the session's application object) and
  invalidates the delta cache and arming memory as teardown step 2, with a fact that arms the
  shared state and asserts the teardown clears it.

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
