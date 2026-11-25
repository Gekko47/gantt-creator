# P1b and P7 — pre-implementation review (2026-11-25)

Owner instruction: review both items before implementing, investigate the best option and the
installation procedure. Both were approved earlier the same day
([`office-suite-runtime-implementation-plan.md`](office-suite-runtime-implementation-plan.md),
"Owner decisions, approved 2026-11-25"); this review fixes the *how*, and in P7's case the
*whether-now*.

Evidence base: the 19:30 live run's teardown records
(`scripts/_artifacts/office-evidence/office-forced-kills-shard-0.log`, 68 records with
`BodyMs`/`SetupMs`/`TeardownMs`), the eight classes' sources, the xUnit documentation for
trait placement and class-fixture lifetime, and [`verify-office.ps1`](../scripts/verify-office.ps1).

## Verdict

- **P1b — built, and the live run confirms the shape.** The plan's "~12 bodies" was an estimate
  made before counting: the eight named classes carry **24 test cases across 22 methods** (one method
  is the 3-row `Initialise_preserves_a_cell_empty_sheet_with_live_non_cell_state` Theory; discovery
  counts distinct methods, so the coverage line reads 22). Measured wall on the first live `-Smoke`
  run: **165.4 s test process, ~3.5 min end to end**, against the 420 s bound.
- **P7 — approved, but do not start yet.** Its measured net ceiling is **~30–60 s**, the same order
  as P12's refused 35 s, and the classes with the most bodies are exactly the classes whose bodies
  deliberately start from *different* states, so a shared workbook spends on reset what it saves on
  initialise. It starts only if P1b's measured smoke run leaves the full gate too slow for the
  branch cadence, and its first step is a per-class state-delta audit, not a fixture rewrite.

## P1b — the best option and the installation procedure

### Scope correction (the plan's estimate vs the tree)

| Class | Cases | Notes |
|---|---|---|
| `AddRowIntegrationTests` | 7 | all write teardown records |
| `InitialiseSheetIntegrationTests` | 8 | 5 facts + the 3-row Theory |
| `StagingLifecycleIntegrationTests` | 2 | both record |
| `AutoRefreshTriggerIntegrationTests` | 2 | 1 records, 1 suppresses |
| `TextRenderIntegrationTests` | 2 | both suppress |
| `ShapeRenderIntegrationTests` | 1 | suppresses |
| `StyleRenderIntegrationTests` | 1 | suppresses |
| `RefreshIdempotenceIntegrationTests` | 1 | suppresses |
| **Total** | **24** | 18 record, 6 suppress |

The six suppressed bodies set `SuppressLeakSignal = true` (they render real shapes, which legitimately
forces a kill), so they are absent from the kill log but still cost full fixture time — a targeted
`ShapeRenderIntegrationTests` run measured **13 s wall**. The smoke estimate must count them.

### Wall estimate, from the 19:30 run's records

- 18 recorded bodies: body ≈ **45.3 s** summed, setup ≈ 15.3 s (mean 0.85 s), teardown ≈ 57.6 s
  (mean 3.2 s, dominated by the 3 s graceful poll). The "runner ≈ 2.3 s per test" term written
  here was **withdrawn after the smoke run**: the host-free control measured the runner at
  ~2.5 s *fixed* (39 facts in 2,655 ms wall against 210 ms of test time), so this estimate was
  high by ~40 s; the observed run came in at 160.3 s.
- 6 suppressed render bodies: ≈ **55–65 s** (render + read-back + their own setup/teardown).
- Host start and discovery: ≈ 20 s.
- **Total ≈ 215–235 s ≈ 4 min.** The acceptance run replaces this estimate with a measurement.

### The option chosen, and the one risk that decides the design

**Tag placement: class-level `[Trait("Category", "OfficeSmoke")]` on the eight classes.** xUnit
applies traits declared on a test class to every method in it (the docs state traits attach to "a
test method, class, or assembly"), so one line per class covers existing and future bodies.
Rejected alternatives: per-body tagging (drifts the moment a new test is added — the exact failure
mode the plan's own mitigation warns about) and a class-name list inside the gate script (a second
source of truth for one set, which the coverage proof would then have to reconcile).

**The integration risk:** [`verify-office.ps1`](../scripts/verify-office.ps1) hard-codes
`Category=OfficeIntegration` in **three** places — discovery
([`:447`](../scripts/verify-office.ps1:447)), the single-shard filter
([`:512`](../scripts/verify-office.ps1:512) and [`:528`](../scripts/verify-office.ps1:528)), and the
multi-shard conjunction terms ([`:540`](../scripts/verify-office.ps1:540)). A `-Smoke` run must
thread **one** filter variable through all three. If discovery keeps the full category while the
run filters to smoke, the coverage proof compares the smoke trx against the full discovered list and
fails `exit 6` in the missing direction — a false failure wearing the proof's own verdict. The
filter variable is therefore the whole design.

### Installation procedure (one commit, one live run)

1. Add `[Trait("Category", "OfficeSmoke")]` at class level on the eight classes. No method-level
   change; `OfficeIntegration` stays where it is.
2. **Mechanical verification before touching the gate:** `dotnet test
   tests/GanttCreator.Office.IntegrationTests -c Release --no-build --list-tests --filter
   "Category=OfficeSmoke"` must list exactly the 24 expected cases. If the count differs, stop — the
   class-level trait did not reach the methods and nothing downstream can be trusted.
3. [`verify-office.ps1`](../scripts/verify-office.ps1): add `[switch]$Smoke`; derive
   `$categoryFilter = $Smoke ? 'Category=OfficeSmoke' : 'Category=OfficeIntegration'` once, near the
   parameters; use it at all three filter sites; default the effective deadline to **420 s** when
   `-Smoke` is passed and no explicit `-DeadlineSeconds` was given (the plan's own number).
4. Pester, in the same commit: the stub-driven PASS case gains a `-Smoke` variant asserting the
   filter text and the 420 s deadline; a source-pin fact asserts all three filter sites use the one
   variable (positive control included, per the validator rule).
5. Review the tag set against the entity guide's renderer-equivalence table — the plan's own
   mitigation for a tag that drifts from the risk surface.
6. **Acceptance:** one live `verify-office.ps1 -Smoke` — wall ≤ 7 min,
   `coverage: all 24 discovered test(s) ran in a shard, and nothing else ran.`, the reference-count
   and forced-exit verdicts unchanged — and one full run unchanged (104 tests, same verdicts).

## P7 — the options, the measurement that gates it, and the procedure if it starts

### The four options

- **A. `IClassFixture<OfficeFixture>` as-is — rejected.** xUnit constructs a class fixture *before
  any test in the class runs*, and [`OfficeFixture`](../tests/GanttCreator.Office.IntegrationTests/OfficeFixture.cs)
  resolves the test name at construction via a stack walk (`_testName = ResolveCallingTestName()`).
  Under a class fixture that walk finds no test frame, so the per-body attribution in the teardown
  record breaks, and with it the per-teardown forced-exit ceiling and the per-body meaning of the
  `Rcw` gate. Option A is P12's isolation risk *plus* attribution damage.
- **B. Share the workbook, keep per-body fixtures — impossible.** The workbook lives inside the
  Excel process the fixture owns; sharing the workbook means sharing the process, which is P12.
- **C. Per-class Excel + one initialised workbook + a reset protocol between bodies — the real P7.**
  The reset surface, enumerated from the landed model: table body rows and the initial blank row,
  hidden flags, outline levels and indent, row heights, the VeryHidden config sheet's catalogues,
  defined names (plot anchor, type options), protection state, owned shapes across the entity
  guide's 26 types, validation notes, and the reserved anchor and padding rows.
- **D. Cut the initialise cost itself — out of scope.** That is product performance work
  ([`refresh-latency-trace-and-opportunities.md`](refresh-latency-trace-and-opportunities.md)'s
  territory), not test-harness work.

### The measurement that gates it

Initialise costs **≈2.5–3.6 s per body** on this host (`Initialise_adopts…` BodyMs 3592,
`ConfigCatalogue…` 3544, the AddRow bodies 1366–1882 including initialise). But the saving is only
*(bodies-per-class − 1) × initialise*, and only for bodies that share a start state. The multi-body
classes are the state-variant ones:

- `InitialiseSheetIntegrationTests`' eight cases each test initialise **from a different starting
  state** — blank, protected, table-exists, three non-pristine artefacts, non-blank active. A shared
  initialised workbook is impossible without rebuilding each state, which spends what it saves.
- `AddRowIntegrationTests`' seven bodies insert from different selections (in-table, outside,
  first-activity, protected, second workbook).
- The render classes have 1–2 bodies each: nothing to share.

The honest net ceiling is **≈30–60 s** — the same order as P12's refused 35 s, against a much larger
risk surface (ADR-0020/ADR-0055 treat retained workbook state as a first-class hazard).

### Installation procedure (only if it starts)

1. **Per-class state-delta audit first:** for the candidate class, list each body's required start
   state. A class qualifies only if every body starts from the same post-initialise baseline.
   Classes that fail the audit are recorded and excluded — that record is the deliverable, not a
   failure.
2. Write the reset protocol and the test-strategy note (the ADR-0020/ADR-0055 hazard) before any
   fixture change. The protocol must restore every surface listed under option C, and it must be
   fail-closed: a reset that cannot prove a surface clean fails the body rather than passing it.
3. Migrate **one** class — `AddRowIntegrationTests` is the cheapest proof (seven bodies, one shared
   baseline, no render residue) — with the **two-order proof in the same commit**: the class's facts
   pass in two different orders.
4. Measure from the phase signal: the class's wall before and after. **Revert if the net saving is
   under ~20 s** — the plan's own rollback rule.
5. Only then consider the next class, one commit at a time.

## Sequence

1. **P1b now** — one commit, one live `-Smoke` run, one unchanged full run.
2. **P7 waits** for P1b's measured result. If the smoke run meets the ≤ 7 min acceptance and the
   full gate stays a release-only gate, P7's ceiling does not justify its risk and the decision is
   re-read, not re-made.
3. The runner work item is **closed on its own measurement, as a negative.** The host-free control
   (39 facts, 2,655 ms wall against 210 ms of test time) and the trx sum (650.2 s of a 661.1 s
   process) put the runner's whole overhead at **10.9 s (1.6%)**; the rest of what the first phase
   line called "outside every fixture" is fixture time in the 38 bodies that write no teardown
   record. There is no test-host work item to do. The lever that replaced it is the
   **graceful-exit poll: 144.9 s of the 650 s**, because two-thirds of teardowns are killed and each
   pays the full 3 s while a clean exit is observed at 0 ms — a 1 s window would save ~92 s, and it
   needs its own measured run before adoption.
