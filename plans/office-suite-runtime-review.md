# Office integration suite — runtime triage

**Verdict:** the gate is slow because it runs one real Office process per test body, serially, and
because about four minutes of it is a host-state stall rather than work. The 780 s deadline cannot
pass on this host; raising it is the only change that makes the gate able to decide, and the stall
is the cheapest real saving.

**Scope:** `tests/GanttCreator.Office.IntegrationTests`, `scripts/verify-office.ps1`,
`tests/GanttCreator.Office.IntegrationTests/xunit.runner.json`. Measured on this machine
(Excel/PowerPoint 16.0.20625.20108, x64, en-us) on 2026-11-25. Not reviewed: `verify-quick.ps1`
timings, the AddIn suite, and Office behaviour on other hosts.

**Counts:** 1 blocker, 3 high, 3 medium, 1 low.

**Top risk:** `verify-office.ps1` kills the suite at 13 minutes and reports nothing, so every run
on this host is undecided rather than red.

> **Corrections, 2026-11-25.** Three statements below are wrong; they are kept as the record of
> what was measured. See
> [`plans/office-suite-runtime-implementation-plan.md`](office-suite-runtime-implementation-plan.md)
> for the corrected baseline and the plan.
>
> 1. **The soak is not in the gate.** `PowerPointTransferSoakIntegrationTests` is tagged
>    `TransferSoak` **only**, and `scripts/verify-office.ps1:414` filters
>    `Category=OfficeIntegration`. Finding 8 and option E are withdrawn; they are already satisfied.
> 2. **The 25.2 min figure is the unfiltered assembly.** The gate's own archived run held 105 tests
>    in 16.9 min, so the deadline is ~4 minutes short of ~17, not 12 short of 25. The table below
>    mixes both scopes and must be re-measured with the gate's filter before any change.
> 3. **Finding 5's "1–5 min" contradicts its own measurement.** A teardown that routinely spent the
>    10 s window could not produce the 3.2 s open/close cycle quoted in the same finding. Treat the
>    saving as a hypothesis until a per-class teardown measurement exists.

---

## Measurements

| Source | Tests | Sum of test durations | Notes |
|---|---|---|---|
| `office-manual.trx`, 2026-11-25 00:43–01:09 | 142 | **25.2 min** | wall 25 m 13 s; 140 passed, 2 failed |
| `office-20261124-101942282.trx`, 2026-11-24 | 105 | **16.9 min** | 101 passed, 4 failed |
| `verify-office.ps1` deadline | — | 780 s (13 min) | stopped before either figure |

Sum of durations equals wall time to within 10 s. The suite is fully serial, so there is no idle
time to reclaim: every second saved has to come out of a test body.

---

## Findings

### 1. The gate's deadline is below the measured runtime (blocker)

Where: `scripts/verify-office.ps1:41` (`DeadlineSeconds = 780`).

What happens: the suite needs 1,516 s end to end on this host; the watchdog stops it at 780 s, no
trx is produced, and the script reports `TIMEOUT`.

Why it matters: the gate can never go green here whatever the code does, so it stops being a gate.
The script's own comment already says to raise it only on a fresh measurement — that measurement
now exists.

Fix: raise the default to 2,400 s (40 min) and record the measurement beside it. This is a policy
change, not a code change; it saves no time.

### 2. Serial execution is a deliberate setting, not an accident (high)

Where: `tests/GanttCreator.Office.IntegrationTests/xunit.runner.json:1`.

What happens: `"parallelizeAssembly": false` and `"parallelizeTestCollections": false`.

Why it matters: 142 tests against COM hosts, a global clipboard, and shared Excel application
state. Turning either flag on without isolation would trade a slow gate for a flaky one.

Fix: do not flip the flags. Add a second collection only for provably independent read-only Excel
bodies, and only after a spike proves two Excel instances coexist without state bleed.

### 3. Four PowerPoint bodies stall for 60 seconds each (high)

Where: `PowerPointTargetDecisionIntegrationTests.cs:1` (`A_foreground_protected_view…`,
`Two_presentations_resolve…`), `PowerPointTransferFailureIntegrationTests.cs:1`
(`A_presentation_with_no_slides…`), `PowerPointPasteSpecialIntegrationTests.cs:1`
(`An_empty_clipboard_paste…`).

What happens: each took **61.5–61.6 s** in the full run. The same body alone takes **4 s**
(`A_presentation_with_no_slides_maps_to_NoSlide`, measured 2026-11-25). In the archived run of
2026-11-24 the same bodies took **0.5 s**.

Why it matters: 4 min of the 25 min is a stall, not work, and it is invisible in the gate because
the run never finishes. A 60 s cluster with 0.1 s spread is a timeout elapsing, not variable work.

Proof: durations above, plus the 4 s `-filter` run of one of the four. Root cause is **not**
proven. Two candidates, both testable: (a) the next body activates
`new PowerPoint.Application()` while the previous fixture's force-killed process is still leaving,
and the DCOM activation waits out its default timeout; (b) `Presentations.Open` on the
Protected-View bodies waits on a host-side timer.

Fix: instrument `PowerPointFixture` with per-step timings (start, open, attach, quit) and run the
four bodies with `--filter` in the full-suite order. Then make teardown wait — bounded, with the
owned-PID manifest — for its process to disappear before returning, so the next activation never
races a dying server.

### 4. PowerPoint process churn per body (high)

Where: `PowerPointFixture.cs:141` (`Initialize` calls `new PowerPoint.Application()`), used by
every PowerPoint body (`PowerPointTargetDecisionIntegrationTests.cs:71`, `:92`, `:112` and so on).

What happens: every body creates a fixture, starts a real PowerPoint, and quits it. There is no
shared host. Even at the healthy 4 s figure, ~20 bodies pay start plus quit twice.

Why it matters: PowerPoint start/quit is the single largest repeated fixed cost in the suite after
Excel, and most bodies do not need a pristine process — they need a deck.

Fix: one fixture per test class, handed to bodies that accept an existing host.
`New_power_point_with_no_presentations_decides_CreatePresentation` asserts that no POWERPNT is
running (`:36`), so that body must stay isolated; everything else can share.

### 5. The Excel graceful-exit window is usually spent for nothing (medium)

Where: `tests/GanttCreator.Office.IntegrationTests/OfficeFixture.cs:113`
(`DefaultGracefulExitSeconds = 10`), `:748` (`DisposeAsync`).

What happens: after `Quit()`, teardown waits the window, then kills the process. The script's own
teardown record says 40 of 52 teardowns needed a kill, and that a widened grace did not help
(`verify-office.ps1:46`–`:58`).

Why it matters: 39 test classes × up to 10 s is up to 6.5 min of waiting that the recorded
evidence says buys nothing. The measured open/close cost is **3.2 s** per cycle
(`OfficeFixtureTests.Excel_open_close_five_times_no_orphan`, 16 s for five).

Fix: shorten the graceful window for fixture-owned hosts to ~3 s, keep the kill and the forced-exit
signal, and record the change against the existing measurement. Low risk; the wait is a courtesy,
not a guarantee.

### 6. Per-test workbook setup in classes that could share one (medium)

Where: 19 of 39 test files hold an `OfficeFixture` and most bodies build a workbook and run
`Initialise` before asserting.

What happens: the expensive setup — create workbook, initialise sheet, materialise catalogues, add
rows — repeats per body even when the body only reads.

Why it matters: at ~2–4 s per setup, 142 bodies is 5–9 min of the suite.

Fix: give each class a shared initialised workbook and let mutating bodies opt out. Needs care: a
body that deletes or protects must not see its neighbour's state.

### 7. Fixed settle delays in the clipboard bodies (medium)

Where: `CopyAfterDeleteSpikeIntegrationTests.cs:65` (`ShortDelay = 2 s`), `:66`
(`LongDelay = 10 s`), `:463`.

What happens: the copy/delete bodies wait fixed windows for the clipboard to settle;
`Delete_after_copy…for_an_excel_paste` took 39.5 s in the last run and 39.7 s in the archive.

Why it matters: fixed delays are the one thing the repository's test rules forbid in favour of a
polled condition. They also hide real settling time, so nobody can tell whether 10 s is needed.

Fix: replace with a polled, bounded clipboard-format check (`ClipboardFormatProbe` already exists)
and keep the delay only as a fallback.

### 8. The 25-transfer soak runs inside this gate (low)

Where: `PowerPointTransferSoakIntegrationTests.cs:1`; it is the single most expensive body at
**99.8 s**.

What happens: a soak with its own gate (`scripts/verify-transfer-soak.ps1`) also runs here.

Why it matters: 1.7 min on every Office gate run for a body whose category is meant to be run on
its own schedule.

Fix: confirm the soak trait split is still intended and move the body to the soak category, or
state in the test strategy that the Office gate deliberately includes it.

---

## Reduction options, ranked by saving per unit of risk

| # | Change | Saving (estimate) | Risk |
|---|---|---|---|
| A | Raise `DeadlineSeconds` to 2,400 from the 1,516 s measurement | 0 min; gate becomes able to decide | none |
| B | Remove the 60 s stall (finding 3) | ~4 min, measured at stake | low once instrumented |
| C | Shorten the graceful-exit window to ~3 s (finding 5) | 1–5 min | low |
| D | Share one PowerPoint host per class (finding 4) | 1.5–3 min | medium |
| E | Move the soak to its own category (finding 8) | 1.7 min | low, policy |
| F | Share one initialised workbook per class (finding 6) | 2–4 min | medium |
| G | Poll instead of fixed clipboard delays (finding 7) | ~10–30 s | low |
| H | Second collection for read-only Excel bodies (finding 2) | 3–6 min | high; needs a spike |

Recommended order: **A now** (the gate must be able to report), then **B** because it is measured
waste, then **C** and **G** as small changes in the same file, then **E** as a documented policy
choice, then **F**. **H** only if the gate must fall below ten minutes, and only behind a spike:
two Excel instances, one clipboard, and one `Application` static is exactly the situation the
serial setting exists to prevent.

Do not raise the deadline alone and call the suite fast. A 2,400 s bound on a 25 min suite has
seven minutes of margin; the same bound on a 20 min suite has thirteen, and the difference is
what makes a future red run meaningful.

---

## Next owner

Code mode for A, B, C, G (test and script changes, each with its own measurement). Architect mode
for E and F if the test strategy needs the category split recorded first. The two failing bodies
(`PowerPointTransferSoakIntegrationTests` handle trend, `ConfigRepairIntegrationTests` duplicate-id)
are separate rows and are not addressed here.
