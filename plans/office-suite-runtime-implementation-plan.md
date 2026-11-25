# Office suite runtime — critique of the review, and the implementation plan

**Status:** plan only. Nothing here is implemented, and no product, test, or script file has been
changed. The review under critique is [`plans/office-suite-runtime-review.md`](office-suite-runtime-review.md).

**Verdict on the review:** sound direction, two material errors, and one missing option that is
better than everything it proposed. Fix the baseline before spending a day on savings.

---

## Part 1 — Critique of the review, ranked by harm

### 1. The soak is not in the gate; the review says it is (must correct)

`PowerPointTransferSoakIntegrationTests.cs:81` tags the body `TransferSoak` **only**, and
`scripts/verify-office.ps1:414` selects `--filter Category=OfficeIntegration`. The soak cannot run
through the gate, and `verify-transfer-soak.ps1` is its home — the file header states the split
explicitly ("D4: the routine verify-office.ps1 gate selects Category=OfficeIntegration … this soak
is tagged TransferSoak ONLY").

Consequence: **option E is void and finding 8 is retracted.** It also means the review's headline
number is not the gate's number.

### 2. The baseline is the wrong suite, so every deadline statement is off (must correct)

The 25.2 min figure came from an **unfiltered** run of the assembly: 142 tests, including the
99.8 s soak and every untagged body. The gate's own archived run (`office-20261124-101942282.trx`)
held **105 tests and 16.9 min**.

So the true gap is **780 s against ~1,020 s** — four minutes short, not twelve. That changes the
plan materially: the first action is a filtered measurement, not a 2,400 s constant.

### 3. Finding 5 contradicts its own measurement (must correct)

The review says teardowns "usually" spend the 10 s graceful-exit window, then quotes a measured
**3.2 s** open/close cycle (`OfficeFixtureTests.Excel_open_close_five_times_no_orphan`, 16 s for
five). Both cannot be true: a cycle that routinely waited 10 s could not average 3.2 s.

Consequence: the "1–5 min" saving from shortening the window is unproven. Downgrade it to a
hypothesis with a named measurement (per-class teardown ms), not a proposal with a number.

### 4. Finding 3 has no proven cause and its control run is contaminated (must fix the method)

The four 61.6 s bodies were 0.5 s in the archive and one of them takes 4 s alone. The "alone" run
was executed with a warm host after a suite had already started PowerPoint, so it does not
separate "warm host" from "no neighbour exiting". The review says the cause is unproven, then still
proposes a fix. That order is wrong: instrument first, then fix the named step.

### 5. Missing option — split the gate (the best proposal, absent from the review)

No amount of tweaking makes a 17-minute, serial, real-Office suite a comfortable per-branch gate.
The structural fix is a tagged **smoke subset**: the ~12 bodies that cover the live paths a change
can break (shape render, style, text, refresh idempotence, auto-refresh trigger, staging lifecycle,
row insert), run with a 420 s budget on every branch; the full `Category=OfficeIntegration` set at
release. That is cheaper and safer than every option in the review, and it should be considered
alongside, not after, the deadline change.

### 6. The deadline is a ratchet, and the review repeats the pattern (must fix the policy)

The script's own history is 24 → 26 → 28 → 30 → 35 → 37 → 38 for the old forced-exit ceiling, then
660 → 780 for the deadline. Each raise was correct at the time and each bought one more echo of the
same problem. Raising to 2,400 as a bare constant will do it again.

Consequence: the plan must derive the deadline from a recorded baseline and a stated headroom rule,
and must add the smoke split so the full suite stops being the thing a developer waits for.

### 7. An unverified coverage claim hides in the plan (needs a check)

Some bodies (fixture tests, soak-trend unit facts, clipboard probe facts) may carry no category, in
which case they run in neither the Office gate nor `test-non-office.ps1`. The review neither checks
nor mentions it. Unknown until counted; if true it is a coverage gap that outranks every speed
proposal here.

### 8. Finding 7 assumes the clipboard delay is a settle wait (needs a read first)

`CopyAfterDeleteSpikeIntegrationTests` proves that a copy stays materialisable after a delete. The
fixed 2 s and 10 s windows may model *elapsed user time* rather than await a condition, in which
case polling would remove the thing under test. Read the bodies before proposing the change; the
review did not.

### 9. No measurement or rollback protocol (must add)

Every option needs a before/after figure from the same host, the same filter, and the same command,
plus a one-commit rollback. The review gives ranges with no protocol, which is how a performance
programme produces anecdotes instead of numbers.

**What the review gets right, and should be kept:** the suite is serial by design
(`xunit.runner.json`), so savings must come from test bodies; the 60 s cluster is real and is the
largest single anomaly; PowerPoint and Excel process churn are the repeated fixed costs; and the
gate's job is to decide, which it currently cannot.

---

## Part 2 — Correction to the plan's starting facts

| Item | Review said | Evidence says | Consequence |
|---|---|---|---|
| Gate content | whole assembly, 142 tests | `Category=OfficeIntegration`, 105 tests in the archive | measure filtered first |
| Gate runtime | 25.2 min | 16.9 min (archive), ≈ 1,020 s | deadline gap is ~4 min |
| Soak inside gate | yes | no — `TransferSoak` only, own gate | option E withdrawn |
| Graceful-exit waste | 1–5 min | contradicted by the 3.2 s cycle measurement | hypothesis, not a proposal |
| 60 s stall cause | two candidates | one filtered reproduction needed | instrument before fixing |

---

## Part 3 — Implementation plan

Each phase is one commit, one concern, one measurement. Nothing here deletes or skips a test.

### P0 — Establish the gate baseline (no behaviour change)

- Run `dotnet test tests/GanttCreator.Office.IntegrationTests -c Release --filter "Category=OfficeIntegration" --logger "trx;..."`.
- Record: wall time, sum of durations, top ten bodies, per-class totals, and the count of bodies
  that carry no category at all.
- Acceptance: a filtered trx exists and its numbers replace the review's table.
- Effort: 30 min, of which 17 is the run.

### P1 — Shard the gate so a deadline cannot force everything (replaces the deadline ratchet)

Investigated 2026-11-25; see Part 6 for the numbers. The gate runs one `dotnet test` for the whole
assembly, so the watchdog's `taskkill /T /F` (`verify-office.ps1:442`) takes the tree and every
Office PID in one manifest with it, and no trx survives.

- Split the `Category=OfficeIntegration` set into three shards by a `Shard` trait on each class
  (or by class-name filter), and run each shard as its own `dotnet test` with its own deadline,
  its own owned-PID manifest and its own trx.
- Derive each shard deadline from that shard's P0 measurement plus 120 s; keep the existing
  per-teardown forced-exit rules unchanged and apply them per shard.
- Merge the three trx files for the verdict; a hung shard is reported as one failed shard with
  evidence, not as a blank run.
- Acceptance: a deliberately hung shard (a test double that sleeps) yields two shards' results plus
  one clearly attributed timeout; a normal run yields the same verdict as the single-process run.
- Cost: one extra test-host start per shard, measured at ~2.8 s each (6.8 s wall for a 4 s body).
- Risk: low. It changes the runner's shape, not the tests.
- Effort: 3–4 h plus two gate runs.

### P1a — Deadline from a recorded baseline (retained, but per shard)

- Keep a single number for the whole gate only as the sum of the shard deadlines, and write the P0
  measurements beside it. The rule replaces the 660→780 history: "each shard deadline = measured
  shard baseline + 120 s; re-derive when the shard composition changes".

### P1b — Smoke subset (new; recommended alongside P1)

- Tag ~12 bodies `OfficeSmoke` **in addition to** `OfficeIntegration`, chosen for live coverage of
  the paths a typical change touches: `ShapeRender`, `StyleRender`, `TextRender`,
  `RefreshIdempotence`, `AutoRefreshTrigger`, `StagingLifecycle`, `AddRow`, `InitialiseSheet`.
- Add `-Smoke` to `verify-office.ps1` selecting `Category=OfficeSmoke` with a 420 s deadline.
- Acceptance: smoke run ≤ 7 min wall on this host; full run unchanged.
- Risk: a tag that drifts from the risk surface. Mitigate by reviewing the tag set with the entity
  guide's renderer-equivalence table, and by keeping the full gate at release.

### P2 — Instrument the PowerPoint fixture (feeds P3)

- `tests/GanttCreator.Office.IntegrationTests/PowerPointFixture.cs`: record per-step milliseconds
  for application creation, presentation open, quit, and the teardown wait, emitted through the
  existing test output or an opt-in trace sink. No behaviour change.
- Acceptance: the four suspect bodies, run filtered, report which step owns the 60 s.
- Stop rule: if the stall does not reproduce in a filtered run, close finding 3 as host-specific,
  record the evidence, and skip P3.
- Effort: 1–2 h.

### P3 — Remove the stall (option B), only for the step P2 names

- Candidates, in order of preference once attributed:
  1. Teardown waits, bounded, for the owned process to disappear before returning, so the next
     activation never races a dying server (the file already keeps an owned-PID manifest).
  2. Avoid activating while a force-killed process still appears in the snapshot (the pattern
     `WaitForNoPowerPointAsync` already uses).
  3. If the cost is `Presentations.Open` on a Protected-View file, build the deck in memory.
- Acceptance: each of the four bodies ≤ 10 s in a filtered run; gate wall down by ≥ 3 min; orphan
  count zero; per-teardown forced exits ≤ 1.
- Risk: low once attributed; it touches teardown ordering, which the suite treats as an invariant.
- Effort: 2–4 h including the measurement run.

### P4 — Graceful-exit window (option C, conditional)

- Gate on P2's teardown numbers. If the window is spent, default to 3 s for fixture-owned hosts and
  keep the kill and the forced-exit signal. If it is not spent, record that and close the idea.
- Acceptance: teardown record unchanged (per-teardown ≤ 1); wall measured; no orphan.
- Effort: 30 min plus a run.

### P5 — Clipboard waits (option G, after reading the bodies)

- Read `CopyAfterDeleteSpikeIntegrationTests` first. If the windows model elapsed time, keep them
  and record why in the review. If they await a settle condition, poll `ClipboardFormatProbe` with
  a bounded deadline and keep the same assertion.
- Acceptance: the body still proves the copy stays materialisable after the delete; wall measured.
- Effort: 1–2 h.

### P6 — One PowerPoint host per class (option D)

- Move each PowerPoint test class to a class fixture; keep
  `New_power_point_with_no_presentations_decides_CreatePresentation` isolated because it asserts
  that no POWERPNT is running (`PowerPointTargetDecisionIntegrationTests.cs:36`).
- Acceptance: PowerPoint bodies' wall down ≥ 1 min; no new orphans; per-teardown forced exits ≤ 1;
  the no-host body still starts from a clean snapshot.
- Risk: medium — shared host state. Mitigate by recreating the deck per body.
- Effort: 3–5 h.

### P7 — One initialised workbook per class (option F)

- Architect decision first: a reset protocol and a test-strategy note, because ADR-0020/ADR-0055
  treat retained workbook state as a first-class hazard.
- Acceptance: each class's facts pass in two different orders; wall measured.
- Effort: 1–2 days. Only if P1b is rejected and the full gate must shrink further.

### P8 — Parallel read-only collection (option H)

- Spike only: two Excel instances, one clipboard, `Application` statics. Stop at the first shared
  state, and record the result either way.
- Default answer: no. `xunit.runner.json` is serial because of exactly this.
- Effort: 2–4 h for the spike, no production change unless it proves isolation.

### P9 — Close-out

- `docs/STATUS.md`: one bullet per landed phase with the observed before/after.
- `docs/04-TEST-STRATEGY.md`: the smoke subset, the deadline rule, and the category audit result.
- The two failing bodies (soak handle trend, config-repair duplicate id) stay separate rows.

---

## Part 4 - Sequence, measurement protocol, rollback

> **Superseded by Part 8 for the sequence.** P2, P3 and P3b were withdrawn on measurement, P4 landed,
> P12 was withdrawn, and the smoke subset (P1b) was never adopted into a roadmap row. Read Part 8 for
> what actually happened; keep this section for the measurement protocol and the rollback rule.

**Sequence:** P0 → P1 (shards) → P1b (decision) → P2 → P3 → P4 and P5 in either order → P6 → stop
or P7/P8 by decision. P1 lands before any speed work so the gate exists to measure the rest and so
a slow phase cannot take the whole run down with it.

**Measurement protocol:** same host, same filter, same command; record wall, sum, and the top ten
bodies; for any delta under 10% run three times and report the median; never quote a filtered
number and an unfiltered number in the same sentence.

**Rollback:** one commit per phase, each revertible alone; no test deleted, skipped, or loosened; a
phase that does not move the measured wall time is reverted rather than kept "because it is
tidier".

**Measured outcome instead of the predicted one (2026-11-25):** the live gate is **848 s (14 m 8 s)**
and was reached by **P4 alone**; P3, P3b, P5, P6 and P12 contributed nothing measurable and were
withdrawn. The per-branch smoke run (P1b) does not exist.

**The residual is measured — and the first reading of it was wrong (2026-11-25).** The fixture
records setup and teardown beside the body, and the gate reconciles them. One live run: the test
process ran **701.1 s** — setup **56.7 s** (of which **50.4 s** is the host launch), body
**197.5 s**, teardown **206.3 s** (of which **148.0 s** is the graceful exit poll). The 68
**recorded** fixtures own **460.5 s**, and the first version of the phase line called the remaining
**240.6 s** "outside every fixture — the runner's per-test turnaround".

That reading was wrong, and two measurements disproved it:

1. **The host-free control:** the same assembly's 39 Office facts that launch nothing ran in
   **2,655 ms wall against 210 ms of test time** — the runner's own cost is ~2.5 s fixed, not 2.3 s
   per test.
2. **The trx sum:** the full run's trx reported **106 results summing 650.2 s** against a 661.1 s
   process, so the runner's whole overhead is **10.9 s (1.6%)**.

The 240.6 s was therefore not runner time: it was **fixture time in tests that write no teardown
record** — the 15 bodies that suppress the leak signal because they legitimately force a kill, and
the ~23 PowerPoint bodies, whose fixture writes to no Office record. The gate now prints three
figures from three sources (recorded fixtures, the runner's reported test time, the process clock)
and the arithmetic between them, so the mistake cannot be repeated silently. The corrected split:
**650.2 s of test time** (477.8 s recorded fixtures, **172.4 s in the 38 unrecorded bodies**) and
**10.9 s of runner overhead**.

The lever this exposes is not the test host. It is the **graceful exit poll: 144.9 s of the 650 s**,
because two-thirds of teardowns are killed and each pays the full 3 s window, while a clean exit is
observed at 0 ms. A 1 s window would save ~92 s on the full run; it changes a calibrated safety
bound and needs its own measured run before it is adopted, so it is recorded here as the next
candidate rather than taken.

**What the plan predicted:** a ~17-minute gate becomes ~13–14 minutes from P3 + P4 + P5 + P6, and
the per-branch gate becomes a ≤ 7-minute smoke run from P1b. If P3's stall is host-specific, only P1
and P1b remain, and the honest statement is that the full live suite is a release gate measured in
tens of minutes.

---

## Part 5 - What not to do, and why

> **Partly superseded by Part 8.** The teardown-ordering rule below was conditioned on P2 naming the
> stall; P2 was withdrawn (the stall is positional), so that condition no longer binds. The rest of
> this section still holds.

- **Do not raise the deadline and stop.** It restores a verdict without restoring confidence: a
  2,400 s bound on a 1,020 s suite hides a 15-minute regression.
- **Do not flip the xunit parallelism flags.** The suite is serial because of shared clipboard, one
  Excel per class, and `Application` statics; P8 is a spike, not a change.
- **Do not retag a slow body out of the gate to make it faster.** The soak is already out by
  design; anything else moved out is coverage lost, and the plan's P1b makes the same saving
  transparently by adding a tag, not removing one.
- **Do not change teardown ordering before P2 names the stall.** The teardown is where the suite's
  ownership guarantees live.

---

## Part 6 — Investigation: what is running at the deadline, and can it close cleanly (2026-11-25)

### What is in flight when the watchdog fires

Reconstructed from the gate's own 105-test archive (`office-20261124-101942282.trx`, total
1,013.5 s): at 780 s the body in flight was
`TableGridBorderIntegrationTests.Every_boundary_survives_the_reports_fill_followed_by_an_add`
(775 s → 789 s). **Twelve bodies and 233 s remained.** Immediately before it ran the clipboard
trio — 15.9 s + **39.7 s** + 15.5 s = 71 s — then protection arming (23.6 s) and paste-special
(14.6 s).

So the deadline does not fall on a random body: it falls on the slow tail, because class order is
effectively alphabetical and the expensive classes cluster late.

### Why the force-all is structural, not a scheduling mistake

The gate runs **one** `dotnet test` for the whole assembly. On timeout the script kills the launcher
tree (`verify-office.ps1:442`) and then every PID in the fixture manifest (`:283`). The trx is
written at the end of the run, so a run killed mid-flight has no results to save. Given that shape,
killing everything is the only thing the script can do.

### Can the bodies close cleanly? Three mechanisms, measured where possible

| Mechanism | What it fixes | Effect |
|---|---|---|
| Per-shard process boundaries (P1) | A hang costs one shard, not the run; other trx files survive | ~2.8 s overhead per shard (measured: 6.8 s wall for a 4 s body) |
| Shorter exit poll (P4, revised) | The 10 s graceful wait before the kill | up to ~7 s saved per class that hits the race |
| Heavy-first ordering (P3b, new) | A budget hit lands on cheap work, not a 40 s body | 0 s on a normal run; bounds the loss when it happens |

**The honest limit.** The fixture already closes every workbook, releases every proxy, collects,
and polls for exit; `OfficeFixture.ComScope` is in **113** places. The script's own record says the
survivor is an **Excel shutdown race**, not an unreleased proxy — even a body that releases nothing
can hit it, and the surviving process is idle and responding. So "close cleanly" cannot mean
"guarantee Excel exits". It means: **no forced kill on the happy path, at most one kill per hung
shard, and always a result file.**

### Can the work at the deadline be reduced? Yes, and the list is short

- Clipboard trio: 71 s, of which two bodies use fixed 2 s and 10 s waits (P5).
- Table-border pair: ~63 s in the 2026-11-25 run (15.9 s in the archive), per-class workbook setup
  (P6/F).
- Protection arming 23.6 s, paste-special 14.6 s, initialise class 1.7 min over 8 bodies (F).
- The 60 s PowerPoint cluster when it appears (P2/P3).

Shaving 233 s from the front would make 780 s viable again **by coincidence of composition** — add
two bodies and it breaks. The shard split is the mechanism that makes the gate independent of the
total; P1b (smoke) is the mechanism that makes a per-branch run short. Reducing the tail is what
keeps both inside their budgets.

### P3b — Heavy-first ordering (new, mitigation only)

- Add a `Weight` trait to bodies over ~15 s and an `ITestCaseOrderer` that runs them first.
- Acceptance: the ordering is deterministic across two runs; no behavioural change; the gate still
  passes.
- Risk: low, but it reorders execution, so any accidental order dependency in the fixture surfaces
  here — which is itself worth knowing before P7.
- Do this only after P1; before P1 it just moves the deck chairs.

---

## Part 7 — Leak elimination, so the ratchet can go (investigated 2026-11-25)

The user's target: clean up the leaky tests and prevent the leaks, rather than live with a forced
-kill ratchet. The investigation found the cause, and it is not what the ratchet's name suggests.

### The experiment

Three arms against a real Excel, run directly on 2026-11-25 with no test code involved:

| Arm | RegisterXLL | Unregister | Excel process after `Quit()` + 10 s |
|---|---|---|---|
| A: plain workbook, no add-in | no | — | **gone** |
| B: register the packed XLL, then quit | yes | — | **still running** |
| C: register, then `Application.UnregisterXLL` | yes | not available | **still running** |

Arm C failed at the call: the Excel `Application` COM object has **no `UnregisterXLL` member**
(`Method invocation failed … does not contain a method named 'UnregisterXLL'`). `RegisterXLL`
exists; its counterpart does not.

**Conclusion:** Excel does not exit because the packed Excel-DNA add-in is loaded, not because a
test leaked an RCW. A workbook holding shapes is a red herring; a bare workbook with the XLL
registered shows the same behaviour. This is a **host lifetime constraint**, and it is why the
XLL-bearing bodies kill on purpose.

### What the current signal actually measures

- `RegisterXll` is called from **21 test files** covering roughly 70 bodies.
- `SuppressLeakSignal` is set in **19 places across 12 files** — all of them XLL-bearing bodies. The
  others still report a forced kill, so the ratchet's number is mostly "one kill per XLL body",
  which the per-teardown ceiling of 1 accepts by construction.
- Bodies with no XLL (arm A) exit cleanly today, so the mechanism that *would* catch a real leak
  works — it is simply drowned out by the XLL population.

So the ratchet is not protecting against leaks at all; it is counting a design constraint. That is
why raising it never converged.

### The fix, in two rules and one fixture

**Rule 1 — no XLL, no kill.** Every body that does not register the pack must reach zero forced
kills. A kill there is a real leak or a real hang and must fail. `OfficeFixtureTests` already
proves five clean open/close cycles are possible.

**Rule 2 — XLL bodies do not own an Excel process per test.** Replace per-test
`new OfficeFixture()` + `Quit()` in XLL classes with one **session fixture per shard/class** that
loads the pack once, creates a fresh workbook per body, and exits once at the end. One deliberate
exit per shard, recorded once, instead of N kills and N × 10 s waits.

That also removes the reason `SuppressLeakSignal` exists: there is no per-test teardown to
suppress. The flag can then be deleted, and its 19 uses with it.

**Rule 3 — proof of no leak, not absence of a kill.** Two additions to the fixture:
- Capture the value `Marshal.ReleaseComObject(_excel)` **returns** at teardown (it is currently
  discarded). `> 0` means a live RCW still holds the application: a real leak, named per test.
  `0` with the process alive means the XLL (arm B) — not a leak.
- Record handle and working-set counts at shard start and end and assert a flat trend with the
  existing [`SoakTrendTests`](../tests/GanttCreator.Office.IntegrationTests/SoakTrendTests.cs)
  logic. That catches growth that never keeps Excel alive, which the process-exit signal cannot see.

### What the gate becomes

| Signal | Today | After |
|---|---|---|
| Forced exits per teardown | ceiling 1, total reported | **0** for non-XLL; **exactly 1 per XLL shard, by design** |
| Suppressed bodies | 19, uncounted | flag deleted; XLL shards are the only exclusion, counted and capped |
| Leak proof | process exit only | RCW count at 0 **and** flat handle/working-set trend |
| Ratchet | raised repeatedly | removed; the rule is absolute, not a ceiling |

The `verify-office.ps1` per-teardown ceiling and its long explanatory comment then reduce to one
sentence per class, and the "forced-exit total tracks test count" problem disappears because there
is no longer a forced exit to count.

### Work items

- **P10 — classify (1 h).** List every teardown as `Xll` or `Plain` from its `RegisterXll` calls;
  record the split in the report. Expected: ~70 XLL bodies, ~30 plain.
- **P11 — plain bodies to zero (1–2 d).** Run the plain subset with the forced-kill log; for each
  body that kills, capture the RCW count and fix the holder (a `ComScope` gap or an undisposed
  adapter). Acceptance: plain subset reports 0 forced exits across three consecutive runs.
- **P12 — XLL session fixture (2–3 d).** One Excel per XLL class, pack loaded once, workbook per
  body, single exit at the end. Acceptance: one kill per class at most, XLL classes wall time down
  (no per-body start/quit), no cross-body state bleed (run each class's bodies in two orders), and
  `SuppressLeakSignal` deleted.
- **P13 — gate rewrite (0.5 d).** `verify-office.ps1`: 0-forced-exit rule for plain teardowns,
  one-per-shard for XLL, RCW-count and handle-trend assertions, suppression cap removed.
  Acceptance: a deliberately leaked RCW in a scratch test fails the gate with the test named.

### Honest limits

- Excel will still not exit while the pack is loaded; P12 works *with* that, it does not defeat it.
- Excel-DNA may hold its own process for a moment after the last workbook closes; the single
  end-of-shard exit absorbs that once instead of once per test.
- If Excel-DNA ever supports a clean unload, P12 collapses further, but no current API does — arm C
  is the evidence.

---

## Part 8 — Executed 2026-11-25 (observed, not planned)

### P0 — the gate's own baseline

`dotnet test --filter Category=OfficeIntegration` on the clean tree: **105 tests, 105 passed, sum
19.3 min, wall ≈ 21.5 min**. Top bodies: one PowerPoint body at **62 s** (the stall appeared once,
not four times, so it is intermittent), the clipboard copy/delete at **39.5 s**, slide-fit probe
19.4 s, paste-special 18.5 s. Against 780 s, the gate is about **8.5 minutes short**.

The same run with `GANTTCREATOR_FORCED_KILLS_PATH` set sampled 32 teardowns before the log was read:
**25 forced a kill, 7 were clean, `Quit` returned in all 32**. That is the same ~77% the script's
comment records, and it is the population P10–P13 exist to separate.

### P1 — landed as infrastructure, sharding default OFF

Two measurements changed the design:

1. `Category=OfficeIntegration&(FullyQualifiedName~A|...|~N)` — VSTest has no grouping, so the
   expression did not mean what it read. It ran **82 of 106** tests and every shard exited 0:
   **a false PASS**.
2. A bare OR list without the category — `~` is a substring match, so the lists overlap and the
   filters also match **15 tests that carry no category** (117 listed vs 102 discovered).

Landed instead: discovery of the gate's own test list from `--list-tests` (anchored parse — a loose
match captured the `(.NETCoreApp,Version=v10.0)` banner as a class), per-shard trx/manifest/kill-log
plumbing, per-shard deadlines derived from `-DeadlineSeconds`, and — the important part — a
**coverage check** that compares every executed test name in the shards' trx against the discovered
list and fails with `exit 6` and the first missing name. `ShardCount` defaults to **1**, which keeps
the exact category filter and the old single pass, now with a proof that it ran everything.

Multi-shard selection is deferred until it is expressed as an exact `Shard=N` trait; the two
failures above are recorded in the script beside the parameter so nobody re-tries the OR form.

### P10 — the teardown record now says whether the XLL was loaded

`OfficeFixture.XllRegistered` is set by `RegisterXll` and written into every teardown record as
`"Xll":true|false`. The contract test drives both values (`Teardown_records_the_quit_outcome...`,
positive control included). This is what makes rule 1 of Part 7 enforceable: a forced kill with
`Xll=false` is a leak or a hang; with `Xll=true` it is the host-lifetime constraint.

### Still owed

**Acceptance run, observed:** `verify-office.ps1 -DeadlineSeconds 2400` on `bfbe47f3`'s tree
passed, with the new line

```
coverage: all 102 discovered test(s) ran in a shard.
Forced-exit signal: 47 forced exit(s) across 68 recorded teardown(s) (69% escalated ...)
```

**The Xll flag already pays for itself.** Split of those 68 teardowns:

| Xll | Records | Forced kills |
|---|---|---|
| `false` | 49 | **28** |
| `true` | 19 | **19** |

Every pack-loaded teardown forced a kill (19/19), which is the host constraint, as predicted. But
**28 forced kills happened with no pack loaded** — `AddRow` (six bodies), `GanttTableReader` (five),
`OutlineGroup` (four), `CopyEditable` (two), `StagingLifecycle` (two), `ColumnVisibility` (two), and
others. Arm A of the probe showed a plain workbook exits, so these are genuine leaks or hangs, and
they are now a named list rather than part of an undifferentiated ratchet. That list is P11's work.

### P11 result — the leaks do not exist (2026-11-25, negative)

`ReleaseComObject(Application)` now returns the remaining reference count into every teardown
record (`"Rcw"`). One full gate run later:

| Rcw | Forced kill | Records |
|---|---|---|
| `0` | 0 (clean) | 22 |
| `0` | 1 | **46** |

**Every teardown released to zero.** The fixture was the last managed holder in all 68 teardowns,
and Excel still survived in 46 of them. There is **no RCW leak in the Office gate.** That kills the
premise of P11 as written ("clean up the leaky tests") and of Part 7's rule 1 ("no XLL, no kill"):
a forced kill is not evidence of a leaked reference, and no amount of `ComScope` adoption will
remove it. My earlier classification by `Xll` was correct as a *correlate* (19/19 killed with the
pack loaded, 28/49 without) but wrong as a *cause*: plain bodies with `Rcw=0` kill too, so touching a
real workbook is enough.

**Revised rules:**

1. **`Rcw == 0` on every teardown is the leak gate.** It is measurable, it is already green, and a
   non-zero value names a genuine bug. This replaces "forced exits must be zero".
2. **Forced exits are an Excel residency metric.** Keep the existing per-teardown ceiling of 1 and
   report the total and the trend; do not gate a zero that the host controls.
3. **P12 was re-scoped to a pure runtime change and has now been measured and withdrawn** (below):
   the host launch it would save is 743 ms, not the 2.8 s an earlier experiment suggested.
4. **P13 became smaller and is landed** (below): gate `Rcw`, keep the ceiling, delete the
   "escalating fraction" narrative that reads the kill as a leak, and keep the Xll/Rcw fields for
   diagnosis.

### P13 — landed (2026-11-25)

`verify-office.ps1` parses `Rcw` from every teardown record and prints a reference-count line beside
the forced-exit line. The forced-exit line no longer computes an "escalating percent": it reports the
killed fraction and states that the kill is residency, citing the 68-teardown measurement in place.
Three verdicts, in order:

| Signal | Rule | Exit |
|---|---|---|
| Application reference count, not zero | a leaked proxy — a real defect | `7` |
| Application reference count, never reported | unmeasured is not zero | `5` |
| Forced exits above the per-teardown ceiling | more than one owned process | `4` |

`verify-office.Tests.ps1` gained four facts (42/42 PASS) and its stub a `VERIFY_OFFICE_STUB_RCW` knob
(`0` default, `1` leak, `absent` unmeasured). The four that matter: a teardown that needed one kill
and released to zero still **passes** (the corrected meaning), a leaked count exits 7 and names the
test, an unreported count exits 5, and a positive control proves the three source assertions fail on
a source that has none of the new idioms. The data behind the gate is the measured P11 run (68
teardowns, all zero); the new verdict itself has harness evidence only until the next live run.

### P4 — the exit poll is 3 s, and it is the biggest measured lever (landed 2026-11-25)

The same 68 teardown records answer P4 without another run:

| Record set | Graceful wait |
|---|---|
| 22 clean teardowns | **0 ms** each (maximum 0) |
| 46 killed teardowns | 10,056 ms minimum, **464,040 ms** total |

Test-body work in the same run was 247.3 s, so **two-thirds of the suite's wall time was a window
no survivor could ever satisfy.** `DefaultGracefulExitSeconds` is 3, the comment carries this table,
and `The_production_graceful_bound_is_three_seconds_and_the_diagnostic_override_wins` pins the bound
and proves the diagnostic env override still wins. This is cheaper and safer than P12: it changes one
constant, not the fixture's isolation model.

**Confirmed live the same day.** `verify-office.ps1 -DeadlineSeconds 2400` **PASS** on Excel
16.0.20625.20108 x64: 103 of 103 discovered tests ran, 68 teardowns, 47 forced exits, every
`Quit()` returned, `Reference-count signal: 68 of 68 ... and 0 left one behind`, and the graceful
wait totalled **144,920 ms (maximum 3,112 ms)** against 464,040 ms at the 10 s bound — **319 s
saved**, the prediction to within a second. End to end the run took 14 m 8 s including the Release
build, against 25 m 13 s for the 10 s tree, while test-body work *rose* 247.3 s -> 319.9 s, so the
difference is the wait and not a lighter suite.

### P2 — the "~60 s stall" is positional, not a hang (2026-11-25, negative)

The 66 s body in the live run was
`SendToPowerPointIntegrationTests.The_zero_deck_create_arm_pastes_and_ungroups_on_a_fixture_owned_host`,
the **66th teardown of 68**. Run alone through the same assembly and filter it takes **10 s** (its
sibling 14 s; both together 26.6 s wall). So there is no fixed stall inside the body, and nothing to
shrink: the same work is six times slower at the end of the suite than on its own.

The position effect is in the records and is not subtle:

| Teardown positions | Mean body time |
|---|---|
| 1-34 | 2.7 s |
| 35-68 | **6.7 s** |

**It is not leaked hosts.** Immediately after the local runs this session, `Get-Process EXCEL,POWERPNT`
reported none alive, and the gate's own sweep found no stray, so the degradation is inside the test
host or below it, not in accumulating Office processes.

**A fresh host shows no drift at all.** The four heavy classes run alone — `SendToPowerPoint`,
`CopyAfterDeleteSpike`, `PowerPointSlideFit` and `PowerPointSlideFitRouteProbe`, 7 tests — reproduce
their in-suite times almost exactly (22, 10, 19, 39, 13, 10, 13 s), including the zero-deck body at
**10 s**. So the 66 s reading needs a long-lived host, not a particular neighbour.

**The 2-shard rehearsal was not a measurement (recorded as such).** Splitting the assembly's 44
classes into two ~6-minute runs produced half A 92 tests in 5 m 52 s and half B **51 tests in
6 m 36 s with one failure**. That is 143 tests, not the gate's 103: a bare `FullyQualifiedName~Class`
OR list matches by substring and also ignores the category, so it pulled in the 36 non-live bodies
and the `Category=TransferSoak` body the routine gate deliberately excludes (confirmed in the
`--list-tests` output). The failing body was identified afterwards by re-running it alone: it is the
soak itself, and it fails reproducibly (below). The rehearsal therefore reproduced P1's own filtering
lesson rather than answering the sharding question.

**Open defect, reproducible: the soak's handle-trend tolerance is defeated by its own cold start.**
`dotnet test ... --filter "FullyQualifiedName~PowerPointTransferSoakIntegrationTests"` alone:
**Failed in 1 m 44 s**, `handlesMin=1892 handlesMax=1956 delta=64 > tolerance=50` over 15 attach-arm
samples. The series names the cause — the attach samples read **1892** (first), then 1951, 1956, and
settle at **1936-1941** for the remaining twelve; `excelHandles` climbs **1443 -> 1521** across the
first six attach iterations and then flattens. So the max-minus-min is a **warm-up step, not a rising
leak**: the steady region varies by 36, inside the tolerance.

**Repaired on the owner's ruling the same day.** `SoakTrend.Handles` excludes the first two attach
samples (`SoakTrend.HandleWarmUpSamples`) and gates max − min over the steady region; the tolerance
stays **50**, a steady region shorter than three samples still fails closed, and the failure message
names the exclusion. Four host-free facts pin it: the recorded series (`Baseline=1920 End=1956
Delta=36`), a climb that begins inside the warm-up (`Delta=90`, still fails), a two-sample steady
region, and the growing-series fact now reading `Delta=6` with its verdict unchanged. **Live: the soak
body passed in 1 m 46 s** with `handlesMin=1952 handlesMax=1972 handlesDelta=20 handlesTolerance=50
orphans=0` — and the warm-up step on that run was only 12, which is why the old whole-series rule
failed intermittently rather than always.

**Consequences.** P2 is withdrawn as written — there is no 60 s timeout to remove. The condition for
testing the sharding prediction was a filter that cannot over-match: each shard term self-contained as
`Category=OfficeIntegration&FullyQualifiedName~Class`, plus a coverage assertion that also fails on
*extra* tests. **Both landed, and P3b is now answered — negative.** With those in place the gate ran
two shards twice, both **PASS** with the both-directions coverage proof (`all 104 discovered test(s)
ran in a shard, and nothing else ran`):

| Run | Wall | Shards |
|---|---|---|
| 1 shard (2026-11-25) | 848 s | 778 s of test time |
| 2 shards, first | 767 s | — |
| 2 shards, repeat | 863 s | 375.6 s + 390.6 s = **766 s** |

The shard sum equals the single pass, and the wall time is inside run-to-run noise in both
directions. **Splitting the suite does not recover the position effect**, so `ShardCount` stays 1 and
P3b is withdrawn. The durable gains of the work are the filter and the two verdicts: they found a real
discovery defect (every `[Theory]` was invisible to the gate's test list, because the host lists a
theory only as one row per data set) and a count mismatch that reported 104 discovered against 106
covered.

### P5/P6 — PowerPoint hosts are not the cost (2026-11-25, negative)

The Excel wait was worth 319 s, so the obvious next question was whether its PowerPoint twin behaved
the same way. `PowerPointFixture` now records the wait it actually spends, in
`LastGracefulExitElapsedMs`, and the two send-to-PowerPoint bodies print it: **1 ms** and **0 ms**.
PowerPoint is gone before the first poll, so its 10 s bound costs nothing, and shortening it would
change nothing. **P6 (per-class PowerPoint host) therefore has no teardown-wait motive either** — the
10 s zero-deck body is real work (Excel start, stage, copy, paste, ungroup), not a wait.

**P5 is answered the same way: there is no clipboard wait to remove.** `WindowsClipboardFormatReader`
has no poll, retry or deadline — it reads a format once. The only delays anywhere in the clipboard
path are `CopyAfterDeleteSpikeIntegrationTests`' two **measurement windows** (2 s and 10 s), and they
are the experiment's parameters: the spike exists to observe whether Excel's delayed rendering
survives deleting the source shapes, so the window is the observation, not a synchronisation cost.
That closes the whole wait family with numbers — Excel 464 s (fixed), PowerPoint 0–1 ms, clipboard
12 s of deliberate measurement — and leaves body work plus host start-up as the suite's real cost.

### P12 — one host per class is worth 35 s, and is withdrawn (2026-11-25)

The remaining candidate was one Excel host per test **class** instead of per body. `OfficeFixture` now
records the launch it actually pays (`HostMs`, the `new Application()` call), and 16 real bodies across
`AddRow`, `OutlineGroup` and `GanttTableReader` measured **mean 743 ms, max 1,158 ms, 11,881 ms total**
— not the 2.8 s an earlier experiment recorded, which must have included the workbook work that
follows the launch.

The arithmetic follows directly: 68 bodies across ~21 workbook classes means ~47 launches avoided,
i.e. about **35 s per run** — 4% of a 14-minute suite. Against that, a per-class host replaces
per-body process isolation with per-class processes, so every body after the first inherits the
previous body's leftover sheets, defined names and add-in state; that is exactly what the suite's
fresh-workbook-per-body design avoids, and no test asserts it can tolerate the difference. **35 s does
not buy that risk**, so P12 is withdrawn with the measurement in the record rather than implemented.
The `HostMs` field stays: it makes the next such question a one-run answer instead of an estimate.

### P1b — landed and measured live (2026-11-25)

Eight classes carry `OfficeSmoke` at class level in addition to `OfficeIntegration`:
ShapeRender, StyleRender, TextRender, RefreshIdempotence, AutoRefreshTrigger, StagingLifecycle,
AddRow and InitialiseSheet. Class-level, so a body added later joins without a second edit; the
tag set was reviewed against the entity guide's renderer-equivalence table. The subset is
**24 cases across 22 methods** (one method is a three-row Theory), verified with
`--list-tests --filter Category=OfficeSmoke` before the gate was touched.

`verify-office.ps1 -Smoke` derives the category **once** and reads it at discovery, the
single-shard filter and the per-shard conjunction terms. That single variable is the design: a
literal left at discovery beside a smoke run filter would make the coverage proof report every
non-smoke test as missing. The bound defaults to **420 s**, and an explicit `-DeadlineSeconds`
still wins.

**Observed, first live run:** `Deadline: 420 s` adopted, **24 of 24 passed**, shard 0 exit 0
after **165.4 s** (`coverage: all 22 discovered test(s) ran in a shard, and nothing else ran.`),
`Reference-count signal: 18 of 18 ... and 0 left one behind`. The corrected phase split on the rerun: **160.3 s process, 151.4 s of runner-reported test time, 8.9 s of runner overhead**, and of that test time **106.8 s in the 18 recorded fixtures, 44.6 s in the six suppressed render bodies that write no record** — the figure the first version of the line had mislabelled as runner cost. Wall end to end ~3.5 min including the build, half the 7-minute
acceptance. The full gate re-ran unchanged the same evening: **104 tests, 661.1 s, all verdicts
the same**. P7 stays gated: its measured net ceiling is ~30–60 s (see
[`p1b-p7-pre-implementation-review.md`](p1b-p7-pre-implementation-review.md)), so it starts only if
this smoke run is judged insufficient.

### P4b — the graceful bound tested downwards: rejected on its own measurement (2026-11-25)

The corrected phase data made the exit poll the largest named block (144.9 s of a 650 s test
time, two-thirds of teardowns killed, clean exits observed at 0 ms), so the bound was taken from
3 s to 1 s and measured. **It is rejected, and the bound is back at 3 s.**

| Signal | 3 s bound | 1 s bound |
|---|---|---|
| Graceful poll, total | 144.9 s | **68.7 s** |
| Forced exits | 47 of 68 | **62 of 68** |
| Shard wall | 661.2 s | **731.2 s** |

The poll fell as predicted and the run still got **slower**, because fifteen processes that had
been exiting on their own within three seconds were now killed: a kill costs more than the wait it
replaces, and the gate also loses fifteen clean teardowns from its strongest signal. The lesson is
the one the P4 round already implied and this round proves: the window's job is to let a real exit
happen, so it has a floor, and 3 s is it. The 1 s experiment is recorded in the fixture's comment
so nobody repeats it.

### P7 — the state-delta audit, and the ceiling it puts on one workbook per class (2026-11-25)

P7 (option F in [`p1b-p7-pre-implementation-review.md`](p1b-p7-pre-implementation-review.md)) was
approved "behind its prerequisites", and the first prerequisite is this audit: what state does a body
leave behind that the next body in the same class would inherit? Counting the tree answers the
structural half of it, and the answer re-prices the item.

**The fixture does not own a workbook.** Every live class constructs its own
[`OfficeFixture`](../tests/GanttCreator.Office.IntegrationTests/OfficeFixture.cs) inside the test
body and then calls `CreateWorkbook()` itself:

| Measure | Count |
|---|---|
| Live classes that construct a fixture in the body | 22 |
| `new OfficeFixture(...)` sites | 59 (46 outside `OfficeFixtureTests`) |
| `CreateWorkbook()` sites | 47 (43 outside `OfficeFixtureTests`) |
| `RegisterXll(...)` sites | 23 |
| Classes with more than one body, and the cases they hold | 10 classes, 35 of the 46 live cases |

So P7 is not "move one assignment to a class fixture". The workbook is created by the body and
tracked by the fixture for teardown, so a per-class workbook needs a **reset-to-pristine API** plus
rewiring 46 fixture constructions, 43 workbook creations and 23 pack registrations.

**The state deltas a reset would have to clear**, enumerated from what the classes' own bodies do:

1. live and staging shapes, their groups, and their ownership tags;
2. the `tblGanttData` list object — rows, hidden columns, hidden rows and outline levels;
3. the `_GanttCreatorConfig` VeryHidden sheet and its schema version;
4. the sheet-scoped defined name (plot anchor);
5. sheet protection, which the add-in's own automatic refresh arms (ADR-0055);
6. row heights and column widths, written by the normaliser and column-visibility bodies;
7. validation highlight fills, written by the `ValidationReporter` bodies;
8. workbook close/save state, which some bodies change mid-test;
9. application state — `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, status bar, selection —
   already restored per ADR-0020, so this one is a delta the current discipline covers;
10. clipboard contents, from the copy bodies;
11. the loaded pack and any scheduler state, which under P7 would span a whole class instead of one
    body.

**One delta of eleven has live evidence.** `InitialiseSheetIntegrationTests` already has a
`NonPristineArtefact` theory — `Shape`, `Comment`, `ThreadedComment` — that pollutes a sheet and
asserts `Initialise` repairs it. That is the seed of the reset protocol and proof for delta 1 only;
the other ten would be new.

**The ceiling, from the measurements already recorded.** Only the 10 multi-body classes can save
anything, and the saving is one launch and one teardown per body avoided:

| Lever | Arithmetic | Value |
|---|---|---|
| Host launches avoided | 35 bodies − 10 classes = 25 × 743 ms (P12 measurement, max 1,158 ms) | **18.6 s** |
| Graceful polls avoided | 25 teardowns ≤ 3 s, minus the 10 per-class teardowns that still pay one | **≤ 45 s** |
| Forced kills avoided | the same 25, at the kill cost the P4b run measured | **≤ 20 s** |
| Ceiling | | **≈ 65–85 s of a 661 s run, ~10%** |

**Recommendation: do not run P7 as a programme.** Per-body process isolation is the property every
flake in this suite has so far been explained by — the AddRow protection race, the soak handle
warm-up, the positional drift — and it is what delta 11 puts at risk for a ~10% saving. Three
dispositions, for the owner:

- **F1 — narrow it.** Apply per-class hosting only to the five classes with three or more bodies
  whose bodies already re-initialise from a pristine sheet at their start (AddRow 7, InitialiseSheet
  6, GanttTableReader 5, OutlineGroup 4, ValidationReporter 3): 25 bodies in 5 classes, ~40 s, and
  still the full reset protocol.
- **F2 — stop here.** P7 stays a documented option with its ceiling measured rather than estimated,
  which is what P12 became.
- **F3 — spike it.** Rewire AddRow alone behind the reset API and measure the class in the gate: one
  class, ~10 minutes of gate time, and the number becomes a measurement instead of arithmetic.

Honest limits: the launch figure is measured on this host, the poll and kill figures are computed
from teardown records rather than from a P7 run, and delta 11 — a pack loaded across a class's
successive bodies — has no measurement at all.

### Review close-out (2026-11-25)

A code review of this plan against the tree raised nine findings. Closed in code the same day:

- **The soak gate now gates the leak.** `verify-transfer-soak.ps1` parses `Rcw`, exits 7 on a
  non-zero count, and exits 5 when no teardown reported one — the same signal as the routine gate,
  which matters here because the soak is the only gate that performs 25 transfers. Live proof:
  `Reference-count signal: 1 of 1 teardown(s) ... and 0 left one behind`, PASS.
- **Discovery failure is no longer a pass.** `verify-office.ps1` exits 6 when the host lists no
  tests, with `-AllowUnprovenDiscovery` as the documented opt-out for the Pester harness.
- **The host-launch figure is reported**, so the P12-style question no longer needs the raw JSONL.
- **The residual is measured, and the measurement corrected itself.** The fixture records setup and
  teardown beside the body, and the gate now prints three figures from three sources: the recorded
  fixtures, the runner's reported test time (every trx result), and the process clock. The first
  version summed only the records and called the difference runner time — 240.6 s of a 701.1 s
  process. The host-free control (39 facts in 2,655 ms wall against 210 ms of test time) and the trx
  sum (650.2 s of 661.1 s) disproved that: the runner costs **10.9 s**, and **172.4 s is fixture
  time in 38 bodies that write no record**. The next candidate lever is the **144.9 s graceful-exit
  poll**, not the test host.
- **The leak claim is scoped** to the `Application` proxy in `KNOWN-LIMITATIONS` L19.

### Owner decisions, approved 2026-11-25

All three open decisions are approved. Each keeps the gate its own plan text already stated, so
"approved" means the work may start, not that its prerequisites are waived.

1. **P1b — the smoke subset is approved, and it is first.** Tag the eight classes named in P1b with
   `OfficeSmoke` **in addition to** `OfficeIntegration`, add `-Smoke` to `verify-office.ps1` selecting
   `Category=OfficeSmoke` with a 420 s deadline, and keep the full gate unchanged. Acceptance stands:
   smoke wall ≤ 7 min on this host, full run unchanged. The tag set is reviewed against the entity
   guide's renderer-equivalence table before it lands, because a tag that drifts from the risk surface
   is a false sense of safety rather than a saving.
2. **P7 — one initialised workbook per class is approved, behind its own prerequisites.** It does not
   start until the reset protocol and the test-strategy note exist, because ADR-0020/ADR-0055 treat
   retained workbook state as a first-class hazard; and its acceptance stands: every migrated class's
   facts pass in **two different orders**. It is sequenced after P1b, one class at a time, with the
   reorder proof in the same commit as each class. If P1b meets the per-branch target, P7 is not
   needed and this decision is re-read, not re-made.
3. **The runner overhead is approved as a work item whose first step is attribution, not a migration.**
   The phase signal already isolates it (240.6 s of a 701.1 s process, ~2.3 s per test). Step 1 is a
   control measurement that separates per-test runner cost from per-test fixture cost — the same
   assembly's host-free facts, whose trx durations are already known to sum to ~1 s — and only then a
   choice between a test-host/platform change and a runner-configuration change. No product code
   changes until that measurement names the mechanism.

Still open with owners, outside these three: the `ConfigRepair` `IdentityDamage` intermittent (needs a
reproduction record before it is touched), and the 25-transfer soak's deliberate exclusion from the
routine gate.
