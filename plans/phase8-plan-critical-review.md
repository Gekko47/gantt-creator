# Independent critical appraisal — revised Phase 8 delivery plan

Reviewer stance: external, no prior involvement. Object under review:
[`phase8-png-export-review.md`](phase8-png-export-review.md) (the "plan"), appraised
against the landed tree at the Phase-7 exit and the governing sources
([`../docs/03-ROADMAP.md`](../docs/03-ROADMAP.md) Phase 8,
[`../docs/02-ARCHITECTURE.md`](../docs/02-ARCHITECTURE.md),
[`../docs/STATUS.md`](../docs/STATUS.md),
[`../docs/REPO-MAP.md`](../docs/REPO-MAP.md),
[`../docs/KNOWN-LIMITATIONS.md`](../docs/KNOWN-LIMITATIONS.md),
[`../docs/DECISIONS.md`](../docs/DECISIONS.md), and the nine R8 work-item guides).
Every load-bearing claim was re-verified against source before being accepted
here; claims the plan asserts without verification are called out as such.

---

## 1. Verdict in one paragraph

The plan correctly interprets the Phase 8 intent and its revisions are
code-grounded and directionally right: the clipboard-independent 300-DPI
channel as a SkiaSharp renderer in the Raster project (Core-only project
graph), consuming the immutable scene, cropped to
[`ExportBounds`](../src/GanttCreator.Core/Scene/ExportBounds.cs), with the
ribbon surface and gating following the landed R7.6 pattern. Its two most
consequential revisions (fixture via
[`SceneSnapshot.Deserialize`](../src/GanttCreator.Core/Scene/SceneSnapshot.cs:31);
scene reuse via [`EditableExportPreflight`](../src/GanttCreator.AddIn/EditableExportPreflight.cs))
survive stress-testing and are strengthened by facts the plan itself missed.
It is, however, **not yet approvable as written**: it propagates one stale
work-item reference, leaves the preview contract's empty state undefined,
omits two obligations (SBOM/third-party notices; the synchronous-render
boundary), and under-specifies the de-risking spike's acceptance criteria.
Judgment: **approve with revisions** — the eight findings in §4, none of which
change the phase's shape, all cheap to fold in before R8.1 starts.

---

## 2. Re-verified facts (accepted as sound)

| Plan claim | Re-verification |
| --- | --- |
| `ExportSize` landed, hardened, culture-invariant, safety-limited | Confirmed: [`ExportSize.cs`](../src/GanttCreator.Raster/ExportSize.cs) — `MaxPixelDimension` 65,535, `MaxTotalPixels` 100M, zero/sub-pixel rejection before allocation |
| Scene model complete for raster (primitives, style, z-order, deterministic order) | Confirmed: [`SceneStyle.cs`](../src/GanttCreator.Core/Scene/SceneStyle.cs) carries fill/stroke/width/hatch/font/bold/alignment/text-colour/dash; [`GanttScene.TryCreate`](../src/GanttCreator.Core/Scene/GanttScene.cs) orders deterministically |
| Crop authority is `ExportBounds` + `PlacementDelta`; scene rides the outcome | Confirmed: [`ExportBoundsResolver.cs`](../src/GanttCreator.Core/Scene/ExportBoundsResolver.cs) pins `EditableExport` and carries `Scene` |
| `SceneBuilder` treats all non-live profiles identically | Confirmed from source: profile gates at [`SceneBuilder.cs`](../src/GanttCreator.Core/Scene/SceneBuilder.cs) test only `== LiveExcel` (title, panel-bounds, sheet-coordinates) |
| `ReferenceSceneFixture` is internal to the Core test project | Confirmed: [`ReferenceSceneFixture.cs`](../tests/GanttCreator.Core.Tests/Scene/ReferenceSceneFixture.cs) — Revision A's premise for replacing it is correct |
| AddIn has no Raster reference; REPO-MAP claims one | Confirmed drift: [`GanttCreator.AddIn.csproj`](../src/GanttCreator.AddIn/GanttCreator.AddIn.csproj) vs [`REPO-MAP.md`](../docs/REPO-MAP.md) dependency table |
| No file-dialog port exists | Confirmed: zero `FileDialog` hits in `src/` |
| Reference scene carries no hatched primitive | Confirmed: the entity guide's Procurement row records the field is unexercised |
| Office-gate 780 s deadline conflict is open | Confirmed: [`STATUS.md`](../docs/STATUS.md) R7.7 close-out |
| D-G4 open; D-G5 must start by Phase 8; latest ADR is 0069 | Confirmed: [`IMPLEMENTATION-PLAN.md`](../docs/work-items/IMPLEMENTATION-PLAN.md), [`DECISIONS.md`](../docs/DECISIONS.md) |

**Two facts the plan did NOT verify, re-verified here — both favourable:**

1. **The snapshot fixture route is lossless.** Revision A depends on
   [`SceneSnapshot`](../src/GanttCreator.Core/Scene/SceneSnapshot.cs)
   serializing every renderer-relevant field. It does:
   `LeftInsetPt` (null-when-zero, lines 121/379), `TextColour` (274),
   `StrokeDash` (275), `HatchPattern` (269), `FontFamily`/`FontSizePt`/`Bold`/
   `Alignment`. The plan's fixture route stands on solid ground — but it never
   checked, and the golden test pins the serialize side only (see finding F2).
2. **The architecture boundary is stronger than the plan claims.**
   [`CoreBoundaryTests.cs`](../tests/GanttCreator.Architecture.Tests/CoreBoundaryTests.cs)
   already enumerates `SkiaSharp`, `SkiaSharp.Views` and `HarfBuzzSharp` as
   forbidden **Core** references, and the authority tests
   (metric/plot-geometry/sheet-layout) already scan
   `src/GanttCreator.Raster`. The suite was built anticipating this phase; no
   package-allowlist change is needed, and the plan's residual uncertainty
   here is closed favourably.

**A fact the plan states but under-uses:** the committed golden scene was
built with `Profile = SceneCompositionProfile.Raster`
([`SceneBuilderTests.cs`](../tests/GanttCreator.Core.Tests/Scene/SceneBuilderTests.cs)
line 2651), while R8.8's production scene is built with `EditableExport` via
the resolver. Without the plan's equivalence pin, the phase's golden baselines
(Raster-profile fixture) and its production scene (EditableExport) would
differ by profile label with **no test proving them equal**. Revision E is
therefore *required*, not merely recommended — see finding F3.

---

## 3. Stress tests of the four named focus areas

### 3.1 SkiaSharp version pinning and native asset loading

The plan is right that this is the phase's only go/no-go engineering unknown,
and right to refuse coding from memory (no SkiaSharp entry exists in
[`Directory.Packages.props`](../Directory.Packages.props)). Two gaps:

- **The spike's acceptance criteria are undefined (F6).** "Verify the packed
  add-in can load the native library" is a direction, not a gate. The spike
  must state pass/fail conditions: the packed XLL loads in Excel;
  `GanttCreator.Raster.dll` and the managed `SkiaSharp.dll` resolve from the
  pack; the native `libSkiaSharp` loads by the chosen route
  (beside-the-XLL / unpack-to-temp / load-from-bytes); and **one smoke render
  succeeds inside the Excel host**, not only in a unit test. Without the
  in-host arm the spike can pass while the shipped add-in fails.
- **The managed packing side is unaddressed.** The plan treats the spike as a
  native-asset question only. Excel-DNA packs managed assemblies into the
  XLL; the spike must confirm the Raster assembly and SkiaSharp's managed
  assembly are actually in the pack, since the AddIn gains its first Raster
  reference in this phase.

### 3.2 Residual `ExportSize` handling

Sound. The residual is genuinely small and the plan's sequencing clarification
(the equality pin lands with R8.7, where the preview producer exists) is
correct — the pin cannot exist before its producer. One check the plan passes
silently and should state: with Revision D, the preview needs exactly the
mathematics `ExportSize.ToPixels` already expresses (value, unit, aspect →
pixels), so the R8.1 stop-condition ("preview needs math `ExportSize` cannot
express") is unreachable by design. No parser change, no second formula — the
one-boundary discipline holds.

### 3.3 Preview contract definition

**Incomplete — the plan's largest contract gap (F4).** Revision D names the
aspect source but not its lifecycle. Undefined: what the preview shows before
any render has recorded an aspect (fresh workbook, first open), what happens
on workbook switch (ribbon state is workbook-scoped by R1.5, and the
warnings-snapshot precedent clears state on switch), and whether a stale
aspect can ever be shown against a changed table. The contract must state:
the cached export aspect is workbook-scoped state, cleared on workbook
switch, and the preview/export controls are disabled until a completed render
records an aspect. Without this, R8.7's "fast, side-effect-free getter" has an
undefined input and the row will hit a stop condition mid-implementation.

A second, smaller gap (F5): the roadmap's revision-10 note binds R8.7 to
"consume the size-preset contract R4.7H owns rather than defining its own
widths". The plan's R8.7 section designs free-form cm/in/px entry and mentions
[`SizePreset`](../src/GanttCreator.Core/SizePreset.cs) only inside the R8.4
widths proposal. The control spec must say where preset widths appear
(defaults or a preset list beside free-form entry) and that their values come
from `SizePresets.All`, never new literals.

### 3.4 Step-0 de-risking spike versus starting at R8.1

**The plan's sequencing is endorsed.** R8.1 is a documentation/verification
commit over renderer-agnostic arithmetic: it is valid under every possible
spike outcome, so nothing is wasted if the spike forces a technology ADR, and
it costs one small commit. Inverting the two buys nothing: the spike is R8.2
Step 0 and may begin immediately after R8.1 lands; deferring it past R8.2's
first code commit is what would be indefensible, and the plan already forbids
that. The reviewer rejects "spike first" as unnecessary and confirms the
plan's order — with the sharpening that the spike must not be allowed to
slip behind any renderer code (its whole value is being first).

---

## 4. Findings (prioritized)

| # | Finding | Severity | Fix |
| --- | --- | --- | --- |
| F1 | **Stale row reference propagated.** The R8.2 work item names "the packaging row (R9.5)"; the roadmap's packaging row is **R10.2** (R9.5 is accessibility). The plan echoes this in §3 ("feeds R9.5/R10.2") while §6 correctly says R10.2 — internally inconsistent. Neither the work item nor the plan names **R10.7** (SBOM/third-party notices), which SkiaSharp's licence obligates. | High (traceability; cheap) | Plan: cite R10.2 and R10.7 only. R8.2's commit fixes the work item's stale reference as a docs defect. |
| F2 | **Snapshot round-trip unguarded.** Revision A's fixture route is lossless today (verified §2), but the golden test pins serialize-side only; a future `SceneText`/`SceneStyle` field not added to the snapshot document would silently default in every deserialized baseline. | Medium-high (protects the phase's core fixture route) | R8.4 Step 0 adds a serialize→deserialize→re-serialize byte-identity round-trip guard. |
| F3 | **Profile-equivalence pin is required, not optional.** R8.4's baselines come from a Raster-profile fixture scene; R8.8's production scene is EditableExport. Without the pin, the renderer's contract tests and the production scene differ by an unproven label. | Medium-high (correctness) | Make Revision E's equivalence test a prerequisite of R8.4 landing, not just an R8.8 companion. |
| F4 | **Preview contract incomplete** (empty-cache state, workbook-switch clearing, disabled-until-rendered) — see §3.3. | High (contract gap; surfaces as a stop condition) | Complete the contract in the plan and in R8.7's Tier-A refresh. |
| F5 | **R8.7's size-preset obligation under-addressed** — see §3.3. | Medium | Name the `SizePresets`-derived width options in the control spec. |
| F6 | **Spike acceptance criteria undefined; managed packing unverified** — see §3.1. | Medium (shippability) | Define pass/fail conditions incl. the in-host smoke render. |
| F7 | **Synchronous-render boundary unstated.** The render runs on Excel's command thread; the 100M-pixel budget permits a ~400 MB frame that could block the host for a long time on a legal-but-absurd request. Cancellation/deadline is R9.1's row, so Phase 8 ships synchronous by design — but the plan should say so, and record typical (~3–10M pixel) versus worst-case budgets. | Medium (risk coverage) | One paragraph in the plan; R8.8's work item records the accepted behaviour and R9.1 ownership. |
| F8 | **R8.6's moved Office gate must be recorded as owed.** Revision C is consistent with repo precedent (multiple STATUS entries record "Required and owed"), but the roadmap forbids marking a row done with a Required gate silently unrun. | Medium (process compliance) | R8.6's STATUS entry explicitly records the live unwritable-path case as owed until R8.8's exit evidence. |
| F9 | **Prose overconfidence on `pHYs`.** §3 R8.5 says the probe is "likely to conclude" a patch step is required; the ledger correctly marks it inference. The prose should match the ledger. | Low | Rephrase as hypothesis; the Step-0 probe decides. |
| F10 | **R8.2 acceptance tests under-enumerated.** The plan describes the renderer but not its pins (exact-pixel regions, z-order draw-sequence, clipping, determinism, and the **no-recalculation positive** — the phase's core architectural claim, and the repo's positive-test rule applies). | Low-medium | Enumerate in the plan; the work item already carries them. |
| F11 | **Determinism metadata incomplete.** The determinism pin should name the CPU raster surface (no GPU) and the pinned antialias settings as part of the recorded baseline metadata. | Low | One line in R8.2's design decisions. |
| F12 | **Dialog threading unstated.** The save dialog runs on Excel's main STA thread from a ribbon callback — safe, and consistent with the WinForms dialog family; the plan should say so to preempt a `Task.Run` suggestion (forbidden around COM). | Low | One line in Revision B. |

---

## 5. Clarity, testability, maintainability

**Clarity.** Good. The plan separates intent, verified baseline, per-item
findings, revisions, and sequence, and its evidence ledger honestly marks
inference and unknowns. Weaknesses: the internal R9.5/R10.2 inconsistency (F1)
and the undefined preview empty-state (F4) are the two places a reader could
act on incomplete information.

**Testability.** Strong where it follows the work items (byte-level PNG
assertions, failure matrices with positives, golden policy with human
approval, independent re-read verification). Gaps: the round-trip guard (F2),
the equivalence pin's promotion (F3), and R8.2's test enumeration (F10). The
plan correctly keeps every validator paired with a positive test and keeps
the Raster 90/85 coverage thresholds in view.

**Maintainability.** The plan's surface claim is accurate and verified: one
new Raster namespace, one AddIn command plus one injected port, additive
Office-side state, **zero Core production changes** (Revision D's outcome
member is Office-side; the equivalence pin is a Core test). The closed
hatch/dash enums and the single font seam keep the renderer's mapping tables
data-driven. The one maintainability debt the plan accepts honestly — golden
baselines as reference-machine artefacts under the font ADR's Option B — is
already the test strategy's stated position.

---

## 6. Judgment

**Approve with revisions.** The plan's interpretation of Phase 8, its
sequencing, and its two structural revisions are sound and survive
stress-testing; the findings above are omissions and hygiene, not design
errors. Required before implementation begins, in priority order:

1. **F1** — correct the packaging-row references (R10.2, R10.7) in the plan;
   fix the R8.2 work item's stale "R9.5" in that row's commit.
2. **F4** — complete the preview contract (workbook-scoped aspect cache,
   cleared on switch, disabled-until-rendered).
3. **F3 + F2** — promote the profile-equivalence pin to a prerequisite of
   R8.4 and add the snapshot round-trip guard at R8.4 Step 0.
4. **F6** — define the spike's pass/fail criteria, including the in-host
   smoke render and the managed-packing check.
5. **F7** — state the synchronous-render behaviour, the pixel-budget
   reality, and R9.1 ownership.
6. **F5, F8, F9, F10, F11, F12** — fold in during the respective rows'
   Tier-A refreshes; each is a sentence or a test name, none changes scope.

With those folded in, the plan is fit to govern Phase 8 from R8.1 through the
phase-exit demonstration.
