# Implementation guides manifest and execution plan

> **Roadmap revision 10 (29 September 2026) added nine Tier-A guides ahead of
> R4.9:** `R4.7A` identity and hierarchy, `R4.7B` projection, `R4.7C` workbook
> presentation, `R4.7D` row geometry, `R4.7E` critical interval, `R4.7F` duration
> and date semantics, `R4.7G` label positions, `R4.7H` size presets, and `R4.8A`
> Refresh orchestration. Tier A goes from **41 to 50**. Three of them reverse
> landed behaviour and name what they delete: `R4.7D` removes
> `LaneLayoutBuilder`'s auto-growth and the auto-growth test, `R4.7E` changes the
> critical overlay from a rect-to-line reinterpretation to a filled rectangle
> drawn as built, and `R4.7G` removes two label positions. Their decisions are
> ADR-0026, ADR-0027 and ADR-0028. Read
> [`../PRE-R4.9-DECISION-REGISTER.md`](../PRE-R4.9-DECISION-REGISTER.md) first:
> the two input documents these guides implement are not self-consistent with the
> entity guide, and the register records which of their proposals were rejected.
>
> **Implementation state as of 2026-09-30: all eight R4.7 rows have landed;
> R4.8A is next.** This header previously read "the guides are authored; **none of
> their product code has been written**. `R4.7A` is next", which had been true
> when the guides were authored and was **stale** by the time it was read. The
> verified state, from `docs/STATUS.md` and each guide's own status section:
>
> | Row | State | Notes |
> | --- | --- | --- |
> | R4.7A | **Landed** | identity and hierarchy; took the schema 3 → 4 |
> | R4.7B | **Landed** | projection; golden regenerated (`14b47f3`, `11130c3`) |
> | R4.7C | **Landed** | workbook presentation; took the schema 4 → 5 |
> | R4.7D | **Landed** | row geometry and outline groups; Office gate 46/46; golden regenerated; collapse-preservation defect in the height pass fixed 2026-10-11 (see the amendment note below) |
> | R4.7E | **Complete** | critical interval as a parent-independent filled rectangle; golden regenerated (`806a632`) |
> | R4.7F | **Implemented** | derived `Duration`; Office gate **not run**, so the row is not complete |
> | R4.7G | **Landed** | horizontal-only label positions; took the schema 5 → 6 |
> | R4.7H | **Implemented** | `SizePreset` and the single `PlotGeometryResolver` |
> | **R4.7I** | **Source landed; coverage closed 2026-10-04** | live chart anchored to the worksheet, not the page ([ADR-0030](../adr/0030-live-chart-vertical-anchoring.md)). All five slices are in `src/`, so the guide's own "not started" header was wrong. Its lane anchoring had **zero** test coverage — mutating it off left 2,477 tests green — and is now pinned by 18 resolver tests plus 2 factory wiring/geometry tests. **Office gate PASSED 2026-10-04**, in the same human-driven F5 session that closed R4.9's. |
> | **R4.8A** | **Implemented** | Refresh orchestration — the last row before R4.9; Office gate still outstanding |
>
> `R4.7A` → `R4.7B` → `R4.7D` were strictly sequential and are done; `R4.7C`,
> `R4.7F`, `R4.7G` and `R4.7H` were independent; `R4.7E` needed `R4.7B` and
> `R4.7C`; **`R4.8A` needs all eight R4.7 rows**, which is why it is the only one
> still open. `scripts/check-status.ps1` requires a letter-suffixed roadmap ID to
> have a guide here, so a row cannot ship without one.
>
> **`R4.7I` was added 2026-10-01** from a live defect report and is **not** part of
> that original set of eight. It amends three of them: R4.7D's measured lane
> geometry, R4.7H's plot-geometry authority, and R4.8A's request factory. Its ADR
> records the finding that ADR-0026 D3's `lane Top == Excel row Top` was
> **unrepresentable** — `PanelCellGrid` carried no absolute sheet origin, and the
> plot's top and height came from the size preset's page coordinates — so no bar
> could line up with its row. It is sequenced before R4.9's live Refresh gate.
>
> **R4.7D amendment (2026-10-11):** a live defect showed Refresh
> re-expanding a parent the user had collapsed. D5's promise ("collapse/expand
> state is preserved across Refresh") was violated by the row-height
> normaliser, not the outline writer: its level-1 padding pass wrote
> `OutlineLevel` across every body row. The normaliser's two passes are now
> disjoint (level 1 → ungrouped rows only, level 2 → grouped rows only), so
> a grouped row's outline is written exactly once. Pinned live by
> `RowHeightNormaliserIntegrationTests.A_collapsed_group_stays_collapsed_through_height_normalisation`
> (Excel 16.0.20430.20146) and by the pass-split contract tests. The
> outline writer's own reset pass was fixed the same day by `b9e2145`;
> both `OutlineLevel`-writing paths now honour D5.
>
> **The schema version is now 7**, four bumps past the 3 this header predates
> (3→4 R4.7A, 4→5 R4.7C, 5→6 R4.7G, 6→7 the R4 QA review's `SizePreset` and
> `RangePaddingDays` settings keys); ADR-0029 D5 holds the full progression.
> The standing risk into R4.8A is the Office COM leak ratchet, at **26 forced
> kills against a ceiling of 26** (`KNOWN-LIMITATIONS.md` L19) — R4.8A's new
> adapter surface is the first thing that could push it over.

> Reauthored 23 September 2026 from the landed R2.7a state to incorporate the approved [`GanttCreator_R2_Implementation_Plan.md`](../GanttCreator_R2_Implementation_Plan.md), and revised 26 September 2026 at the Phase-3 exit: R3.14 landed, and R4.1–R4.10 were upgraded from Tier B to Tier A against the landed Phase-3 code. The current index contains **105** roadmap-guide records: **43 Tier A and 62 Tier B**, counted from the table on 2026-10-31, including the two later Tier-A additions R4.7I and R4.11. The earlier 94/41/53 count was a dated snapshot and omitted later roadmap additions. Historical recount notes below explain the earlier discrepancies; they do not describe the current index. [`../STATUS.md`](../STATUS.md) records implementation and acceptance state; `AGENTS.md` retains requirement precedence.

## Tier model (approved 2026-09-21)

- **Tier A — prescriptive.** 41 guides: the Phase 2 remainder plus nine R2
  hardening rows, Phase 3, and the ten Phase-4 rows upgraded at the
  Phase-3 exit. They name exact files, types, seams, tests, and
  evidence commands; each is re-verified at implementation time.
- **Tier B — binding contract.** 53 guides for Phases 5–10. They fix the
  behaviour, acceptance criteria, governing document sections, design
  decisions, stop-points, and evidence commands, while implementation steps
  reference predecessor artifacts by work-item ID plus a mandatory Step 0
  verification, because the renderer types they build on do not exist
  yet. Inventing them now would violate the anti-hallucination rules.
- **Just-in-time upgrade.** At each phase exit, the next phase's Tier-B
  guides are upgraded to Tier A detail against the landed code (exact
  files, tests, commands; prerequisites re-verified) **before any of them
  is implemented**. The upgrade is recorded in each guide's status line.
  The 2026-09-26 Phase-3 exit upgrade deliberately left four questions
  **unprobed rather than guessed** — the Office shape kind for scene
  lines (R4.3), the freeform point semantics (R4.5), the pattern/alpha
  members (R4.6), and the per-property interop shapes (R4.2). Each is a
  named Step-0 obligation in its guide with an `unknown until probed`
  ledger row; a Tier-A upgrade fixes names against landed code, it does
  not license inventing host behaviour.
- **Two of those four were answered by R4.1's Step-0 probe (2026-09-27),
  and the probe falsified a fifth assumption nobody had flagged.** The
  probe against the installed PIA resolved the R4.2 per-property
  interop shapes (all six readable, and `Selection` confirmed read-only,
  so restore is `Range.Select()`) and half of R4.3's line question
  (endpoints are exact; R4.3 then decides plain line over connector). The
  freeform (R4.5) and pattern/alpha (R4.6) questions remain unprobed.
  **The falsified assumption: `Microsoft.Office.Interop.Excel.Shape` has
  no `Tag` member**, which had been R4.1 D2's ownership carrier and was
  inherited as a "fact" by R4.3 D2, R4.8 D1, and R9.4. `Tag` is a
  *PowerPoint* member. [ADR-0019](../adr/0019-shape-ownership-carrier.md)
  moves the carrier to `AlternativeText` and bounds the tag as
  `GanttCreator.Owned.v1:{hash}`, and the four dependent guides are
  amended. The lesson for the next just-in-time upgrade is the guide's own
  rule working: a Tier-A upgrade may fix names against landed code, but a
  **host-behaviour** claim that no probe has touched is a proposal, not a
  fact, however plausible the member name looks.

## Implementer contract (read before executing any guide)

1. The `AGENTS.md` source-of-truth order governs. A guide never overrides
   it. If a guide conflicts with `AGENTS.md`,
   [`../02-ARCHITECTURE.md`](../02-ARCHITECTURE.md), or
   [`../07-GANTT-ENTITY-GUIDE.md`](../07-GANTT-ENTITY-GUIDE.md), stop and
   report the conflict; the scope-change protocol in
   [`../03-ROADMAP.md`](../03-ROADMAP.md) is the only override path.
2. Read the whole guide, then every section it links, before editing.
3. **Step 0 of every guide verifies prerequisites against current source.**
   A mismatch is drift: record it in the guide's "Notes during
   implementation" ledger and reconcile against the landed work item; a
   real conflict stops the work item.
4. Implement in guide order. One roadmap row is one commit (Conventional
   Commit prefix; list the certified `docs/08-TEST-CHECKLIST.md` sections
   in the commit message).
5. Every new `if (bad) { error }` validator ships with a positive test in
   the same commit (AGENTS.md rule).
6. Gates: `pwsh ./scripts/verify-quick.ps1` every commit;
   `pwsh ./scripts/verify.ps1` at branch end; `pwsh ./scripts/verify-office.ps1`
   when the guide's Office gate is Required. Never claim a gate ran without
   observed output.
7. Never invent Office, Excel-DNA, or SkiaSharp behaviour. Verify against
   installed metadata or vendor documentation and record an evidence-ledger
   row (statuses: fact / inference / proposal / unknown).
8. `AGENTS.md` stop conditions apply everywhere; each guide adds
   item-specific ones.
9. After landing: update `docs/STATUS.md`; update `docs/REPO-MAP.md` when a
   mapped responsibility, entry point, test location, or verification
   command changes; persist the work-item completion fact and any
   irreversible decision (memra); advance the guide's status line.
10. Office-host gates marked "human" are recorded by the human operator in
    the guide's Notes section; an agent never marks them satisfied.

## Phase-exit obligations

Each phase exit must satisfy the roadmap row's stated automated
demonstration and its compatibility-matrix subset
([`../03-ROADMAP.md`](../03-ROADMAP.md) "Cross-phase compatibility matrix"),
and additionally **upgrade the next phase's guides to Tier A** as described
above. Phase 3's guides are Tier A from initial authoring. ADR-0010 is the
only approved exception to table-order execution: after R4.9, execute R2.10,
then R3.13, before R4.10.

## Index of guides (105)

One row per roadmap-guide record. Status values distinguish implementation
from acceptance evidence: `Landed` means the implementation is committed and
its required evidence is closed; `Implemented` means code/tests are present but
one or more required gates remain; `Approved` means a decision exists but the
implementation has not started; `Authored` means a binding guide exists but
implementation has not started. A row note names material open evidence.
Filenames are backticked, not linked, so link checking stays on authored
files only; the local audit verifies every file exists.

### Phase 2 remainder — configuration, hardening, and data commands (Tier A)

| ID | Guide | Status |
| --- | --- | --- |
| R2.7 | `R2.7-config-catalogues.md` | Landed |
| R2.7a | `R2.7a-destructive-command-policy.md` | Landed |
| R2.1a | `R2.1a-type-identity-contract.md` | Landed (`3ef29b6`) |
| R2.2a | `R2.2a-safe-worksheet-adoption.md` | Landed (`b315b14`) |
| R2.4a | `R2.4a-date-system-boundary.md` | Landed (`982cf05`) |
| R2.4b | `R2.4b-neutral-cell-state.md` | Landed (`bfb334c`) |
| R2.5a | `R2.5a-validation-normalisation.md` | Landed (`f21a908`) |
| R2.6a | `R2.6a-validation-reporter-boundary.md` | Landed (`bf9727c`) |
| R2.7b | `R2.7b-style-capability-schema.md` | Landed (`49abd4f`; ADR-0009) |
| R2.7c | `R2.7c-style-registry-formatting.md` | Landed (`69a986d`) |
| R2.7d | `R2.7d-date-display-format-setting.md` | Landed (ADR-0016) |
| R2.8 | `R2.8-add-row-commands.md` | Landed (`61e67f9`; fix `75df937`) |
| R2.9 | `R2.9-type-dropdown-materialisation.md` | Implemented; required Excel dropdown/paste acceptance evidence remains to be reconciled |
| R2.10 | `R2.10-config-repair-migration.md` | Implemented; automated synthetic-damage coverage is present, human Excel acceptance remains required |

### Phase 3 — Core scene engine (Tier A)

| ID | Guide | Status |
| --- | --- | --- |
| R3.1 | `R3.1-geometry-value-objects.md` | Landed (`b5314c7`) |
| R3.2 | `R3.2-scene-primitives.md` | Landed (`e244d3c`) |
| R3.3 | `R3.3-time-scale.md` | Landed (`e8a77e8`) |
| R3.4 | `R3.4-lane-stack-geometry.md` | Landed (`4ba0601`) |
| R3.5 | `R3.5-plot-frame-bands.md` | Landed (`9e088bf`, `0f577bd`) |
| R3.6 | `R3.6-span-bar-label-layout.md` | Landed (`5d49f9c`) |
| R3.7 | `R3.7-milestone-marker-layout.md` | Landed (`e518ca6`) |
| R3.8 | `R3.8-critical-interval-overlay.md` | Landed (`31b72ac`, `eb00c71`) |
| R3.9 | `R3.9-multi-event-stack-lanes.md` | Landed (`6b733a4`) |
| R3.10 | `R3.10-delineator-lines-labels.md` | Landed (ADR-0017) |
| R3.11 | `R3.11-table-header-primitives.md` | Landed (`d763e67` step 1 grid, `aeea924` step 2 panel/header, `2012174` step 3 date labels) |
| R3.12 | `R3.12-scene-validator-benchmark.md` | Landed (`abd52c9` steps 1-2, `808162f` step 3, `2e40998` step 4, `b46c59c` steps 5-6) |
| R3.13 | `R3.13-mutation-testing.md` | Implemented; first full-Core mutation baseline and enabling the local verification step remain required |
| R3.14 | `R3.14-equivalence-thin-slice.md` | Landed (`aa50633`, exit record `0b985bb`, guide upgrade `98e744a`) |
| R3.15 | `R3.15-scene-chart-bounds-single-source.md` | Landed (`e3efd4e`; golden `d8acd0d`) |
| R3.16 | `R3.16-equivalence-field-table.md` | Landed (guide revision 6; D-G14 accepted 2026-09-27) |
| R3.17 | `R3.17-header-label-identifier-convention.md` | Landed (one convention; golden regenerated) |

### Phase 4 — live renderer (Tier B; upgrade to Tier A at Phase 3 exit)

| ID | Guide | Status |
| --- | --- | --- |
| R4.1 | `R4.1-excel-adapter-interfaces.md` | **Landed** (`IShapeWritePort`, `IPanelGridMeasurementPort`, `ShapeOwnershipTag`; D2 corrected by ADR-0019; live `AlternativeText` probe **discharged** 2026-09-27) |
| R4.2 | `R4.2-application-state-scope.md` | **Landed** (`ExcelApplicationStateScope`, five settings per ADR-0020; Office gate PASS 39/39) |
| R4.3 | `R4.3-shape-render-conversion.md` | Landed; required renderer evidence closed |
| R4.4 | `R4.4-text-alignment-conversion.md` | Landed; required renderer evidence closed |
| R4.5 | `R4.5-polygons-z-order.md` | Landed; required renderer evidence closed |
| R4.6 | `R4.6-style-token-mapping.md` | Landed; required renderer evidence closed |
| R4.7 | `R4.7-refresh-idempotence.md` | **Landed** (`ShapeReconciler` + `ShapeReconcilePlan`; D1's create-only premise expired with R4.6 and **D5 was added** replacing the unachievable "zero operations" pin with a fixed-point property, both approved 2026-09-28; **L18 decided** — refuse and surface; Office gate PASS 46/46) |
| R4.8 | `R4.8-unowned-content-preservation.md` | Landed; required preservation evidence closed |
| R4.8A | `R4.8A-refresh-orchestrator.md` | **Implemented** (`GanttRefreshOrchestrator`, `ExcelSceneBuildRequestFactory`; D8 wired). **Office gate CLOSED 2026-10-04** by R4.9's session (Excel `16.0.20430.20032`) — all six named conditions observed, including milestones |
| R4.9 | `R4.9-refresh-command.md` | **Implemented; Office gate PASSED 2026-10-04** (Excel 16.0.20430.20032, human-driven F5 session; Tier A 2026-09-26; first live slice. D1 amended onto R4.8A D1 — the command is a thin boundary, not the pipeline) |
| R4.7I | `R4.7I-live-chart-vertical-anchoring.md` | **Source landed**; anchoring test coverage closed 2026-10-04; **Office gate PASSED 2026-10-04** in the same session |
| R4.10 | `R4.10-thousand-event-performance.md` | Excluded by owner from this implementation; do not implement or run its benchmark |
| R4.11 | `R4.11-label-text-colour-propagation.md` | **Landed** (`e286dfe`; golden `e80bc8f`); **Office gate PASSED 45/45**, Office `16.0.20326.20158` x64. Independent of R4.9 and closed on its own live run. **Added to this manifest 2026-10-04** — the row existed in the roadmap, had a guide, and had landed, but was absent here |

### Phase 5 — interaction and UX (Tier B guides; implementation-state snapshot below)

| ID | Guide | Status |
| --- | --- | --- |
| R5.1 | `R5.1-plot-range-modes.md` | Implemented; live review of date edits, warnings, and refresh behavior remains |
| R5.2 | `R5.2-scale-layout-choices.md` | Implemented; live scale, margin, preset, and displayed-width review remains |
| R5.3 | `R5.3-parent-child-commands.md` | Implemented; live parent/child, Add Child, delete/promote, and selector review remains |
| R5.3 WI-E | `R5.3-wi-e-add-child.md` | Implemented; live Add Child and rendered-result review remains |
| R5.4 | `R5.4-move-expand-collapse.md` | Implemented; live ordering/alignment review and native outline-handle trigger proof remain |
| R5.5 | `R5.5-label-controls.md` | Implemented; live description-label control review remains |
| R5.6 | `R5.6-style-theme-settings.md` | Implemented; live theme edit, contrast warning, save/reopen, and rendered-result review remain |
| R5.6a | `R5.6a-named-style-presets.md` | Implemented; live custom-style create/assign/rename and rendered-result review remains |
| R5.6b | `R5.6b-chart-title-settings.md` | Implemented; live title visibility/edit and save/reopen review remains |
| R5.6c | `R5.6c-theme-dialog-modernisation.md` | Implemented; visual Excel-host review remains |
| R5.6d | `R5.6d-theme-dialog-and-chrome-fixes.md` | Implemented; live chrome/style/row-height review remains |
| R5.6e | `R5.6e-chart-frame-encloses-the-table.md` | Implemented; live frame-enclosure review remains |
| R5.7 | `R5.7-delineator-workflow.md` | Implemented; live create/edit/delete and refresh review remains |
| R5.8 | `R5.8-warnings-panel.md` | Implemented; live warning, non-mutation, and invalid-entry recovery review remains |
| R5.9 | `R5.9-ribbon-completion.md` | Implemented; keyboard/accessibility/visual review and the documented property-column scope decision remain |
| R5.10 | `R5.10-selection-overrides.md` | Implemented; live preventive gating and override persistence review remains |
| R5.11 | `R5.11-refresh-only-enforcement.md` | Implemented; production scheduler-to-render, empty-chart, trigger matrix, and failure-recovery host evidence remain |
| R5.12 | `R5.12-period-label-fit-ladder.md` | Implemented; live Month MM/MMM result and setting persistence review remains |
| R5.13 | `R5.13-group-aware-row-insert.md` | Implemented; live group-aware insertion and outline review remains |
| R5.14 | `R5.14-date-labels.md` | Implemented; live date-label selection and placement review remains |
| R5.15 | `R5.15-managed-worksheet-protection.md` | Not implemented; proof-first live feasibility and ADR decision required before any protection adoption |
| R5.16 | `R5.16-ribbon-first-run-and-data-group.md` | Implemented 2026-11-10/11 (ADR-0056) across sixteen commits; `verify-quick`, `verify` and `verify-office` green (61/61 live tests); first-run surface observed live by the owner; the state-2 swap and the menu/gate/insertion passes are still owed |
| R5.17 | `R5.17-repair-rebuilds-a-missing-data-table.md` | Implemented 2026-11-11 (ADR-0057, amendments 2026-11-11) with R5.17a–d in one recovery pipeline; build `-warnaserror` green and the targeted suites pass; `verify-office` not run, so the worksheet-mutating paths are unproved in the host; the Gantt-sheet-writer extraction is recorded debt |
| R5.17e | `R5.17e-single-chart-locator-seam.md` | Planned 2026-11-11; not started. Deferred deliberately: the measured adoption surface is 28 comparison sites across 20 adapters and a locator with one implementation has no consumer yet |
| R5.17f | `R5.17f-no-owner-damage-states.md` | Planned 2026-11-11; not started. Collects the states that survive a refresh and Fix, each needing an integrity finding before a repair |

**Phase 5 completion snapshot (2026-10-31):** R5.1–R5.14 feature code and
automated coverage are present in the current source. Their required live
acceptance checks are not uniformly closed; the R5.11 end-to-end scheduler host
test has been added but could not start Excel in this session (`0x80070520`).
Native outline-handle events remain unproven. R5.15 has not started. Complete
the live acceptance matrix and the unresolved R5.9 property-column scope
decision before claiming Phase 5 exit. The code-side refresh foundation does
not include R4.10 or its benchmark.

### Phase 6 — PowerPoint editable export (Tier B; upgrade at Phase 5 exit; R6.7 spike first)

| ID | Guide | Status |
| --- | --- | --- |
| R6.7 | `R6.7-powerpoint-compatibility-spike.md` | Authored (Tier B; spike — runs first) |
| R6.1 | `R6.1-export-bounds-model.md` | Authored (Tier B) |
| R6.2 | `R6.2-table-composition.md` | Authored (Tier B) |
| R6.3 | `R6.3-chart-staging.md` | Authored (Tier B) |
| R6.4 | `R6.4-grouping.md` | Authored (Tier B) |
| R6.5 | `R6.5-staging-lifecycle.md` | Authored (Tier B) |
| R6.6 | `R6.6-clipboard-copy.md` | Authored (Tier B) |
| R6.8 | `R6.8-copy-editable-command.md` | Authored (Tier B) |

### Phase 7 — direct PowerPoint transfer (Tier B; upgrade at Phase 6 exit)

| ID | Guide | Status |
| --- | --- | --- |
| R7.1 | `R7.1-powerpoint-adapters.md` | Authored (Tier B) |
| R7.2 | `R7.2-attach-slide-cases.md` | Authored (Tier B) |
| R7.3 | `R7.3-paste-special-verify.md` | Authored (Tier B) |
| R7.4 | `R7.4-slide-fit-position.md` | Authored (Tier B) |
| R7.5 | `R7.5-transfer-failures.md` | Authored (Tier B) |
| R7.6 | `R7.6-send-to-powerpoint.md` | Authored (Tier B) |
| R7.7 | `R7.7-transfer-soak.md` | Authored (Tier B) |

### Phase 8 — 300-DPI PNG export (Tier B; upgrade at Phase 7 exit; R8.2a precedes R8.3)

| ID | Guide | Status |
| --- | --- | --- |
| R8.1 | `R8.1-width-parsing-residual.md` | Authored (Tier B) |
| R8.2 | `R8.2-skia-primitive-renderer.md` | Authored (Tier B) |
| R8.2a | `R8.2a-font-pinning-adr.md` | Authored (Tier B; ADR) |
| R8.3 | `R8.3-text-hatch-rendering.md` | Authored (Tier B) |
| R8.4 | `R8.4-representative-render.md` | Authored (Tier B) |
| R8.5 | `R8.5-png-metadata.md` | Authored (Tier B) |
| R8.6 | `R8.6-atomic-write-validator.md` | Authored (Tier B) |
| R8.7 | `R8.7-width-controls-preview.md` | Authored (Tier B) |
| R8.8 | `R8.8-export-png-command.md` | Authored (Tier B) |

### Phase 9 — recovery, resilience, and maintenance (Tier B; upgrade at Phase 8 exit)

| ID | Guide | Status |
| --- | --- | --- |
| R9.1 | `R9.1-cancellation-deadline.md` | Authored (Tier B) |
| R9.2 | `R9.2-structured-diagnostics.md` | Authored (Tier B) |
| R9.3 | `R9.3-corrupted-settings-recovery.md` | Authored (Tier B) |
| R9.4 | `R9.4-shape-ownership-repair.md` | Authored (Tier B) |
| R9.5 | `R9.5-accessibility-labels.md` | Authored (Tier B) |
| R9.6 | `R9.6-legacy-mapping-docs.md` | Authored (Tier B; docs-first) |
| R9.7 | `R9.7-crash-recovery-matrix.md` | Authored (Tier B) |
| R9.8 | `R9.8-offline-audit.md` | Authored (Tier B) |

### Phase 10 — packaging and release candidate (Tier B; upgrade at Phase 9 exit) — FINAL PHASE

| ID | Guide | Status |
| --- | --- | --- |
| R10.1 | `R10.1-supported-office-matrix.md` | Authored (Tier B) |
| R10.2 | `R10.2-release-packaging.md` | Authored (Tier B) |
| R10.3 | `R10.3-code-signing.md` | Authored (Tier B) |
| R10.4 | `R10.4-install-instructions.md` | Authored (Tier B) |
| R10.5 | `R10.5-full-suite-run.md` | Authored (Tier B) |
| R10.6 | `R10.6-exploratory-workflow.md` | Authored (Tier B) |
| R10.7 | `R10.7-sbom-freeze.md` | Authored (Tier B) |
| R10.8 | `R10.8-release-candidate.md` | Authored (Tier B; final roadmap row) |

## Decision register (product decisions a guide may not make alone)

| ID | Decision | First needed by | Where recorded | Status |
| --- | --- | --- | --- | --- |
| D-G1 | Configuration-sheet table set, settings storage, and storage format | R2.7 | ADR-0007 | Accepted and landed |
| D-G2 | Destructive-command undo semantics | R2.7a | ADR-0008 | Accepted and landed |
| D-G3 | Mutation-testing tool and version | R3.13 | R3.13 guide | **Approved 2026-10-05** by the human's instruction to install and proceed — `dotnet-stryker` 5.0.0 pinned in `tool-versions.psd1`, gate wired into `ci.yml` (the `verify.ps1` step of the same name is a disabled stub pending the first full-Core baseline run). Pester guard in `scripts/mutation-gate.Tests.ps1` (26 tests, 12 positive); **a green mutation-gate run has not yet been observed** — the guard's first observed pass is recorded in `docs/STATUS.md` when it happens. No longer open. |
| D-G4 | Pinned-font provisioning | R8.2a | R8.2a guide | Open — licence-sensitive; human choice required |
| D-G5 | Code-signing certificate procurement | by Phase 8 | R10.3 guide | Open — external lead time; start no later than Phase 8 |
| D-G6 | PowerPoint interop dependency (package + version pin) | R7.1 | R7.1 guide; ADR + `Directory.Packages.props` | Open — new production dependency; human approval required before first install |
| D-G7 | Named-style label/colour capability schema | R2.7b | ADR-0009 | Accepted 2026-09-23 |
| D-G8 | Defer R2.10 and R3.13 until after the R4.9 first-live slice | R2.9/R3.12 | ADR-0010 | Accepted 2026-09-23 |
| D-G9 | Label `Auto` cascade order and the blocked-label widest-gap truncation fallback | R3.6 | ADR-0015 | Accepted 2026-09-26 |
| D-G10 | Approved event-date display format | R3.11 | ADR-0016 | Accepted 2026-09-26 — `dd/mm/yyyy`, new `DateDisplayFormat` setting, schema 2→3 |
| D-G11 | Clipped-event date-label policy | R3.11 | Entity guide §23 revision 5; product owner 2026-09-26 | Accepted — a clipped event always shows the TRUE date; code-owned rule, not a workbook setting, so no schema advance |
| D-G12 | Shared scene ownership for a deduplicated entity | R3.10 | ADR-0017 | Accepted 2026-09-26 — third owner kind `Rows`; amends ADR-0013 |
| D-G13 | Text alignment gets its own enum | R3.11 | [ADR-0018](../adr/0018-text-alignment-is-not-a-label-position.md) | Accepted 2026-09-26 - add a closed `GanttTextAlignment` (`Left`, `Centre`, `Right`) and retype `SceneText.Alignment` and `SceneStyle.Alignment` from `GanttLabelPosition`. Section 4's centred-header rule is currently inexpressible because a label-position enum is being reused as a text-alignment enum. **Landed before `R3.12`**, which is the row that commits the golden scene snapshot. Landing it first avoided a golden-fixture regeneration, not a frozen format: `SceneSnapshot` carries an explicit `Version` and refuses anything else, so a later change is a version bump plus a regenerated fixture (ADR-0018 D2, which corrects the earlier "point of no return" claim). ADR-0013's real constraint is that *consumers* migrate together, and there are none yet. Rejected: adding `Centre` to `GanttLabelPosition` (entrenches the type confusion); shipping `Left`/`Right` only (a visible section 4 violation) |
| D-G14 | The per-entity equivalence field table is a **restatement** of the landed model, not an extension of it | R3.16 | R3.16 guide; entity guide revision 6 | **Accepted 2026-09-27 by the product owner**; proposed 2026-09-26 — product-owner approval required before the guide is edited.** Verified against the committed golden on 2026-09-26, which corrected three rows of the first draft: the panel row has no fixture, because panel primitives are emitted only when a `PanelTheme` is supplied and the reference build supplies none; the "Labels" row was two rows, since description/date labels sit at `ZLayer` 70 with lane and stack keys while a delineator label sits at `ZLayer` 75 with both null; and the header labels are separate identities at `ZLayer` 80, the year label using a `chart:year-label:` prefix and the period label appending `:label` to its rectangle's id. The delineator row now also states that a same-date pair shares one line but keeps one label per row. The table names only fields the model already carries; any cell needing a new field reads "not carried by the model; excluded", and no token value, z-layer, or label position changes. It is keyed to **primitive kind, owner, and lifetime** rather than to the guide's prose entity names, because one prose row bundles primitives with different owners and lifetimes, and owner plus lifetime are what change renderer behaviour. Chosen over extending the model to fill the table, which would make the contract fit the model instead of recording it, and over leaving the field lists in the test file, which leaves the renderer contract in an agent's reading rather than in the guide. The optional legend §25 becomes a row marked empty rather than an omission. **Identifier scope:** the role suffix and `{placeholder}` structure are contractual because the identifier is the shape name and the reconciliation key, but the formatting of an interpolated value is not — `chart:grid:{x}` formats a double `"R"` and `chart:period:{yyyy-MM-dd}` a date today, and changing either changes the identifier for every affected shape, so reformatting is an owned-set replacement rather than a cosmetic edit. `SceneGroup` membership and child order stay out (R6.4). The table states the panel row's contract but explicitly not its test coverage, because none exists. Instrument is a decision-register entry, not a new ADR, because no architectural rule changes — D-G14 makes an existing model rule explicit and consultable |
| D-G15 | A renderer **translates the chart bounds to a zero origin and carries the delta**, and the **outer padding stays inside the bounds on all four sides** | R3.15 (question it left open) | Entity guide §1 revision 7; R4.3 Step 0; R6.1 D1 | **Decided 2026-09-27 by the product owner; PROVISIONAL pending the R4.3 Step-0 probe.** `ChartBounds` is derived by expanding the content union by `ChartOuterPaddingPt` (`FrameBandsBuilder`), so a chart whose content starts at the origin has a **negative** bounds origin. That is legal in the scene — `RectD` permits negative coordinates deliberately — but Excel cannot express a negative shape offset, so placement must be computed rather than copied. Chosen over anchoring to the caller's `PlotBounds`, which would make the panel's measured relationship to the plot the chart's frame and reintroduce a second source of truth of exactly the kind R3.15 removed; and over refusing a negative origin, which would reject the catalogue's own default `ChartOuterPaddingPt = 6` for the ordinary panel-at-origin layout. The translation is uniform across every primitive kind — a rect's `Bounds`, a line's `From`/`To`, a polygon's `Points`, a text's `TextBounds` — and must change no relationship inside the scene, including whether a label sits outside the chart. **The padding is inside the bounds, not a renderer-side margin**, because §1 also says empty whitespace outside the bounds is not exported: a crop computed from an un-translated `ChartBounds` would silently drop exactly `ChartOuterPaddingPt` of margin, which is why R6.1 D1 now requires the delta. **Both rules were already true of the code before this decision** — the padding was always applied inside the bounds, and `FrameBandsBuilderTests` already asserted the resulting `-6` origin — so this records an existing behaviour and adds the placement rule rather than changing geometry. Provisional means exactly one thing: how the *host* handles a negative offset is unprobed, and R4.3 carries a named Step-0 obligation to establish whether it is rejected, clamped, or re-anchored. The rule is what the host must end up doing, not evidence that it does |

## Known drift resolved by revision 5

- KNOWN-LIMITATIONS L4 now names R8.4 as the first approved-PNG gate; R3.12 produces only the deterministic scene snapshot.
- The Show-Properties contract is defined by entity-guide revision 3. R5.9 still owns the separate undefined in-cell validation scope for `StyleKey`/`LabelPosition`.

## Authoring and reconciliation verification

The original 80-guide corpus was authored on 21 September 2026. Revision 5 added eight roadmap-guide records (seven R2 hardening rows and R5.6a), bringing the manifest to 88. The current local reconciliation is verified by Markdown link checking, roadmap-ID/guide existence checks, `git diff --check`, and the relevant automated test/gate commands as work items land. Office-host gates are never inferred from documentation-only authoring.
