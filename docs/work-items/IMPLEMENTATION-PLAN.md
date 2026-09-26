# Implementation guides manifest and execution plan

> Reauthored 23 September 2026 from the landed R2.7a state to incorporate the approved [`GanttCreator_R2_Implementation_Plan.md`](../GanttCreator_R2_Implementation_Plan.md). The manifest now covers 90 roadmap-guide records: 27 Tier A and 63 Tier B. R2.7, R2.7a, and R2.7b are landed; the R2.1a/R2.2a/R2.4a/R2.4b/R2.5a/R2.6a hardening rows and R5.6a have approved guides. ADR-0009 defines style capabilities, R2.7c carries the style registry's resolved formatting, and ADR-0010 records the executable first-live-slice order. [`../STATUS.md`](../STATUS.md) records landed state and `AGENTS.md` retains requirement precedence.

## Tier model (approved 2026-09-21)

- **Tier A — prescriptive.** 27 guides: the Phase 2 remainder plus eight R2
  hardening rows, and Phase 3. They name exact files, types, seams, tests, and
  evidence commands; each is re-verified at implementation time.
- **Tier B — binding contract.** 63 guides for Phases 4–10. They fix the
  behaviour, acceptance criteria, governing document sections, design
  decisions, stop-points, and evidence commands, while implementation steps
  reference predecessor artifacts by work-item ID plus a mandatory Step 0
  verification, because the scene/renderer types they build on do not exist
  yet. Inventing them now would violate the anti-hallucination rules.
- **Just-in-time upgrade.** At each phase exit, the next phase's Tier-B
  guides are upgraded to Tier A detail against the landed code (exact
  files, tests, commands; prerequisites re-verified) **before any of them
  is implemented**. The upgrade is recorded in each guide's status line.

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

## Index of guides (90)

One row per roadmap-guide record. Status values: `Landed`, `Implemented`
(working-tree implementation with observed gates; pending commit), `Approved`
(file and decision exist; not implemented), or `Authored` (binding contract).
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
| R2.7c | `R2.7c-style-registry-formatting.md` | Implemented (working-tree implementation with observed gates; pending commit) |
| R2.8 | `R2.8-add-row-commands.md` | Implemented (working-tree implementation with observed gates; pending commit) |
| R2.9 | `R2.9-type-dropdown-materialisation.md` | Authored (Tier A) |
| R2.10 | `R2.10-config-repair-migration.md` | Authored (Tier A; deferred by ADR-0010) |

### Phase 3 — Core scene engine (Tier A)

| ID | Guide | Status |
| --- | --- | --- |
| R3.1 | `R3.1-geometry-value-objects.md` | Landed (`b5314c7`) |
| R3.2 | `R3.2-scene-primitives.md` | Landed (`e244d3c`) |
| R3.3 | `R3.3-time-scale.md` | Landed (`e8a77e8`) |
| R3.4 | `R3.4-lane-stack-geometry.md` | Landed (`4ba0601`) |
| R3.5 | `R3.5-plot-frame-bands.md` | Landed (`9e088bf`, `0f577bd`) |
| R3.6 | `R3.6-span-bar-label-layout.md` | Implemented (pending commit) |
| R3.7 | `R3.7-milestone-marker-layout.md` | Implemented (pending commit) |
| R3.8 | `R3.8-critical-interval-overlay.md` | Authored (Tier A) |
| R3.9 | `R3.9-multi-event-stack-lanes.md` | Authored (Tier A) |
| R3.10 | `R3.10-delineator-lines-labels.md` | Authored (Tier A) |
| R3.11 | `R3.11-table-header-primitives.md` | Authored (Tier A) |
| R3.12 | `R3.12-scene-validator-benchmark.md` | Authored (Tier A) |
| R3.13 | `R3.13-mutation-testing.md` | Authored (Tier A; deferred by ADR-0010) |
| R3.14 | `R3.14-equivalence-thin-slice.md` | Authored (Tier A) |

### Phase 4 — live renderer (Tier B; upgrade to Tier A at Phase 3 exit)

| ID | Guide | Status |
| --- | --- | --- |
| R4.1 | `R4.1-excel-adapter-interfaces.md` | Authored (Tier B) |
| R4.2 | `R4.2-application-state-scope.md` | Authored (Tier B) |
| R4.3 | `R4.3-shape-render-conversion.md` | Authored (Tier B) |
| R4.4 | `R4.4-text-alignment-conversion.md` | Authored (Tier B) |
| R4.5 | `R4.5-polygons-z-order.md` | Authored (Tier B) |
| R4.6 | `R4.6-style-token-mapping.md` | Authored (Tier B) |
| R4.7 | `R4.7-refresh-idempotence.md` | Authored (Tier B) |
| R4.8 | `R4.8-unowned-content-preservation.md` | Authored (Tier B) |
| R4.9 | `R4.9-refresh-command.md` | Authored (Tier B; first live slice) |
| R4.10 | `R4.10-thousand-event-performance.md` | Authored (Tier B) |

### Phase 5 — interaction and UX (Tier B; upgrade at Phase 4 exit)

| ID | Guide | Status |
| --- | --- | --- |
| R5.1 | `R5.1-plot-range-modes.md` | Authored (Tier B) |
| R5.2 | `R5.2-scale-layout-choices.md` | Authored (Tier B) |
| R5.3 | `R5.3-parent-child-commands.md` | Authored (Tier B) |
| R5.4 | `R5.4-move-expand-collapse.md` | Authored (Tier B) |
| R5.5 | `R5.5-label-controls.md` | Authored (Tier B) |
| R5.6 | `R5.6-style-theme-settings.md` | Authored (Tier B) |
| R5.6a | `R5.6a-named-style-presets.md` | Authored (Tier B; added 2026-09-23) |
| R5.6b | `R5.6b-chart-title-settings.md` | Authored (Tier B; added 2026-09-25) |
| R5.7 | `R5.7-delineator-workflow.md` | Authored (Tier B) |
| R5.8 | `R5.8-warnings-panel.md` | Authored (Tier B) |
| R5.9 | `R5.9-ribbon-completion.md` | Authored (Tier B) |
| R5.10 | `R5.10-selection-overrides.md` | Authored (Tier B) |
| R5.11 | `R5.11-refresh-only-enforcement.md` | Authored (Tier B) |

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
| D-G3 | Mutation-testing tool and version | R3.13 | R3.13 guide | Open — new test-only dependency; human approval required |
| D-G4 | Pinned-font provisioning | R8.2a | R8.2a guide | Open — licence-sensitive; human choice required |
| D-G5 | Code-signing certificate procurement | by Phase 8 | R10.3 guide | Open — external lead time; start no later than Phase 8 |
| D-G6 | PowerPoint interop dependency (package + version pin) | R7.1 | R7.1 guide; ADR + `Directory.Packages.props` | Open — new production dependency; human approval required before first install |
| D-G7 | Named-style label/colour capability schema | R2.7b | ADR-0009 | Accepted 2026-09-23 |
| D-G8 | Defer R2.10 and R3.13 until after the R4.9 first-live slice | R2.9/R3.12 | ADR-0010 | Accepted 2026-09-23 |
| D-G9 | Label `Auto` cascade order and the blocked-label widest-gap truncation fallback | R3.6 | ADR-0015 | Accepted 2026-09-26 |

## Known drift resolved by revision 5

- KNOWN-LIMITATIONS L4 now names R8.4 as the first approved-PNG gate; R3.12 produces only the deterministic scene snapshot.
- The Show-Properties contract is defined by entity-guide revision 3. R5.9 still owns the separate undefined in-cell validation scope for `StyleKey`/`LabelPosition`.

## Authoring and reconciliation verification

The original 80-guide corpus was authored on 21 September 2026. Revision 5 added eight roadmap-guide records (seven R2 hardening rows and R5.6a), bringing the manifest to 88. The current local reconciliation is verified by Markdown link checking, roadmap-ID/guide existence checks, `git diff --check`, and the relevant automated test/gate commands as work items land. Office-host gates are never inferred from documentation-only authoring.
