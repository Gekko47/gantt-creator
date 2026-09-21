# Implementation guides manifest and execution plan

> Authored 21 September 2026 on branch `cline/vd5dq9wz` from `stage-inspect`
> commit `8d53d39`. This manifest records the product owner's approved
> two-tier plan: every remaining roadmap work item — 80 of the 100 rows in
> [`../03-ROADMAP.md`](../03-ROADMAP.md) (R0.1–R0.8, R1.1–R1.6, R2.1–R2.6 are
> landed) — gets an individual implementation guide in this directory.
> The guides are the pre-created work-item files the roadmap requires
> ("Create a work-item file before implementation"); none of them records
> implementation state. [`../STATUS.md`](../STATUS.md) remains the record of
> what has actually landed, and `AGENTS.md` retains requirement precedence.

## Tier model (approved 2026-09-21)

- **Tier A — prescriptive.** 19 guides for the Phase 2 remainder
  (R2.7–R2.10) and Phase 3 (R3.1–R3.14), verified against the source tree as
  of commit `8d53d39`. They name exact files, types, seams, test files,
  `[Fact]` names, and evidence commands.
- **Tier B — binding contract.** 61 guides for Phases 4–10. They fix the
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
above. Phase 3's guides are Tier A from initial authoring.

## Index of guides (80)

One row per remaining roadmap row, in roadmap order. Status values:
`Authored` (file exists; tier as stated), `Pending` (next authoring batch).
Filenames are backticked, not linked, so link checking stays on authored
files only; the authoring audit verifies every `Authored` row's file exists.

### Phase 2 remainder — configuration and data commands (Tier A)

| ID | Guide | Status |
| --- | --- | --- |
| R2.7 | `R2.7-config-catalogues.md` | Authored (Tier A) |
| R2.7a | `R2.7a-destructive-command-policy.md` | Authored (Tier A) |
| R2.8 | `R2.8-add-row-commands.md` | Authored (Tier A) |
| R2.9 | `R2.9-type-dropdown-materialisation.md` | Authored (Tier A) |
| R2.10 | `R2.10-config-repair-migration.md` | Authored (Tier A) |

### Phase 3 — Core scene engine (Tier A)

| ID | Guide | Status |
| --- | --- | --- |
| R3.1 | `R3.1-geometry-value-objects.md` | Authored (Tier A) |
| R3.2 | `R3.2-scene-primitives.md` | Authored (Tier A) |
| R3.3 | `R3.3-time-scale.md` | Authored (Tier A) |
| R3.4 | `R3.4-lane-stack-geometry.md` | Authored (Tier A) |
| R3.5 | `R3.5-plot-frame-bands.md` | Authored (Tier A) |
| R3.6 | `R3.6-span-bar-label-layout.md` | Authored (Tier A) |
| R3.7 | `R3.7-milestone-marker-layout.md` | Authored (Tier A) |
| R3.8 | `R3.8-critical-interval-overlay.md` | Authored (Tier A) |
| R3.9 | `R3.9-multi-event-stack-lanes.md` | Authored (Tier A) |
| R3.10 | `R3.10-delineator-lines-labels.md` | Authored (Tier A) |
| R3.11 | `R3.11-table-header-primitives.md` | Authored (Tier A) |
| R3.12 | `R3.12-scene-validator-benchmark.md` | Authored (Tier A) |
| R3.13 | `R3.13-mutation-testing.md` | Authored (Tier A) |
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
| R4.9 | `R4.9-refresh-command.md` | Authored (Tier B) |
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

### Phase 8 — raster fidelity and golden images (Tier B; upgrade at Phase 7 exit)

| ID | Guide | Status |
| --- | --- | --- |
| R8.1 | `R8.1-width-parsing-residual.md` | Pending — Batch 7 |
| R8.2 | `R8.2-export-dialog.md` | Pending — Batch 7 |
| R8.2a | `R8.2a-font-pinning-adr.md` | Pending — Batch 7 |
| R8.3 | `R8.3-cross-renderer-equivalence.md` | Pending — Batch 7 |
| R8.4 | `R8.4-golden-images.md` | Pending — Batch 7 |
| R8.5 | `R8.5-hires-render-verification.md` | Pending — Batch 7 |
| R8.6 | `R8.6-slow-machine-verification.md` | Pending — Batch 7 |
| R8.7 | `R8.7-zoom-dpi-evidence.md` | Pending — Batch 7 |
| R8.8 | `R8.8-phase-exit.md` | Pending — Batch 7 |

### Phase 9 — startup, performance, hardening, packaging (Tier B; upgrade at Phase 8 exit)

| ID | Guide | Status |
| --- | --- | --- |
| R9.1 | `R9.1-startup-open-refresh.md` | Pending — Batch 8 |
| R9.2 | `R9.2-performance-budgets.md` | Pending — Batch 8 |
| R9.3 | `R9.3-failure-hardening.md` | Pending — Batch 8 |
| R9.4 | `R9.4-workbook-save-open-hardening.md` | Pending — Batch 8 |
| R9.5 | `R9.5-installer-packaging.md` | Pending — Batch 8 |
| R9.6 | `R9.6-addin-performance-instrumentation.md` | Pending — Batch 8 |
| R9.7 | `R9.7-matrix-perf-stability.md` | Pending — Batch 8 |
| R9.8 | `R9.8-phase-exit.md` | Pending — Batch 8 |

### Phase 10 — release (Tier B; upgrade at Phase 9 exit)

| ID | Guide | Status |
| --- | --- | --- |
| R10.1 | `R10.1-supported-office-matrix.md` | Pending — Batch 9 |
| R10.2 | `R10.2-user-docs.md` | Pending — Batch 9 |
| R10.3 | `R10.3-code-signing.md` | Pending — Batch 9 |
| R10.4 | `R10.4-install-instructions.md` | Pending — Batch 9 |
| R10.5 | `R10.5-full-suite-run.md` | Pending — Batch 9 |
| R10.6 | `R10.6-exploratory-workflow.md` | Pending — Batch 9 |
| R10.7 | `R10.7-sbom-freeze.md` | Pending — Batch 9 |
| R10.8 | `R10.8-release-candidate.md` | Pending — Batch 9 |

## Decision register (product decisions a guide may not make alone)

| ID | Decision | First needed by | Where recorded | Status |
| --- | --- | --- | --- | --- |
| D-G1 | Configuration-sheet table set, settings storage, and storage format (R2.7 ADR) | R2.7 | R2.7 guide decisions D1–D3; ADR-0007 draft | Open — options + recommendation; human approval before first write |
| D-G2 | Destructive-command undo semantics (R2.7a ADR) | R2.7a | R2.7a guide; ADR-0008 draft | Open — human choice required |
| D-G3 | Mutation-testing tool and version | R3.13 | R3.13 guide | Open — new test-only dependency; human approval required |
| D-G4 | Pinned-font provisioning | R8.2a | R8.2a guide | Open — licence-sensitive; human choice required |
| D-G5 | Code-signing certificate procurement | by Phase 8 | R10.3 guide | Open — external lead time; start no later than Phase 8 |
| D-G6 | PowerPoint interop dependency (package + version pin) | R7.1 | R7.1 guide; ADR + `Directory.Packages.props` | Open — new production dependency; human approval required before first install |

## Known drift to resolve during authoring

- [`../KNOWN-LIMITATIONS.md`](../KNOWN-LIMITATIONS.md) L4 still names R3.12
  as the first approved-PNG gate; roadmap revision 4 moved that obligation
  to R8.4. Re-date L4 when KNOWN-LIMITATIONS is next revised.
- The roadmap R5.9 note calls the Show-Properties control undefined in the
  entity guide; entity guide revision 3 "Required worksheet fields" defines
  it. The R5.9 guide resolves which text is current; escalate only if a
  real gap remains.

## Authoring verification

Authored in a Linux sandbox without `pwsh` or `dotnet`. Per-batch checks
actually run: link/reference existence for every authored file (bash
equivalent of `scripts/check-md-links.ps1` rules), `git diff --check`, and
this manifest's `Authored` rows verified to exist on disk.
`scripts/verify-quick.ps1` and the Pester script gates were **Not run
locally** — they run in CI on the pushed branch. Office gates are by
definition Not run at authoring time.


