# Work item — Audit remediation implementation plan

> **Plan status:** Approved for implementation from `docs/GanttCreator_StageInspect_Code_Audit.md`; implementation is phased, verified after every phase, and committed only after all phases are verified.
>
> This work item records the approved audit dispositions and human decisions from 2026-09-25. It is subordinate to `AGENTS.md`, the architecture, the entity guide, and the ADRs.

## Outcome

The identified R2 integrity, repair, style, scene, lifecycle, dependency, and Add Row issues are implemented or explicitly closed/deferred without silently ignoring any audit proposal, with deterministic tests and observed verification evidence after each phase.

## Approved decisions

- **Target-aware protection:** every mutating Office path authorizes the actual target worksheet after read-only target resolution; active-sheet queries remain for UI/state probes only.
- **Pristine workbook requirement:** unrelated workbook-level defined names continue to prevent sheet adoption; no weakening of the existing pristine-workbook rule.
- **Named-style resolution:** a named style supplies base formatting; per-row FillColour, StrokeColour, and LabelPosition are explicit overrides. Changing StyleKey clears all per-row formatting overrides and displays the newly selected style's complete defaults. Malformed styles fall back to the Type default with a deterministic validation/diagnostic result.
- **Event detachment:** use explicit subscription and per-handler state (`Attached`, `Detaching`, `Detached`, `DetachFailed`); do not report a failed handler removal as successful.

## Phase plan

### Phase 1 — Protection, identity, and anchor safety

- Implement target-aware protection while preserving the pristine adoption rule.
- Build reference-aware row identity repair; remap unambiguous ParentId references and refuse ambiguous duplicate-parent repairs.
- Locate plot-anchor repair by `tblGanttData`, use the live table column count, and cover suffixed/arbitrary worksheet names.
- Add positive refusal tests and Office contract coverage for non-active protected targets and repair mutations.

### Phase 2 — Configuration integrity and safe repair

- Enforce exact config key sets, reject duplicate/unexpected/blank keys, and reject duplicate preservation keys.
- Validate and repair the TypeOptions defined-name target only through the explicit repair path.
- Detect currently unused integrity findings, including schedule data on the config sheet and quarantineable user styles.
- Return distinct repair refusal reasons and count only actual successful operations.
- Track catalogue partial failures honestly; do not claim impossible Excel COM atomicity.
- Validate user style rows before preservation; never restore malformed rows into the active catalogue.

### Phase 3 — Style semantics and TypeOptions idempotence

- Implement the approved named-style base plus per-row override resolution.
- Change StyleKey resolution to clear per-row FillColour, StrokeColour, and LabelPosition when the style changes.
- Resolve malformed StyleKey values to the Type default and emit deterministic validation/diagnostic output.
- Make TypeOptions materialisation idempotent/current-checkable and avoid unconditional rewrites during Add Row.
- Keep Custom Activity capabilities style-driven and explicit.

### Phase 4 — Validation reporting and Add Row UX

- Replace broad validation-note sentinel stripping with exact structured ownership.
- Use deterministic line-oriented note formatting, truncation markers, and Unicode-safe boundaries.
- Add a separate inserted-row selection/activation concern without placing UI semantics in Core.
- Preserve Classic Notes as the compatibility reporter; defer the larger validation panel to its existing UI workstream.

### Phase 5 — Scene and geometry hardening

- Reject cyclic scene groups and duplicate group children according to the scene contract.
- Reject negative SortOrder at the scene primitive boundary.
- Reject rectangle edge overflow while preserving valid negative coordinates.
- Add snapshot and scene positive refusal tests; do not duplicate the already-valid alignment check in SceneText.

### Phase 6 — Host lifecycle and dependency boundaries

- Implement per-handler detach state and safe retry/teardown semantics without throwing into Excel.
- Wire actual repair confirmation through DestructiveCommandPolicy, including identity repair classification if confirmed.
- Remove Excel interop from Raster and remove the unused AddIn-to-Raster reference if confirmed by source usage.
- Add architecture/contract tests and live Office evidence where COM behaviour changes.

### Phase 7 — Representative acceptance fixture and final verification

- Add the audit's representative workbook/scene fixture covering overlaps, lanes, critical parentage, milestones, delineators, overrides, hidden rows, warnings, and refresh-only behaviour.
- Run targeted tests after every phase and update this document's evidence ledger.

## Exclusions

- No renderer implementation, PowerPoint transfer, PNG encoding, new Office dependency, new helper worksheet, or workbook schema expansion.
- No change to the accepted 1904 refusal or pristine-workbook adoption rule.
- No silent deletion of user-authored styles/settings and no second helper sheet.
- No Office calls in Core and no Core dependency on UI, clipboard, filesystem dialogs, Excel-DNA, or SkiaSharp.

## Acceptance tests

- Each phase has targeted unit/contract tests and positive tests for every new validator/refusal path.
- Office-dependent phases have tagged `OfficeIntegration` coverage where live behaviour is required.
- After each phase: targeted tests, Release build with warnings as errors, `pwsh ./scripts/verify-quick.ps1`, `git diff --check`, and complete diff review.
- At completion: `pwsh ./scripts/verify.ps1`, `pwsh ./scripts/verify-office.ps1` for applicable Office evidence, `git status --short`, and final commit review.
- Update `docs/STATUS.md`, `docs/REPO-MAP.md` where mapped boundaries change, and this work item's evidence ledger.

## Evidence commands

```powershell
dotnet build src/GanttCreator.slnx /p:Configuration=Release /warnaserror
pwsh ./scripts/verify-quick.ps1
pwsh ./scripts/verify.ps1
pwsh ./scripts/verify-office.ps1
git diff --check
git status --short
```

## Risk and rollback

The principal risk is changing workbook repair or style semantics while the current R2/R3 contracts are still being integrated. Each phase is limited to one concern, has positive refusal tests, and is independently reviewable. Rollback is a conventional commit revert of the final implementation commit; no migration is applied automatically to user workbooks.

## Definition of done

All implementable audit findings are resolved, all rejected/deferred findings are explicitly recorded, every new validator has a positive test, all targeted and full verification gates are observed green, applicable Office evidence is recorded, STATUS/REPO-MAP/work-item evidence is current, no unrelated user changes are overwritten, and the final implementation is committed with an applicable `docs/08-TEST-CHECKLIST.md` certification.

## Notes during implementation

- **Phase 7 observed (2026-09-25):** added `tests/GanttCreator.Core.Tests/Scene/RepresentativeAcceptanceFixtureTests.cs`, a 10-scenario fixture running one representative workbook through the real validator and scene creation, covering overlapping planned/actual spans, multiple events on one lane with distinct stacks, critical-interval parentage, two milestones, a delineator, a per-row colour override, a hidden row, determinism across repeated runs, and a canonical scene. Full non-Office suite PASS: Core 587, Office Contract 216, AddIn 195, Architecture 78, Raster 41, Office Integration (non-Office-tagged) 4.
- **Phase 5 observed (2026-09-25):** `SceneGroup` now refuses duplicate child primitive IDs and a direct self-reference; `ScenePrimitive` refuses a negative `sortOrder` (matching the row validator's `BadSortOrder`); `RectD` refuses an overflowing right/bottom edge while still allowing negative coordinates. The already-valid `SceneText` alignment check was deliberately not duplicated. Core 577/577 PASS.
- **Phase 6 observed (2026-09-25):** removed the unused `ExcelDna.Interop` package from `GanttCreator.Raster` and the unused Raster project reference from `GanttCreator.AddIn`; both claims were verified against source usage first (Raster's only file imports just `System.Globalization`, and no AddIn source names the Raster namespace). Added `EventHandlerState` and `IWorkbookStateSubscriptionStatus` with per-handler detach tracking (Decision E): a failed handler removal is now recorded `DetachFailed` and is never reported as `Detached`, while the remaining removals are still attempted. Full non-Office suite PASS: Core 577, Office Contract 216, AddIn 195, Architecture 78, Raster 41, Office Integration (non-Office-tagged) 4. Live Office evidence: not run.
- **Phase 4 observed (2026-09-25):** validation-note ownership is now line-structured rather than a substring test (`ValidationReportComposer.FindOwnedSectionStart` requires the sentinel at the start of the text or of a line); a user note that merely mentions "Gantt Creator validation:" mid-sentence is no longer truncated. Note bodies are line-oriented, truncation appends an explicit `TruncationMarker` and never splits a surrogate pair. Added AddIn-layer `IInsertedRowSelector`/`ExcelInsertedRowSelector` so a successful Add Row moves the selection to the inserted body row, with selection failure degrading silently. Classic Notes remain the compatibility reporter. Debug build PASS; Office Contract 212/212 PASS; AddIn 195/195 PASS; `verify-quick.ps1` PASS in 181.8s.
- **Phase 2 observed (2026-09-25):** exact `tblGanttConfig` key validation rejects blank/duplicate/unexpected/missing keys; catalogue preservation rejects duplicate settings/config keys with `CataloguePreservationInvalid`; TypeOptions validates the existing name target; repair counts use successful operations; visibility, anchor, and identity repair failures have distinct refusal reasons and positive contract tests. Debug solution build PASS; Office Contract 205/205 PASS; AddIn 191/191 PASS.
- **Phase 3 observed (2026-09-25):** added `src/GanttCreator.Core/GanttStyleResolution.cs` with `GanttStyleResolver` (named-style base + explicit per-row FillColour/StrokeColour/LabelPosition overrides, Type-default fallback with `UsedFallback`, typed `GanttStyleResolutionRefusal`) and `GanttStyleChange.ChangeStyleKey` (clears all per-row formatting overrides on style change); added `ITypeOptionsMaterialiser.EnsureCurrent()` and used it in the Add Row adapter; extended the architecture protection-read patterns with `QueryTarget(`. Debug solution build PASS; Core 571/571 PASS; Office Contract 205/205 PASS; `pwsh -NoProfile -File .\scripts\verify-quick.ps1` PASS in 179.4s. The resolver is not yet consumed by `GanttRowValidator`/scene construction; that wiring is Phase 7 scope.
- **Phase 1 observed (2026-09-25):** target-aware guard API and adapter wiring; identity repair builds a repair plan before writing and remaps unambiguous ParentId references; plot-anchor repair locates `tblGanttData` and uses the live table column count; pristine adoption rule preserved. `dotnet build .\GanttCreator.slnx -c Debug --no-restore` PASS; Architecture 78/78 PASS; Office Contract 199/199 PASS; `pwsh -NoProfile -File .\scripts\verify-quick.ps1` PASS in 183.3s. After the ADR-0011 alignment described below, the phase was re-verified at 182.7s.
- **ADR-0011 alignment (2026-09-25):** the first Phase 1 implementation refused identity repair with `AmbiguousParentReference`, which conflicted with accepted ADR-0011. Human decision: preserve ADR-0011. `AmbiguousParentReference` was removed, ambiguous duplicate-parent `ParentId` values are left unchanged, and a new `GanttValidationCodes.ParentAmbiguous` error reports them through normal validation. No ADR text was changed.

## Explicitly closed, rejected, or deferred audit proposals

- **P1-06 adoption relaxation:** rejected; workbook-level names remain non-pristine.
- **P1-13 Custom Activity semantics:** current style-defined capability model is preserved; expand tests only.
- **P2-03 Excel error state:** current `GanttCellState` model is preserved.
- **P2-04 1904 date system:** existing ADR-0006 refusal is preserved.
- **P2-07 validation panel:** deferred to the existing later UI workstream; Notes remain supported.
- **P2-11 snapshot alignment:** production validation already exists in `SceneText`; add targeted proof only if needed.
- **P2-16 authoring sequence:** no speculative initialisation complexity; prove through the fixture.
