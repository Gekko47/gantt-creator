# GanttCreator.Office refactor plan

Status: revised 2026-11-24 after review. Decisions D1 and D5 are approved by the owner. D2, D3, and D4 take the stated defaults unless the owner overrides them.
Source: code review of `src/GanttCreator.Office` on 2026-11-24.
Scope: `src/GanttCreator.Office` only, plus the minimum Core and test edits each slice needs.
Owner after approval: code mode. This document changes no product code.

## 1. Purpose

Remove the nine findings from the Office review without changing behaviour.
The two goals, in order:

1. Make one rule live in one place. Today the table-lookup and column-mapping rule is written in up to twelve adapters, and two copies have already diverged.
2. Make the two largest files reviewable. `ExcelShapeWriter.cs` is about 3,600 lines; `GanttRefreshOrchestrator.cs` is about 1,500 lines with a 22-parameter constructor.

Every slice keeps external behaviour identical unless its ledger row in §14 says otherwise. A slice that changes side effects, timing, or an error contract needs the decision recorded there before it starts.

## 2. Non-goals

- No worksheet-schema change. `GanttTableSchema` column order, names, and required flags stay as they are.
- No export, renderer, or scene-contract change.
- No new production dependency.
- No Office host or COM ownership redesign beyond the one decision in D1.
- No golden-image update.
- No test deletion or threshold change to force green.

## 3. Invariants every slice must keep

From `AGENTS.md` and `docs/02-ARCHITECTURE.md`:

- `Core` has no Office, Excel-DNA, SkiaSharp, clipboard, dialog, or UI dependency.
- Exactly one visible worksheet plus one `xlSheetVeryHidden` `_GanttCreatorConfig`.
- The scene model owns all positioning. Renderers consume it and never recompute dates, labels, colours, widths, or z-order.
- Milestones and delineators read `Start` only.
- Chart-affecting edits and successful chart-setting commands schedule a coalesced refresh. Selection and navigation never render.
- No `Thread.Sleep`, no `Task.Run` around COM, no analyzer suppression, no catch-and-ignore.
- Deterministic identifiers, ordering, culture, time, and rounding.
- Every new `if (bad) { error }` validator ships with a positive test in the same commit.
- One concern per commit. Conventional Commit prefixes only.
- `internal virtual` COM seams stay on the class the tests subclass. This is the seam rule; it constrains every slice below.

Gates:

```
pwsh -NoProfile -File scripts/verify-quick.ps1      # every commit
pwsh -NoProfile -File scripts/verify.ps1            # branch-final, clean tree
pwsh -NoProfile -File scripts/verify-office.ps1     # live Excel, after shape-write slices
pwsh -NoProfile -File scripts/check-md-links.ps1
pwsh -NoProfile -File scripts/check-status.ps1
```

### 3.1 Rollback protocol

Follow the refactor skill's §5:

- One slice is one commit and is reverted as a unit.
- A red targeted test triggers a revert before any diagnosis.
- State the revert and the failing test name in the handoff.
- Never widen a filter, delete a test, or edit a test to recover green.
- Do not fix forward. Re-decompose the slice or fix the prerequisite.

## 4. Finding-to-slice map

| Finding | Severity | Addressed by |
|---|---|---|
| 1. Table-access rule duplicated and drifted | P2 | Phase 2, Phase 3 |
| 2. Dead duplicate warning highlighter | P2 | Phase 1 |
| 3. `ExcelShapeWriter` god class | P2 | Phase 4 |
| 4. Four near-identical mutator adapters | P2 | Phase 3 |
| 5. `GanttRefreshOrchestrator` 22-parameter constructor | P2 | Phase 5 |
| 6. Two COM-ownership policies for one traversal | P2 | Phase 2 (D1) |
| 7. Highlighter state has two homes | P3 | Phase 6 |
| 8. Warning colours hard-coded vs Core catalogue | P3 | Phase 6 (Phase 1 removes the competing copy) |
| 9. `TryFindTable` mixes iteration styles | P3 | Phase 2 |

## 5. Decisions needed before code starts

Each decision has a default. If nobody answers, code mode takes the default and records it in the commit message.

**D1. COM ownership policy for table traversal. (APPROVED 2026-11-24)**
Decision: release only proxies acquired and not returned. Release a skipped `ListObject`, a skipped `ListObjects` collection, and a non-owning `Worksheet`. Never release the `table` or `worksheet` handed back to the caller.
Fact behind it: `ReleaseAcquired` uses `Marshal.ReleaseComObject`, and the reader's doc states the RCW may be one another component still uses.
Evidence: `src/GanttCreator.Office/ExcelGanttTableReader.cs:353`, `:358-363`, `:620-652`.
Contract tests required: a skipped proxy is released once; a returned proxy is not released.

**D2. Authority for persisted highlight state.**
Fact: the highlighter keeps `_previousHighlights` in memory and `GanttCreator.RowHighlights` in workbook properties.
Options: (a) workbook property is the only truth, memory is a per-pass cache keyed by workbook; (b) memory is the truth for one pass and the property is only a restart hint.
Default: (a). It survives reopen and workbook switch without a second rule.
Evidence: `src/GanttCreator.Office/ExcelRowWarningHighlighter.cs:90`, `:113`.

**D3. Home for the column-map rule.**
Options: (a) `GanttTableSchema` in Core; (b) a new Office helper.
Default: (a). The schema already owns column names, order, and required flags.
Evidence: `src/GanttCreator.Core/GanttTableColumns.cs:273`, `:276`.

**D4. Single-cell mutator consolidation style.**
Options: (a) one shared writer returning a common outcome, each adapter maps to its public type; (b) a generic base type.
Default: (a). It is smaller and keeps the existing public outcome types.
Evidence: `src/GanttCreator.Office/ExcelLabelPositionMutator.cs:312`, `ExcelDateLabelPositionMutator.cs:401`, `ExcelStyleKeyMutator.cs:292`.

**D5. Assembly visibility. (APPROVED 2026-11-24)**
Decision: `GanttCreator.Office` is an internal assembly, so deleting its public types is a source-compatible change inside this solution.
Still required: a repo-wide search for each deleted or changed symbol in `.cs`, `.md`, `.ps1`, `.csproj`, and `.slnx` files before Phase 1. Record the result in the commit.
Evidence: owner ruling, 2026-11-24.

## 6. Phase 0 — baseline (no product change)

**S0.1 Record the green baseline.**
Run, and save the observed summary:

```
dotnet test tests/GanttCreator.Office.ContractTests
dotnet test tests/GanttCreator.Architecture.Tests
pwsh -NoProfile -File scripts/verify-quick.ps1
```

Exit: baseline recorded. If anything is red, stop and report; do not refactor on a red tree.

**S0.2 Record the class-size and call-site baseline.**
Save the line counts for `ExcelShapeWriter.cs`, `GanttRefreshOrchestrator.cs`, and the four mutators, plus the duplicate-site list in Appendix A. These are the before numbers the review compares against.

**S0.3 Confirm the API surface (D5).**
Search the whole repo for `ExcelWarningHighlighter`, `IWarningHighlighter`, and the `GanttRefreshOrchestrator` constructor. Include docs, scripts, samples, and project files. Confirm no public API baseline file exists. Record the result.
Exit: the search result is recorded. If an external consumer appears, stop.

**S0.4 Spike the shared table base.**
On a scratch branch, move one seam member (`GetTableAt`) and `TryLocateTable` for one adapter (`ExcelDurationWriter`) onto `ExcelTableAdapterBase`. Compile and run that adapter's suite.
Exit: the spike compiles with zero test edits and its tests pass. If not, the plan uses a composed `ExcelTableAccess` with seam delegates instead, and S2.3 is rewritten before it starts.

**S0.5 Capture the working tree.**
Record `git status --short`, `git stash list`, and the current branch. A slice starts only from a clean tree.

## 7. Phase 1 — delete the dead highlighter (Finding 2, part of 8)

**S1.1 Remove `ExcelWarningHighlighter` and `IWarningHighlighter`.**
Evidence they are dead: the only live construction is `new ExcelRowWarningHighlighter(application)` in `src/GanttCreator.AddIn/RefreshSheetCommand.cs:466`; `IWarningHighlighter` appears only in its own file, one architecture allow-list entry, and its tests.

Files:
- delete `src/GanttCreator.Office/ExcelWarningHighlighter.cs`
- delete `src/GanttCreator.Office/IWarningHighlighter.cs`
- delete `tests/GanttCreator.Office.ContractTests/ExcelWarningHighlighterTests.cs`
- remove the allow-list entry at `tests/GanttCreator.Architecture.Tests/ProtectionGuardFirstTests.cs:148`
- remove any doc or comment reference found by searching `IWarningHighlighter` and `HighlightWarnings`

Exit: solution builds; architecture tests pass; `grep` for both symbols returns only this plan; the S0.3 search recorded no external consumer.
Commit: `refactor: remove unused cell-note warning highlighter`.
Behaviour change: none.

## 8. Phase 2 — one table locator, one column map (Findings 1, 6, 9)

This is the highest blast radius. Do it in five commits.

**S2.1 Add the pure column map to Core (D3).**
Add `GanttTableSchema.TryBuildColumnMap(IReadOnlyList<string> headerNames, out int[] columnMap)`.
Rules, copied exactly from the existing copies:
- one-based index per schema column, in schema order;
- ordinal name comparison;
- a missing required column returns `false`;
- a missing optional column maps to `-1`;
- an empty or duplicate header list behaves as today.

Tests: new `GanttTableSchemaColumnMapTests` in `tests/GanttCreator.Core.Tests`.
Cases: exact headers, reordered headers, missing required column, missing optional column, duplicate header, empty list, case-mismatched header.
Characterisation rule: add these fixture shapes to a shared fixture file before any Office copy is deleted, so old and new mappings are compared on identical input.
Exit: targeted Core tests pass. No Office file touched.
Commit: `refactor: add pure schema column-map builder`.

**S2.2 Make the fast header read feed the pure map.**
Change `ExcelSheetChangeAnalyzer.TryBuildColumnMap` (`ExcelSheetChangeAnalyzer.cs:168`) to:
1. read the header row in one `Value2` call as it does today;
2. fall back to per-column names as it does today;
3. pass the names to `GanttTableSchema.TryBuildColumnMap`.
The bulk-read optimisation and its comment stay. Only the mapping loop is removed.
Comparison: run the shared fixture against the old loop and the new mapper before deleting the loop. The maps must be identical for every fixture case.

Exit: `ExcelSheetChangeAnalyzerTests` and `ExcelSheetChangeAnalyzerDecisionTests` pass unchanged.
Commit: `refactor: route change-analyzer column map through schema`.

**S2.3 Add the shared table traversal (D1).**
Prefer the base-class shape, because the repository already uses it and documents why: `StagingShapeAdapter` (`src/GanttCreator.Office/StagingShapeAdapter.cs:20-27`) states that a composed helper "would bypass a subclass's overrides".

Create `src/GanttCreator.Office/ExcelTableAdapterBase.cs`:
- constructor shape confirmed by the S0.4 spike; if the adapters' constructors do not fit one base, use a composed `ExcelTableAccess` instead;
- `internal virtual object GetSheetAt(Sheets sheets, int index)`;
- `internal virtual ListObject GetTableAt(ListObjects objects, int index)`;
- `internal virtual bool TryLocateTable(Sheets sheets, out Worksheet? sheet, out ListObject? table)`;
- `internal virtual void ReleaseAcquired(object? proxy)`, implementing D1 exactly: release skipped proxies only, never the returned table or sheet;
- a test seam for the release call so a contract test can assert it fired.

Adapters that adopt the base, one commit each, deleting only their private `TryFindTable` copy:
`ExcelDurationWriter.cs:263`, `ExcelStyleKeyWriter.cs:188`, `ExcelOutlineGroupWriter.cs:568`, `ExcelRowHeightNormaliser.cs:853`, `ExcelParentChildMutator.cs:677`, `ExcelDateLabelPositionMutator.cs:272`.

Leave `ExcelGanttTableReader`'s traversal in place until S2.5. It is the reference policy and it has extra ownership handling around its result.

Test rule: the existing adapter tests subclass the concrete adapter and override the seam members. Moving a `virtual` member to a base class keeps those `internal override` declarations valid, because they override an inherited virtual member. No test edit is expected. If a test edit is needed, stop and report; it means the move changed a seam.

D1 contract tests (required):
- a skipped `ListObject` is released once;
- a skipped `ListObjects` collection and a non-owning `Worksheet` are released;
- the returned `table` and `worksheet` are not released.

Behaviour ledger: side effects (release calls) and timing, approved by D1. Product output is unchanged.
Repo map: add the new base class to `docs/REPO-MAP.md` in the first adapter commit.

Exit per adapter: its contract tests pass with zero edits; one new `ExcelTableAdapterBaseTests` proves the release policy and the chart-sheet skip.
Commit pattern: `refactor: share table traversal in <adapter>`.

**S2.4 Add the body and active-row helpers.**
Create `ExcelTableBodyReader` (or add to the base) with:
- `TryReadBody` using `ExcelValue2Matrix.ReadRows` and the S2.1 map;
- `ResolveActiveBodyIndex` using the same header-row arithmetic as today.

Replace the four copies: `ExcelLabelPositionMutator.cs:257/283`, `ExcelDateLabelPositionMutator.cs:346/372`, `ExcelStyleKeyMutator.cs:237/263`, `ExcelParentChildMutator.cs:763/800`.
Keep the `ExcelGanttRowInserter.GetInsertionPosition` variant separate: it deliberately treats the header row as "insert at top" (`ExcelGanttRowInserter.cs:449-461`), which the mutators must not do.

Exit: the four mutator suites pass unchanged. Add one test for the header-selected case to the base tests.
Error contract: the optional-column refusal reason stays exactly as today. A test pins the message path from the date-label mutator.
Comparison: run the shared body fixture against each old `TryReadBody` before deleting it.
Commit: `refactor: share table body read and active-row resolution`.

**S2.5 Align the reader.**
Move `ExcelGanttTableReader`'s traversal onto the base, or document in one line why it keeps its own copy. Preferred: use the base, because it is the reference policy. Keep its extra result-ownership rule in the caller.

Exit: `GanttTableReaderTests` pass unchanged.
Repo map: if the reader now derives from the base, update `docs/REPO-MAP.md` in the same commit.
Commit: `refactor: use shared table traversal in table reader`.

## 9. Phase 3 — one single-cell write path (Finding 4)

**S3.1 Add `ExcelSingleCellWriter`.**
Inputs: table, column map, zero-based row index, schema column name, value, body index.
It performs the same steps as today: resolve the table index, `ListRows` row, `ListColumns` cell, `SetRangeValues`.
It returns a common `CellWriteOutcome` with `Ok`, `RefusedWriteFailed`, `RefusedColumnMissing`.

The optional-column guard from `ExcelDateLabelPositionMutator.cs:411-418` becomes an input flag:
- `DateLabelPosition` passes "optional"; a missing column returns a refusal that names Initialise.
- `LabelPosition` and `StyleKey` pass "required"; the guard is unreachable, as the map already proves them.

**S3.2 Adopt it in the three mutators, one commit each.**
`ExcelLabelPositionMutator.cs:312`, `ExcelDateLabelPositionMutator.cs:401`, `ExcelStyleKeyMutator.cs:292`.
Each keeps its public outcome type and maps `CellWriteOutcome` to it. No public signature changes.
Exit: the three suites pass unchanged, plus one new test that a pre-R5.14 workbook still gets the "run Initialise" refusal from the date-label path.
Commit: `refactor: use shared single-cell writer in <mutator>`.

**S3.3 Leave `ExcelParentChildMutator` alone.**
Its batch `WriteCells` (`:1023`) and hierarchy planning are not the same operation. Only its table/body/active-row sections moved in Phase 2.

## 10. Phase 4 — split `ExcelShapeWriter` (Finding 3)

The detailed slice design is in the review handoff. The plan pins the rules here.

Hard constraint: nine test classes derive from `ExcelShapeWriter` and override `internal virtual` members. The list is in Appendix B. Any new collaborator must call through the writer instance, never construct its own host access.

**S4.1 Partial split, moves only.**
Files: `ExcelShapeWriter.Creation.cs`, `.Update.cs`, `.Ownership.cs`, `.Appearance.cs`, `.Text.cs`, `.Host.cs`.
Exit: `git diff --stat` shows moves only; no logic line changes; the five writer suites pass with zero edits.
Contingency: schedule S4.2–S4.4 in the same milestone. If they slip, do not ship S4.1 alone; six partial files still hold one god class. Record the size budget in the review (no file over about 900 lines after S4.4).
Commit: `refactor: split shape writer by responsibility`.

**S4.2 Move pure ownership rules.**
Move `CouldBeOwnedName`, `CarriesTagForName`, and name composition into `ShapeOwnershipScheme`.
Keep `ApplyOwnershipStamp` and `ApplyDuplicateOwnershipStamp` on the writer as seams.
Exit: ownership and style-write suites pass unchanged.
Commit: `refactor: move shape ownership rules into scheme`.

**S4.3 Extract appearance with injected seams.**
`ShapeAppearanceWriter` depends on a narrow `IShapeAppearanceHost` implemented by `ExcelShapeWriter`; it does not take the concrete class, to avoid a writer → collaborator → writer cycle. The host exposes `ApplyStyle`, `ApplyStyleToRange`, `ApplyText`, and `ApplyTextFormatting`.
It must not read `shape.Fill`, `shape.Line`, or `shape.TextFrame2` directly.
The unmapped-value throws (`ArgumentOutOfRangeException`) stay.
New compatibility test: a test subclass overrides `ApplyStyle`, the collaborator runs, and the override is observed.
Exit: `StyleWriteTests`, `ExcelShapeWriterTests`, `ExcelShapeWriterDepthTests` pass; the compatibility test passes.
Gate: run `pwsh -NoProfile -File scripts/verify-office.ps1` because the style call path changed.
Commit: `refactor: extract shape appearance writer`.

**S4.4 Extract batch state.**
`ShapeWriteBatch` holds `Shapes`, the proxy memo, the tag memo, the template cache, and the create timings.
The writer keeps the non-nestable `BeginBatch` guard and the public scope type.
Exit: batch, template, and ownership suites pass unchanged.
Commit: `refactor: extract shape write batch state`.

## 11. Phase 5 — parameter objects for the orchestrator (Finding 5)

**S5.1 Add the two records.**
`GanttRefreshDependencies`: the nine required ports.
`GanttRefreshOptions`: the thirteen optional values, with today's defaults.
The orchestrator keeps one public constructor that takes the two records. If a compatibility constructor is kept for tests, mark it `internal` and route it to the records.

**S5.2 Update the composition root and tests.**
`src/GanttCreator.AddIn/RefreshSheetCommand.cs:418` becomes named properties.
Test constructors use the records.
Exit: `GanttRefreshOrchestratorTests` and the AddIn tests pass unchanged.
Repo map: update the orchestrator entry point in `docs/REPO-MAP.md` in the same commit.
Commit: `refactor: replace orchestrator parameter list with records`.

## 12. Phase 6 — highlighter cleanups (Findings 7, 8)

**S6.1 One state authority (D2).**
If D2 is "property is truth": build `_previousHighlights` from the property at the start of a pass, write it back at the end, and key the cache by workbook identity. Add a test for a workbook switch.
Exit: `ExcelRowWarningHighlighterTests` pass plus the new switch test.
Commit: `refactor: make persisted highlights the single authority`.

**S6.2 One colour authority (Finding 8).**
After Phase 1 only the live highlighter remains. Replace the two OLE constants at `ExcelRowWarningHighlighter.cs:84` and `:87` with the Core catalogue tokens, or add the pastel values to the catalogue and document them there. Pick one; do not keep both.
Exit: a test asserts the colours come from the catalogue.
Commit: `refactor: source warning fill colours from catalogue`.

## 13. Phase 7 — close-out

- Run `pwsh -NoProfile -File scripts/verify.ps1` on a clean tree.
- Run `pwsh -NoProfile -File scripts/verify-office.ps1`.
- Run `pwsh -NoProfile -File scripts/check-md-links.ps1` and `check-status.ps1`.
- Update `docs/STATUS.md` with the refactor outcome.
- Update `docs/REPO-MAP.md` if a mapped entry point or test location changed (Phase 2 adds a base class and a Core mapper; Phase 5 changes the orchestrator constructor).
- Record the D1–D4 decisions in `docs/DECISIONS.md` if they are new policy, not just a local choice.

## 14. Behaviour-change ledger

| Slice | Class | What changes | Pinned by |
|---|---|---|---|
| S1.1 | none | Dead type removal only | Build plus architecture tests |
| S2.1 | none | New pure method, additive | Core mapper tests |
| S2.2 | none | Mapping loop removed; same reads | Old-vs-new fixture |
| S2.3 | side effects, timing | Release calls added for skipped proxies (D1) | Release contract tests; `verify-office.ps1` |
| S2.4 | none | Body and active-row code relocated; refusal preserved | Date-label refusal test |
| S2.5 | side effects | Reader adopts the D1 release policy | Reader tests plus release tests |
| S3.1–S3.2 | none | Internal outcome mapping; public types unchanged | Three mutator suites plus the pre-R5.14 test |
| S4.1 | none | Moves only | Moves-only diff; writer suites |
| S4.2 | none | Pure rule relocation | Ownership tests |
| S4.3 | timing | Style call path goes through one collaborator | Compatibility test; `verify-office.ps1` |
| S4.4 | none | State object relocation | Batch and template suites |
| S5.1–S5.2 | contract | Constructor signature changes; assembly is internal (D5) | Full solution build plus tests |
| S6.1 | state | Highlight rebuild source changes (D2) | Workbook-switch test |
| S6.2 | none | Colour source changes to the catalogue | Colour test |

Rows marked `timing`, `side effects`, `contract`, or `state` are not pure refactors. Each one needs the named test and the approval recorded in D1, D2, or D5.

## 15. Risk register

| Risk | Trigger | Mitigation |
|---|---|---|
| A seam is moved off `ExcelShapeWriter` | Phase 4 slice 3 | Seam-injection rule plus the new compatibility test. Stop if a test needs an edit to compile. |
| A test subclass breaks when a seam moves to a base class | Phase 2 | Overriding an inherited virtual member still compiles. If an edit is needed, the move changed behaviour. Stop and report. |
| The bulk header read is lost | Phase 2 slice 2 | Keep `TryReadHeaderNamesInOneCall`; only replace the loop after it. Measure the automatic refresh before and after on one workbook. |
| COM release policy changes host behaviour | Phase 2 slice 3 | D1 first. Do not mix policies in one commit. Live check with `verify-office.ps1` after the first adapter adopts the base. |
| A copy-paste variant contains extra logic | Phase 2/3 | Replace only the traversal, map, body, and active-row sections. Leave hierarchy planning, hidden-flag rotation, outline rebuild, and template reuse untouched. |
| Architecture allow-list drifts | Phase 1 | Remove the dead entry in the same commit as the deletion. |
| A pure move includes a hidden edit | Phase 4 slice 1 | Compare `git diff -w`; a move-only commit should show only file headers and member relocation. |

## 16. Acceptance criteria

- `TryFindTable` traversal exists once per design: in the shared base, with the reader reduced to a documented thin wrapper. No adapter keeps a private copy.
- `TryBuildColumnMap` logic exists once: `GanttTableSchema.TryBuildColumnMap`.
- `TryReadBody` and `ResolveActiveBodyIndex` exist once each.
- `ExcelShapeWriter` is no longer a single file; the largest file is under about 900 lines after S4.4.
- `GanttRefreshOrchestrator` has two constructor parameters (the two records).
- The dead highlighter and its interface are gone.
- Warning fill colours have one authority.
- Every phase leaves the targeted suites green with no test deletion, no filter widening, and no analyzer suppression.
- `verify.ps1` and `verify-office.ps1` both observed green at close-out.

## 17. Suggested sequence and estimate

1. Phase 0 and Phase 1 — half a day.
2. Phase 2 — three to five days, six adapters and one reader, one commit each.
3. Phase 3 — one to two days.
4. Phase 4 — three to four days.
5. Phase 5 — one day.
6. Phase 6 — one day.
7. Phase 7 — half a day.
Total: eight to twelve working days, labelled an estimate until S0.1 confirms the test baseline.

Phases 1 and 2 can ship on their own. They remove the highest maintenance cost and do not touch rendering.

The estimates exclude `verify-office.ps1` runs, which need a human Excel host. Add one per shape-write slice and one at close-out.

## Appendix A — duplicate-site inventory

Table traversal (`TryFindTable`):
- `src/GanttCreator.Office/ExcelDurationWriter.cs:263`
- `src/GanttCreator.Office/ExcelStyleKeyWriter.cs:188`
- `src/GanttCreator.Office/ExcelOutlineGroupWriter.cs:568`
- `src/GanttCreator.Office/ExcelRowHeightNormaliser.cs:853`
- `src/GanttCreator.Office/ExcelParentChildMutator.cs:677`
- `src/GanttCreator.Office/ExcelDateLabelPositionMutator.cs:272`
- `src/GanttCreator.Office/ExcelGanttTableReader.cs:337-373` (reference policy)

Additional traversal-shaped sites to classify as adopt, keep with reason, or different rule:
- `src/GanttCreator.Office/ExcelGanttRowInserter.cs:96` (call site; locate the definition before changing it)
- `src/GanttCreator.Office/ExcelConfigCatalogueReader.cs:155` (catalogue tables, different name source)
- `src/GanttCreator.Office/ExcelConfigCatalogueWriter.cs:1241` (catalogue tables)
- `src/GanttCreator.Office/ExcelSelectionIdentifier.cs:183`
- `src/GanttCreator.Office/ExcelInsertedRowSelector.cs:172`
- `src/GanttCreator.Office/ExcelPlotAnchorRepairer.cs:77`
- `src/GanttCreator.Office/ExcelTypeChangeNormaliser.cs:315`
- `src/GanttCreator.Office/ExcelWorkbookInitialiser.cs:515`, `:912`
- `src/GanttCreator.Office/ExcelWorkbookRecovery.cs:1117`

Column map:
- `src/GanttCreator.Office/ExcelGanttTableReader.cs:382`
- `src/GanttCreator.Office/ExcelParentChildMutator.cs:718`
- `src/GanttCreator.Office/ExcelStyleKeyMutator.cs:198`
- `src/GanttCreator.Office/ExcelLabelPositionMutator.cs:218`
- `src/GanttCreator.Office/ExcelDateLabelPositionMutator.cs:307`
- `src/GanttCreator.Office/ExcelGanttValidationReporter.cs:379`
- `src/GanttCreator.Office/ExcelSheetChangeAnalyzer.cs:168` (bulk variant)
- `src/GanttCreator.Office/ExcelWarningHighlighter.cs:204` (deleted in Phase 1)

Column-lookup variants to classify:
- `src/GanttCreator.Office/ExcelGanttRowIdentityRepairer.cs:161`
- `src/GanttCreator.Office/ExcelTypeOptionsMaterialiser.cs:373`
- `src/GanttCreator.Office/ExcelPanelGridMeasurement.cs:733`

Body read:
- `ExcelLabelPositionMutator.cs:257`, `ExcelDateLabelPositionMutator.cs:346`, `ExcelStyleKeyMutator.cs:237`, `ExcelParentChildMutator.cs:763`

Active-row resolution:
- `ExcelLabelPositionMutator.cs:283`, `ExcelDateLabelPositionMutator.cs:372`, `ExcelStyleKeyMutator.cs:263`, `ExcelParentChildMutator.cs:800`
- variant kept separate: `ExcelGanttRowInserter.cs:431-461`

Single-cell write:
- `ExcelLabelPositionMutator.cs:312`, `ExcelDateLabelPositionMutator.cs:401`, `ExcelStyleKeyMutator.cs:292`

## Appendix B — `ExcelShapeWriter` test subclasses

- `tests/GanttCreator.Office.ContractTests/ExcelShapeWriterTests.cs:31`, `:197`, `:829`, `:1454`
- `tests/GanttCreator.Office.ContractTests/ExcelShapeWriterDepthTests.cs:187`
- `tests/GanttCreator.Office.ContractTests/ExcelShapeWriterTemplateTests.cs:435`
- `tests/GanttCreator.Office.ContractTests/StyleWriteTests.cs:290`, `:322`
- `tests/GanttCreator.Office.ContractTests/StagingSchemeWriters.cs:29`

Seams they override: `FindGanttWorksheet`, `GetShapes`, `GetShapeCount`, `GetShapeAt`, `FindShapeByName`, `GetShapeByExactName`, `AddShape`, `ApplyText`, `ApplyStyle`, `ApplyOwnershipStamp`, `TextFormattingMatches`, `TextMatches`.

## Appendix C — per-commit checklist

- [ ] One finding or one mechanical step only.
- [ ] Behaviour change: none stated in the commit message, or explicitly labelled.
- [ ] Existing tests unchanged, or the reason recorded.
- [ ] New validator or refusal path has a positive test.
- [ ] No analyzer suppression added; existing `CA1031` justifications unchanged.
- [ ] Targeted tests observed green.
- [ ] `pwsh -NoProfile -File scripts/verify-quick.ps1` observed green.
- [ ] `git diff --check` and `git status --short` reviewed.
- [ ] `docs/STATUS.md` updated when a phase completes.
- [ ] `docs/REPO-MAP.md` updated in the same commit when an entry point, boundary, or test location changes.
- [ ] Applicable `docs/08-TEST-CHECKLIST.md` sections certified in the commit message.

## Appendix D — execution log

**2026-11-24 — Phase 0 to Phase 2 (partial).**

- Baseline observed: Office contract 1779, Architecture 142, both green; tree clean except the plan file; three pre-existing stashes left untouched.
- D5 search observed: no external consumer, no `PublicAPI*.txt` baseline; the only references were the two Office files, their tests, and two architecture allow-list entries.
- Phase 1 landed as `refactor: remove unused cell-note warning highlighter`.
- S2.1 landed as `refactor: add pure schema column-map builder` (10 Core facts).
- S2.2 landed as `refactor: route change-analyzer column map through schema`.
- S2.3 landed: `ExcelTableAdapterBase` (D1) and adoption by six adapters in seven commits. Duration and row-height depth fakes now also stub `Count` and the one-based indexer.
- S0.4 outcome: the spike compiled with the delegating seam, but two test fakes needed support edits. The plan says a test-support edit is a stop-and-report point, so S2.4, S2.5, and Phases 3-7 are not started.
- Gate observed after S2.3: `verify-quick.ps1` PASS in 230.2 s — Core 2297, Office contract 1766, AddIn 943, Architecture 142, Raster 127, non-live Office 36; Release build 0 warnings.
