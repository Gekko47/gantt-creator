
# Implementation roadmap — revision 6

<!-- SKILL-SUMMARY:START -->
Phase-by-phase work-item list (R0.x .. R10.8), the cross-phase
compatibility matrix, and the scope-change protocol.

Do not get wrong:
- One roadmap row is normally one commit; split a row if the diff
  gets hard to review, never combine unrelated rows.
- Every acceptance criterion naming a test count, command, or
  artifact must link to a specific test or script step — drift
  between this doc and the code is itself a defect
  (docs/08-TEST-CHECKLIST.md section I).
- A phase exits only on its stated automated demonstration. A
  screenshot supports evidence; it never substitutes for it.
<!-- SKILL-SUMMARY:END -->

<!-- SKILL-TOOLS:START -->
- `github__list_commits` / `git_tool` — verify "one row is one commit" and Conventional Commit prefixes.
- `github__get_pull_request_status` — verify green CI before declaring a phase exit (the W-12 rule).
- `github__get_pull_request_files` — check the diff links to a specific test or script step.
- `dotnet_test` — run the named test or script step an acceptance criterion references.
- `combined-mcp-server__read_file` on `docs/work-items/` — confirm a work-item file exists before implementation.
- `memra_add` — record one work-item completion fact per landed item (3-5 lines: ID, commit, gates, Office status, next item). Not commit text or test counts — see `docs/06-LLM-PROTOCOL.md` "Persisting the ledger".
- `memra_add_decision` — record an irreversible design decision with the context that made it irreversible.
- `memra_bootstrap` — recall prior decisions at session start before restating unknowns.
<!-- SKILL-TOOLS:END -->

> **R3 composition remediation (28 September 2026, `docs/REMEDIATION-PLAN-A.md`)** — An external review of the R3.1–R3.13 composition and the R4.1–R4.5 adapters produced a remediation brief. Every claim was re-verified against source before acting and three were found incomplete or wrong; the verified plan is `docs/REMEDIATION-PLAN-A.md` and the brief itself is **not** a tracked source of requirements. Commit A is done (ADR-0021, ADR-0022) and closes the R3-side items: `Splitter`/`Spacer` reach lane layout for the first time, the §10 band is implemented, "empty scene" means "no visible scene-producing entity", and the delineator style has one authority. **No roadmap row is added, removed, renumbered, or re-scoped** — this is corrective work inside already-approved rows, and the row count is unchanged. The remaining items were the per-row `PanelCellGrid` and one panel-bounds authority (blocking R4.6+), the `ApplyZOrder` ownership preflight (blocking R4.7), and the R4.6/R4.7 sequencing the R4.7 work item already records. **All three are now closed** (remediation commits B, C, and D: ADR-0023 and ADR-0024 for the per-row measurement and the single panel-bounds authority, and the `ApplyZOrder` preflight, recorded in the later sections of [the repository map](REPO-MAP.md) and in R4.7's Step 0). **The golden scene snapshot was regenerated and approved in a separate human-reviewed commit** (`4278a49`, plan and verified delta in `docs/REMEDIATION-PLAN-GOLDEN-AND-D.md`), then regenerated once more after the label-suppression fix (`4c29675`); the D4 policy that regeneration be its own reviewed commit was followed both times.

> **Revision 10 (29 September 2026) — pre-R4.9 hierarchy, projection and orchestration.** Two tracked input documents, [`GanttCreator_R3_R4_Architecture_Roadmap_Review_REV5.md`](GanttCreator_R3_R4_Architecture_Roadmap_Review_REV5.md) and [`GanttCreator_Revised_Implementation_Plan_REV6.md`](GanttCreator_Revised_Implementation_Plan_REV6.md), propose the corrections the product model needs before Refresh can be locked: real parent/child rows, projected child shapes, Critical Interval as a user-authored child, uniform worksheet geometry, hidden engine metadata, Ribbon Row Styling, derived Duration, and a Refresh orchestrator. **They are governance inputs, not the contract** — ADR-0025 places them in `AGENTS.md`'s source-of-truth order below the entity guide, and [`PRE-R4.9-DECISION-REGISTER.md`](PRE-R4.9-DECISION-REGISTER.md) records which of their proposals were changed or rejected, because four of them conflict with landed behaviour and a fifth was implemented wrongly by an earlier draft of the plan. **This revision inserts nine rows, R4.7A–R4.7H and R4.8A, all of which must complete before R4.9.** No existing row is removed, renumbered, or re-scoped; this revision contains **124** commit-sized work items, counted from the phase tables. Revision 9's stated 111 came from an earlier count that did not match the tables, so this figure is measured rather than inherited — the discrepancy predates this revision and is recorded here rather than silently absorbed.
>
> Three of the nine reverse landed behaviour and therefore carry ADRs. **R4.7D** removes live lane auto-growth, deleting `LaneLayoutBuilder`'s `Math.Max(LaneHeightPt, contentHeight)` and the `08-TEST-CHECKLIST.md` §B invariant that states the same rule, because a projected child must share its parent's lane and lane geometry has to flow from the worksheet into Core (ADR-0026). **R4.7E** keeps the Critical Interval a **filled rectangle** — the guide says every host renderer draws a line from the rect, and zero of the three do, so the reinterpretation is removed rather than implemented in three places; `CriticalFill` carries `#FF0000` and `CriticalLinePt` is retired (ADR-0027). **R4.7G** removes `Above`/`Below` and *amends* ADR-0015 rather than superseding it, because the existing truncation fallback is confirmed correct and is reused as built (ADR-0028). **R4.7C** advances the workbook schema 4 → 5 in place, with no migration (R4.7A already advanced it 3 → 4 for `SiblingOrder`), per the owner's confirmation that there are no active users (ADR-0029). Two of the nine — R4.7B and R4.7D — change scene geometry and so require human-reviewed golden regenerations; **CORRECTED 2026-09-30: this revision originally claimed R4.7E "does not" need a golden regeneration "because the committed golden contains no critical overlay primitive". That was FALSE — the snapshot does contain one.** The premise that let the divergence survive was never "the golden is silent" but "no renderer implemented the guide rule". In the event R4.7E geometry was pulled forward into commit `017199f` (owner ruling 2026-09-30: the critical interval is a parent-independent span half the predetermined `ActivityHeightPt`), it DID change the snapshot, and that regeneration shipped as its own commit `806a632`. Each regeneration ships in the same commit as its code change, per R3.17's recorded D3.
>
> **R4.9 becomes integration and must invent no architecture.** It is the row that locks the Refresh sequence, so every contract it depends on lands first: R4.7A → R4.7B → R4.7D are strictly sequential; R4.7C, R4.7F, R4.7G and R4.7H are independent of one another; R4.7E needs R4.7B and R4.7C. R4.8A needs all eight (R4.7A–R4.7H). Cross-phase rows are rescoped rather than added — R5.3 loses the parent/child mechanics to R4.7A and keeps the permitted-Type selector and delete/promote UI, R5.2 hands the size-preset model to R4.7H, R5.5 rescopes to the horizontal-only label set, and R6.2's export panel becomes the collapsed top-level rows while all child graphics still render.

> Created 2 September 2026 under a new filename. Revision 9 (26 September 2026) inserts one row, `R3.16`, which gives the entity guide's per-entity **field** table its own commit. R3.14 could only *assemble* its three equivalence field lists, because the guide's "Entity-to-renderer equivalence" section states which representation each entity takes in each renderer, not which fields it carries — so the contract a renderer author must satisfy was assembled inside a test file rather than stated in the guide. Three gaps make the amendment necessary before Phase 4: the critical interval is a `SceneRect` one `CriticalLinePt` tall in the scene but a *line* in all three host renderers and the table never says so; two rows (data panel/header, validation indicator) draw no live host object at all and a reader cannot tell that from the current text; and seven of ten rows have no field-level proof. The row changes **no token value, z-layer, or label position** — it names only fields the landed model already carries, and any cell that would need a new field reads "not carried by the model; excluded". The optional legend (§25) becomes an eleventh row marked empty rather than an omission. The row is blocked on product-owner approval of the field lists before the guide is edited. No row is removed, renumbered, or re-scoped; this revision contains 111 commit-sized work items. Revision 8 (26 September 2026) inserts one row, `R2.7d`, which carries the approved event-date display format as a workbook setting. ADR-0016 names the value (`dd/mm/yyyy`) because entity guide §3 and §23 referred to "the approved display format" without specifying one, which armed R3.11's own stop condition: the twelve settings approved by ADR-0007 D3 contain no date-format key, and `PeriodLabelFormat` selects the period-header band label only. The row advances the workbook schema version 2 → 3 because ADR-0007 D3 makes the exact settings key set a schema contract. No row is removed, renumbered, or re-scoped; `R2.7d` lands ahead of R3.11 rather than inside it because it changes the schema contract R3.11 reads. This revision contains 110 commit-sized work items. Revision 7 (26 September 2026) records ADR-0015: the label `Auto` cascade order and the blocked-label outcome. Span `Auto` becomes `Right → Left → Inside`; the delay event's `Inside` is a style-level default evaluated once rather than a special case; and a label no candidate can accommodate takes the widest free gap, truncated with a single-character `…` and one warning. No roadmap row is added, removed, renumbered, or re-scoped by this revision — the work-item count is unchanged and every ID keeps its meaning — because ADR-0015 changes the contract R3.6–R3.8 implement, not the set of rows. Revision 6 (25 September 2026) inserts one row, `R2.7c`, which gives the Core style registry's resolved formatting its own commit. The StageInspect audit landed `GanttStyleResolver` as prerequisite work, but that resolver could only read `GanttCatalogues.StylePresets`, so a user-authored style was invisible to resolution — a gap against R2.7b's save/reopen gate and R5.6a's Custom Activity Refresh gate. R3.14 requires resolved style tokens in its equivalence field list, but it is blocked behind unimplemented R3.6–R3.12, so the registry payload lands first as its own tracked row rather than being folded into a test row. Revision 5 (23 September 2026) applied the approved R2 findings implementation plan: eight rows were inserted (`R2.1a`, `R2.2a`, `R2.4a`, `R2.4b`, `R2.5a`, `R2.6a`, `R2.7b`, `R5.6a`), ADR-0009 defines named-style capability columns, and ADR-0010 records the first-live-slice sequencing exception that moves R2.10 and R3.13 after R4.9 without removing or weakening either row. This revision contains 109 commit-sized work items. No existing ID was renumbered, reused, or removed, so work-item files, STATUS, and ADR references keep resolving. Revision 4 (21 September 2026) remains the historical basis for all unchanged rows and gates.

## How to use this roadmap

- Complete items in order unless an approved ADR records why order changed. Order is table order; row IDs are stable historical labels.
- Row IDs are never renumbered or reused. A row inserted between two existing rows carries a letter suffix (for example `R2.7a`) so work-item files, STATUS, and ADR references keep resolving.
- One row is normally one commit. Split a row if the diff becomes difficult to review; do not combine rows merely because they are related.
- Create a work-item file before implementation. Record the exact acceptance tests and evidence.
- Every work-item acceptance criterion that names a test count, a command, or an artifact must link to a specific test or script step. Drift between the doc and the code is a defect (see `docs/08-TEST-CHECKLIST.md` section I).
- Every visual/layout/style/export work item names the affected sections of `docs/07-GANTT-ENTITY-GUIDE.md` and tests the defined cross-renderer contract.
- A row flagged *landed early* in a phase note was already satisfied by an earlier row's commit; implement only the residual the note describes, and link the acceptance criteria to the existing tests.
- A phase note's *picks up* clause lists deferrals recorded in earlier work items that the row owes; implementing the row without them is drift.
- Every code commit must pass `scripts/verify-quick.ps1`; every phase exit and pull request must pass `scripts/verify.ps1`.
- A phase exits only after its stated demonstration. A screenshot is supporting evidence, not a substitute for automated assertions.
- Use synthetic construction data. Do not add customer schedule data to tests or examples.

The **Automated gate** is run from the terminal and CI. A **Visual Studio / Office gate** marked `Required` must be demonstrated against desktop Excel or PowerPoint on Windows, normally launched or debugged through Visual Studio. `None` means the commit does not require Office; it does not waive the automated gate. A work item cannot be marked done when a required Office gate was not run.

A Required Office gate is either a **runner gate** — an assertion expressed as an `OfficeIntegration` test wired into `scripts/verify-office.ps1` on the self-hosted runner, following the R2.2/R2.4/R2.6 pattern of Moq contract tests plus a tagged live test — or a **human gate**, which needs a Visual Studio/manual session because it exercises judgement. Every Required gate that can be written as an assertion is a runner gate. A phase exits only when every item-level Required Office gate in the phase is recorded as run in `docs/STATUS.md` — a PASS, or a defect with an owner; pending gates are listed there explicitly rather than left implicit. Office-gate evidence records the host Windows and Office build every time, so the R10.1 supported matrix is assembled from phase-exit records instead of being discovered at the end.

## Phase 0 — repository and quality foundation

Goal: a clean solution that fails fast on warnings, formatting, dependency drift, and broken tests.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R0.1 | Add governance kit, licence, security policy, and contribution entry point | Validate Markdown links, rule discovery, and skill schemas | None |
| R0.2 | Add `global.json`, solution, project folders, and allowed project references | Release build and architecture reference test | Required: open the solution and confirm every project and x64 configuration loads |
| R0.3 | Add central package management, locked restore, and approved initial packages | Locked restore succeeds twice from clean package state | None |
| R0.4 | Add `.editorconfig`, `Directory.Build.props`, analyzers, nullable, and warnings-as-errors | Format check plus a temporary warning that must fail before removal | None |
| R0.5 | Add xUnit test projects and one non-trivial sample test per test project | `dotnet test` discovers and passes every test project | Required: Test Explorer discovers the same test projects |
| R0.6 | Add quick/full verification scripts and coverage settings | Both scripts pass from a clean clone and enforce configured thresholds | None |
| R0.7 | Add Windows CI for restore, format, build, unit tests, and artifacts | Branch workflow passes and a deliberate failure produces a useful annotation | None |
| R0.8 | Add versioning and local rolling-log abstractions with privacy-safe defaults | Unit tests for version string, rotation, and redaction | None |

> **R0.8 note (do not skip):** while adding the rolling-log abstractions,
> also close the two tooling gaps deferred as **L6** and **L7** in
> `docs/KNOWN-LIMITATIONS.md` — lint the GitHub Actions workflow
> (actionlint or equivalent, pinned) and add unit tests for the PowerShell
> scripts in `scripts/` (Pester or equivalent, wired into
> `verify-quick.ps1`). Phase 0 does not exit until L6 and L7 are closed or
> re-dated by an approved decision.

Exit demonstration: a new developer clones the repository, runs one documented command, and gets a clean Release build and test report without opening Office.

## Phase 1 — Excel-DNA host and Ribbon shell

Goal: reliably load, debug, and unload the add-in before product behaviour is added.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R1.1 | Add Excel-DNA 1.9 entry point with `AutoOpen`/`AutoClose` logging | Entry-point/log contract tests and Release build | Required: F5 loads the correct x64 XLL; `AutoOpen` and `AutoClose` breakpoints hit |
| R1.2 | Add minimal valid RibbonX resource with a Gantt Creator tab | Ribbon XML namespace, IDs, and callback contract tests | Required: Excel displays one Gantt Creator tab without Ribbon errors |
| R1.3 | Add diagnostics command showing add-in/Office/bitness identifiers | Callback/application-command unit tests | Required: click Diagnostics and hit the callback breakpoint in Visual Studio |
| R1.4 | Add one command error boundary with operation IDs and user-safe messages | A thrown fake command produces one log record and one translated result | Required: forced Excel callback failure shows one safe message and retains usability |
| R1.5 | Add Ribbon state service and invalidate mechanism | State getters are deterministic and side-effect-free | Required: controls enable/disable correctly as workbook state changes |
| R1.6 | Add deterministic add-in shutdown and owned-resource cleanup | Lifecycle and cleanup contract tests | Required: repeat Excel open/close five times with no add-in error or owned orphan process |

Exit demonstration: F5 launches Excel, the Ribbon appears, Diagnostics works offline, a breakpoint is hit, and forced failure produces one safe dialog and one useful log record.

## Phase 2 — worksheet contract and data access

Goal: create/read the visible single-sheet data model without rendering.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R2.1 | Define table column names, event types, and schema version in Core | Enum, schema, and serialization round-trip tests | None |
| R2.1a | Pin stable Type machine identity separately from the 16 persisted display labels; clarify selectable Types versus visual guide sections | Exact enum name/value pins plus existing catalogue/display/schema tests | None |
| R2.2 | Implement `Initialise Sheet` to bring a workbook into the supported state: adopt the blank active worksheet as `tblGanttData`'s home, or create one, plus `_GanttCreatorConfig` | Adapter tests assert exact names, the adopted-blank path leaves one visible sheet and one `xlSheetVeryHidden` sheet and no others, the create path leaves two visible sheets (the preserved original and the new Gantt Data) and one `xlSheetVeryHidden` sheet, and typed refusals mutate nothing | Required: initialise a blank workbook and inspect table, plot anchor, helper visibility, and sheet count |
| R2.2a | Restrict worksheet adoption to pristine sheets; preserve cell-empty sheets containing non-cell user state and create a new Gantt sheet | Contract matrix for A1-only used range, shapes, comments/threaded comments, names, tables, pivots, queries, and hyperlinks proves preserved original plus exact create mutation set | Required: create a blank cell-only sheet with real shape/comment/name state and prove it remains unchanged |
| R2.3 | Add stable ID generation and preservation independent of row number | Insert, sort, move, and delete contract tests | Required: sort and insert rows in Excel; IDs remain stable and unique |
| R2.4 | Read cell values into neutral row DTOs without locale display parsing | 1900 date-system and multiple-culture conversion tests | Required: read representative real Excel dates under both supported locale formats |
| R2.4a | Isolate date-system support/refusal behind an Office-free converter while retaining ADR-0006's 1904 rejection | Converter support/serial tests and reader refusal-before-body tests | Required: existing live 1904 refusal remains deterministic |
| R2.4b | Preserve blank/value/Excel-error/unsupported states in neutral typed DTO cells | Known and future Excel error mapping, unsupported typed-value tests, and live error-cell round-trip | Required: real Excel error cells remain distinguishable from blanks |
| R2.5 | Map DTOs into Core events with all-errors validation | Table-driven valid, invalid, and all-errors tests | None |
| R2.5a | Make validation the deterministic cell-state normalizer and require Critical Interval parents to be valid mapped spans | Relevant/irrelevant cell-state matrix, new-validator positives, warning-event inclusion, and deterministic cross-row tests | None |
| R2.6 | Add row-level error reporting without mutating valid input | Fake-worksheet mutation and error-order tests | Required: invalid rows show actionable errors while original cells remain unchanged |
| R2.6a | Keep Notes a replaceable reporter and enforce protection-guard-first note mutation | Zero-mutation protected refusal, ownership preservation, and architecture mutation-discovery tests | Required: protected-sheet Validate changes no user or owned notes |
| R2.7 | Persist versioned workbook settings and style/metric tables on `_GanttCreatorConfig` | Settings/style/metric round-trip plus invalid/missing schema tests | Required: save/reopen and confirm the preserved-original-sheet-or-Gantt-Data visible sheet rule, one VeryHidden helper, and retained settings |
| R2.7a | Add the destructive-command policy ADR: confirmation, undo-or-no-undo semantics, and protected-sheet typed refusals | Policy contract tests for every mutating command: confirmation before delete, the approved undo semantics, and a typed refusal with no partial mutation when the target sheet or workbook is protected, extending R2.2's `TargetProtected` pattern | Required: exercise one confirmed destructive command and one protected-sheet refusal in Excel |
| R2.7b | Materialise each named style's default/allowed label positions and colour-override capability | Built-in projection, user-row round-trip, capability validation, and catalogue-hash tests | Required: save/reopen preserves custom style capabilities and hash |
| R2.7c | Carry each named style's resolved formatting and default label position in the Core style registry so resolution covers user styles, not only built-in presets | Registry-construction guard tests, built-in and user-style projection tests, resolver override-precedence and Type-default-fallback tests, and a round-trip proving a user style's formatting survives a writer/reader regeneration and resolves identically from a rebuilt registry (`A_user_style_resolves_identically_after_a_registry_rebuild`) | None |
| R2.7d | Add the approved `DateDisplayFormat` setting and advance the workbook schema version | Settings key-set, key-order, default-value, strict-parser refusal, and `en-GB`/`en-US` culture-invariance tests; schema-version pin | None |
| R2.8 | Add add-activity, add-milestone, and add-delineator row commands | Command/contract tests for defaults, IDs, and insertion position | Required: invoke all three Ribbon commands and inspect the resulting visible rows |
| R2.9 | Materialise the central Type catalogue and apply its named-range dropdown | Catalogue uniqueness/hash, type/style/date/capability mapping, defined-name, and validation tests | Required: full-name dropdown appears on existing/new rows and resolves only through `_GanttCreatorConfig` |
| R2.10 | Add VeryHidden configuration integrity, migration, and safe-repair workflow | Missing/corrupt/wrong-visibility/version/hash tests preserve valid custom styles or require confirmation | Required: damage/copy/remove configuration in synthetic workbooks and verify detection, repair, and the preserved-original-sheet-or-Gantt-Data visible sheet rule |

> **Phase 2 notes (revision 5):**
> - **R2.1a** pins the durable enum identity separately from persisted display labels; it adds no Type and no localisation.
> - **R2.2a** narrows adoption to pristine worksheets. A cell-empty sheet with non-cell artefacts is preserved and the create path is used.
> - **R2.4a/R2.4b** isolate date-system policy and preserve neutral Excel error/unsupported cell states; ADR-0006 keeps 1904 rejected.
> - **R2.5a/R2.6a** make the validator the deterministic normalizer and Notes a guarded, replaceable reporter.
> - **R2.7** must enumerate the complete first-release settings/style/metric catalogue from `docs/07-GANTT-ENTITY-GUIDE.md` before the first write. **R2.7b**, under ADR-0009, adds explicit named-style label/colour capabilities required by Custom Activity and R2.9.
> - **R2.7a** precedes R2.8 because R2.8's row commands are the first mutating commands after Initialise. COM-driven edits do not enter Excel's undo stack.
> - **R2.8** adds row insertion with catalogue defaults only. The delineator dialog, edit, delete, and ignored-field warnings are **R5.7**.
> - **R2.9** picks up the R2.5 deferrals: style-registry existence checks and the `Custom Activity` label/colour capability gates, now made representable by R2.7b.
> - **R2.10** picks up the R2.2 plot-anchor and R2.3 identity-repair deferrals. ADR-0010 moves it after R4.9; it remains mandatory and still requires its own approved safe-repair ADR before code.

Exit demonstration: initialise a blank workbook, enter representative data, use the full-name Type dropdown, validate it, and save/reopen it. Prove there is exactly one visible Gantt worksheet and one valid `_GanttCreatorConfig` worksheet with `xlSheetVeryHidden` visibility and no schedule data.

## Phase 3 — deterministic Core scene engine

Goal: generate the complete point-based drawing model with no Office process.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R3.1 | Add point, size, rectangle, colour, and tolerance value objects | Boundary, equality, conversion, and invalid-number tests | None |
| R3.2 | Add immutable scene primitives, stable IDs, groups, and z-order | Serialization, ID, grouping, and deterministic-order tests | None |
| R3.3 | Add time-range validation and day-to-point mapping policy | Leap-day, boundary, monotonicity, and finish-policy tests | None |
| R3.4 | Add lane ordering, row height, and stack geometry | Property tests prove shuffled input produces the same geometry | None |
| R3.5 | Add plot frame, monthly/yearly header bands, and alternating time bands | Exact scene-geometry snapshots | None |
| R3.6 | Add span-event bar and label layout with clipping | Before, within, crossing, and after-range geometry tests | None |
| R3.7 | Add milestone marker and label layout | Same-date, boundary, size, and label-side tests | None |
| R3.8 | Add critical child interval overlay | Multiple disjoint, adjacent, and overlapping interval tests | None |
| R3.9 | Add multiple events on one lane using `StackIndex` | Two/three-event stack, gap, overflow, and collision tests | None |
| R3.10 | Add full-height delineator lines and labels | Plot-height, z-order, clipping, and duplicate-date tests | None |
| R3.11 | Add visible table/header scene primitives for editable export | Bounds, cell, grid, and text-style snapshots | None |
| R3.12 | Add scene invariant validator and representative 1,000-event benchmark | Zero invalid geometry and recorded benchmark under the Core budget | None |
| R3.13 | Add the mutation-testing harness for changed Core code (the R0.5 deferral) | A pinned mutation tool runs the Core suite through a new `scripts/verify.ps1` step with a Pester guard; the first full-Core run is archived as the baseline; the 80% changed-code threshold of `docs/04-TEST-STRATEGY.md` applies from the next Core change onward | None |
| R3.14 | Prove the scene model with a cross-renderer primitive-equivalence thin slice | A contract test walks one bar, one external label, and one milestone diamond through the built scene and asserts every field the entity-guide equivalence table requires of all four renderers; a missing field is a defect, not a renderer workaround. A gap that serialization alone cannot prove stops for a spike-plus-ADR under the scope-change protocol, with the spike kept outside production code | None |
| R3.15 | Make the scene's declared chart bounds the single source of truth | The scene's `ChartBounds` equals the `chart:background` primitive it contains, and a caller cannot supply a second, contradictory chart bounds; the now-unreachable `PlotOutsideChart` guard is deleted rather than left vacuous | None |
| R3.16 | State the per-entity equivalence **field** table in the entity guide | Every equivalence row names the scene primitive, its role-derived ID, its owner, and every field a renderer must read, transcribed from the landed model with no token change; a guard test fails if the guide and the test catalogue disagree | None |
| R3.17 | Make one header-label identifier convention | A year label names its parent by appending `:label` to the parent's ID, matching the period header and the row description label, and a convention test proves every `:label` text names a rectangle that exists | None |

> **Phase 3 notes (revision 5):**
> - **R3.1** is landed early: `PointD` already exists in Core with tests. The row completes the value-object set — size, rectangle, colour, tolerance — beside it; it must not re-implement or diverge from `PointD`.
> - **R3.3** picks up the plot-range policy groundwork deferred from R2.5 (decision U11). **R5.1** later adds the explicit range modes and the user-facing warnings.
> - **R3.11** header primitives follow the schema display names in `GanttTableSchema.Default`. `Duration` is not a schema v1 column; do not invent one. A derived `Duration` column requires a product-owner entity-guide amendment plus a schema ADR first.
> - **R3.12**'s artifact is the deterministic scene snapshot and the benchmark only. No PNG renderer exists until Phase 8, so the golden-PNG limitation (L4 in `docs/KNOWN-LIMITATIONS.md`) cannot close at this phase; its replacement gate is the first approved representative render (R8.4). Re-date L4 when `docs/KNOWN-LIMITATIONS.md` is next revised.

Exit demonstration: a command-line/test fixture produces a deterministic scene snapshot containing planned/actual overlaps, three events on one lane, critical segments, milestones, and two labelled delineators.

## Phase 4 — live native Excel renderer

Goal: render and refresh an owned set of Excel shapes beside the source data.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R4.1 | Add narrow Excel application/workbook/worksheet/shape adapter interfaces | Architecture test keeps Core reference-clean; adapter contracts compile | None |
| R4.2 | Add application-state scope with restoration after success/failure | Fake-state tests cover every property and injected failure point | Required: force a live command failure and confirm Excel state is restored |
| R4.3 | Render basic rectangles/lines with explicit point geometry and ownership tags | Shape type, name, tag, and point-geometry contract tests | Required: inspect real Excel shape properties within the documented tolerance |
| R4.4 | Render text, font, alignment, and overflow policy | Text-frame, font, alignment, and overflow contract tests | Required: inspect representative labels at normal and clipped lengths |
| R4.5 | Render polygons/milestones and deterministic z-order | Shape count/type/order and milestone geometry tests | Required: inspect diamonds and overlaps in Excel, including same-date events |
| R4.6 | Map all scene styles including planned, actual, baseline, critical, and hatch patterns | Complete style-token-to-Office-property contract matrix | Required: compare all live styles against the approved visual specification |
| R4.7 | Add idempotent refresh and owned-shape reconciliation | Two refreshes produce identical owned IDs/counts/bounds | Required: refresh twice and confirm no duplicates, flicker defect, or selection corruption |
| R4.7A | Make `ParentId` authoritative and add engine-maintained `SiblingOrder`; assign identity on first meaningful entry; enforce the seven-child maximum; promote children on parent delete; reseed a cloned managed structure atomically | Hierarchy validator reports every broken, ambiguous, cyclic, or over-deep relationship; paste and clone tests prove new identities and preserved relationships; `IsSemanticallyEmptyRow` has one definition reused everywhere | Required: create a parent and seven children, refuse the eighth, collapse and expand, delete the parent and observe the children promoted |
| R4.7B | Introduce `EntityProjection` (`OwnLane`, `OverlayParent`, `StackOnParent`, `MilestoneOnParent`, `PlotGlobal`, structural) with a Core capability matrix, and separate `SourceRowVisible` from `RenderVisible` | Projection is resolved in Core before scene construction; a collapsed child renders with `SourceRowVisible = false`; Excel row-hidden is never mapped to `GanttEvent.Visible` | None |
| R4.7C | Classify every managed column; hide and lock the engine columns; add visible locked `Duration` and hidden `SiblingOrder`; add the `GanttRowHeightPt` token; advance the schema 4 → 5 in place with no migration (R4.7A advanced it 3 → 4 for `SiblingOrder`) | Column-classification contract; hidden/locked state is repaired if a user unhides it; config-integrity reports a version mismatch; `CriticalLinePt` absent, `CriticalFill` present | Required: confirm on a protected sheet that the engine columns and `Duration` reject edits while the authoring cells accept them, and that the version bump is detected |
| R4.7D | Derive live lane geometry from the measured worksheet row; normalise managed row heights to `GanttRowHeightPt`; **remove live lane auto-growth**; rebuild outline groups from `ParentId`; project children onto the parent lane | Every visible lane-owning row's lane Top/Height/Bottom equal its Excel row's; `LaneLayoutBuilder` no longer grows a lane; a projected child creates no lane; **human-reviewed golden regeneration** | Required: expand and collapse a parent with children and confirm the lanes stay aligned with the visible rows in both states |
| R4.7E | Make Critical Interval a real projected child with independent dates and permitted styling; render it as a **filled rectangle** at half the predetermined `ActivityHeightPt`, **centred on its own visual slot** and derived from its **own** dates (owner ruling 2026-09-30 — it is **not** top-aligned to the parent, and its parent governs lane membership only); retire `CriticalLinePt` | `CriticalInterval` gains the `Fill` capability and the validator accepts it; the catalogue, validator and checklist change together; multiple intervals project onto one parent; the R3.16 field row is re-transcribed; the integration fixture carries a critical child | Required: style a critical child and read the fill back off the live host |
| R4.7F | Calculate `Duration` in Core from date-only semantics and write it back in one bulk pass | Inclusive span days correct for one-day and multi-day; milestones and delineators show the approved marker; structural rows blank; invalid dates produce no misleading value and a validation report | Required: edit dates and confirm `Duration` updates only on Refresh |
| R4.7G | Remove `Above`/`Below` from the enum, the capability sets, the planner's milestone order, the guide and the checklist; keep the delineator corners and splitter positions | The retained horizontal set resolves; the removed positions no longer parse; the ADR-0015 truncation and nothing-to-place outcomes still fire unchanged | Required: confirm a label placed to the right of a bar in Excel |
| R4.7H | Introduce first-class output size presets (A4, Presentation 16:9, Presentation 4:3) with a single plot-geometry authority | Text columns keep their measured widths; `PlotWidth = PresetWidth − TextPanelWidth − Margins`; title and time bands use the same composition geometry; insufficient remaining width **refuses** with an actionable error | Required: select each preset and confirm the plot reflows without resizing the text columns |
| R4.8 | Preserve unowned shapes/cells and active worksheet | Sentinel shape/cell and active-sheet adapter tests | Required: refresh with manual content present and prove it is unchanged |
| R4.8A | Add the production Refresh orchestrator and a testable `SceneBuildRequestFactory`; resolve plot bounds from preset, text-panel measurement, date range and scale; enforce the mutation-ownership boundary; make live-versus-export composition an explicit profile | `RefreshSheetCommand` stays thin and only checks availability, invokes the orchestrator, and presents the outcome; every validation, layout and scene step succeeds before any shape mutation; `LiveExcel` supplies no data panel | Required: exercise a full Refresh on a representative construction-delay workbook and record the host build |
| R4.9 | Wire Validate and Refresh Ribbon commands with guarded errors | Command, validation, and error-boundary tests | Required: exercise Validate/Refresh on valid and invalid workbooks |
| R4.10 | Profile 1,000-event refresh and remove only measured bottlenecks | Automated benchmark and no-regression threshold | Required: measure real Excel refresh on the reference machine and record Office build |
| R4.11 | Propagate label text colour through style resolution, the scene model, and the Office adapter, activating §17's inside/outside switch | Resolver, planner, scene, snapshot, mapper, and renderer tests; a non-delay control proves the switch is not widened | Required: read a label's font colour back off the live host on create and after an update |

> **Phase 4 notes (revision 10):**
> - **R4.9** is the first-live-slice demonstration and the Phase 3 guide-upgrade gate. Blocking Refresh preserves the previous valid scene; normal worksheet/property edits never invoke rendering. **As of revision 10 it is also the gate that locks the Refresh sequence**, and every contract it depends on lands first: R4.7A-R4.7H and R4.8A all complete before it starts. R4.9 **integrates settled contracts and invents no architecture**; if it turns out to need a new contract, that is a stop condition, not a row to absorb.
> - **R4.7A → R4.7B → R4.7D are strictly sequential.** Identity feeds projection, and projection feeds measured geometry: a child must have a resolved render lane before lane geometry can be derived from the worksheet rather than from content. **R4.7C, R4.7F, R4.7G and R4.7H are independent of one another.** R4.7E needs R4.7B and R4.7C. **R4.8A needs all eight — R4.7A, R4.7B, R4.7C, R4.7D, R4.7E, R4.7F, R4.7G and R4.7H.** This text previously said "all seven", which under-counted the R4.7 rows by one and left a reader unable to tell which item had been dropped; **no item is excluded.** R4.8A composes a `SceneBuildRequest`, so it needs identity and hierarchy (A), projection (B), the classified schema (C), measured lane geometry and the outline groups (D), the critical-interval contract (E), derived `Duration` (F), the horizontal label set (G), and the single plot-geometry authority (H). Each of the eight is named because R4.8A assembles all of them and a reader implementing it must be able to check its own prerequisites against that list.
> - **Two rows require human-reviewed golden regenerations: R4.7B and R4.7D.** Each needs a stated reason and human review, and each ships **in the same commit as its code change** — splitting them would leave the intermediate commit red, per R3.17's recorded D3. R4.7E did not regenerate the golden *in the colour-capability change*: it adds a critical child to the **vertical integration fixture** rather than the canonical fixture, which is what lets the representation be pinned without another regeneration. **Corrected 2026-09-30:** the parenthetical here once read "the committed scene contains no critical overlay primitive", which was **false** — it contained one, and the geometry change in `017199f` resurrected a second that had been silently deleted for falling outside its parent, so `806a632` regenerated it. The no-regeneration claim now applies only to the later fill-contract commit.
> - **R4.7A owns hierarchy mechanics; R5.3 owns the user-facing UI for it.** The Add Child workflow, promote-on-delete and hierarchy repair land in R4.7A; R5.3 keeps the permitted child-Type selector and the delete/promote Ribbon surface over them. R5.4 (move up/down, expand/collapse) is retargeted at the outline groups R4.7D rebuilds.
> - From this phase onward, every Office evidence row records the host Windows/Office build, feeding the R10.1 matrix.

> **Phase 4 notes (revision 5):**
> - **R4.1** is landed early in part through the existing application/workbook/table/config/reporter ports. The row adds only missing narrow shape/render ports and must not duplicate them.
> - **R4.9** is the first-live-slice demonstration and the Phase 3 guide-upgrade gate. Blocking Refresh preserves the previous valid scene; normal worksheet/property edits never invoke rendering.
> - ADR-0010 moves **R2.10**, then **R3.13**, immediately after R4.9. **R4.10** follows them and profiles both 1,000-event scene/refresh and the existing 1,000-row Validate/report path.
> - From this phase onward, every Office evidence row records the host Windows/Office build, feeding the R10.1 matrix.

Exit demonstration: after R4.9, a valid workbook with the canonical construction-delay fixture renders a real owned Gantt beside `tblGanttData`; a blocking Refresh leaves it unchanged, and a second successful Refresh is idempotent. R2.10 and R3.13 then close their mandatory deferred gates before R4.10 performance evidence.

## Phase 5 — chart controls and construction-delay cases

Goal: complete the main authoring workflow and the screenshot-equivalent Ribbon controls.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R5.1 | Add plot start/finish settings with automatic and explicit modes | Invalid, automatic, explicit, and boundary setting tests | Required: change both modes in Excel and confirm chart range and Ribbon state |
| R5.2 | Add month/quarter/year scale and page-width layout choices | Geometry snapshots for every scale/layout choice | Required: inspect headers, bands, and page-width behaviour in Excel |
| R5.3 | Add parent/child commands and stable lane reassignment | Sort, move, parent, child, and stable-ID contract tests | Required: exercise parent/child commands and refresh the real worksheet |
| R5.4 | Add move-up/down and expand/collapse display commands | Deterministic ordering and visibility tests | Required: exercise controls and confirm visible rows/shapes remain aligned |
| R5.5 | Add label visibility/position controls | Scene snapshots and Type-specific allowed-position tests for every option | Required: inspect label placement, clipping, and retained settings after Refresh |
| R5.6 | Add style/theme settings using explicit local defaults | Style round-trip, validation, and offline-dependency tests | Required: edit colours/heights/diamond size and compare Excel output with approved tokens |
| R5.6a | Create, rename, and delete approved user named-style presets with explicit label/colour capabilities | Unique-key, capability, built-in protection, cancel-before-write, save/reopen, and Custom Activity consumption tests | Required: create a custom style, assign it to Custom Activity, Refresh, and compare the live result |
| R5.6b | Add user-editable chart title and title-visibility settings | Nonblank-title, exact-settings-write, cancel/no-mutation, save/reopen, and Refresh-only tests | Required: edit title, toggle visibility, save/reopen, and confirm the chart changes only after Refresh |
| R5.7 | Add Start-only delineator create/edit/delete workflow | Type/dialog creation, ignored-field warning, same-date, label, visibility, edit, and delete tests | Required: create from dropdown and Add Vertical Line dialog; chart changes only after Refresh |
| R5.8 | Add warnings panel/dialog for clipped, missing, and conflicting data | Deterministic grouping, severity, and message tests | Required: provoke each warning in Excel and confirm it does not mutate data |
| R5.9 | Complete Ribbon layout, icons, keytips, accessibility labels, and offline help | Ribbon XML/callback, resource, and offline-link contract tests | Required: keyboard, screen-reader label, high-contrast, and screenshot-layout review |
| R5.10 | Add single-expanded-entity selection and row fill/line/label overrides | Selection-context, capability, colour, label-position, inheritance, and persistence tests | Required: controls enable only for one expanded entity; edit rectangle/diamond fill and label position |
| R5.11 | Enforce Refresh-only rendering for all worksheet and property edits | Command/event tests prove edits never invoke the renderer; blocking Refresh preserves the last scene | Required: edit Type/dates/colour/label, confirm shapes stay unchanged, then Refresh once to apply all |

> **Phase 5 notes (revision 10):**
> - **R5.2** keeps the month/quarter/year **scale** selector. The **size-preset
> model** (A4, 16:9, 4:3) and the single plot-geometry authority move to R4.7H,
> because Refresh derives plot bounds from the preset and R4.8A needs that
> authority to exist before it can assemble a `SceneBuildRequest`. R5.2 consumes it.
> - **R5.3** is rescoped by R4.7A. Add Child, promote-on-delete, hierarchy repair
> and `SiblingOrder` are R4.7A's mechanics; R5.3 keeps the **permitted child-Type
> selector and the delete/promote Ribbon surface** over them, plus the
> capability matrix's user-facing side.
> - **R5.4** is retargeted at the Excel outline groups R4.7D rebuilds from
> `ParentId`, and at `SiblingOrder`-based ordering rather than raw row number, so
> a user sort cannot silently redefine the hierarchy.
> - **R5.5** is rescoped by R4.7G to the horizontal-only label set: `None`,
> `Auto`, `Left`, `Right`, and `Inside` (shown as Centre), plus the delineator
> corners. `Above`/`Below` are removed and must not appear as options.
> - **R5.10** operates on **row-based selection**: the user selects a worksheet
> row, the command resolves it to the hidden stable ID, and Row Styling styles
> only that entity. A collapsed parent gives group context for navigation and
> expand/collapse, but it is **not** a styleable group — Row Styling must never
> implicitly style children.

> **Phase 5 notes (revision 4):**
> - **R5.7** extends R2.8's add-delineator row command; it does not replace it. R2.8 is row insertion with defaults; R5.7 is the dialog, edit, delete, and ignored-field warnings.
> - **R5.8** extends R2.6's cell notes and counts-only summary; it covers scene-level warnings — clipping and the plot-range warnings deferred from R2.5 — and the deterministic warning list. It must not duplicate or re-surface R2.6's row-level notes.
> - **R5.6b** owns only the user-editable `ChartTitle`/`ShowTitle` settings surface approved by ADR-0014. It reuses R5.6's approved settings-dialog infrastructure where compatible, writes only those two settings, and preserves the Refresh-only rendering rule.
> - **R5.9** scope check: `docs/07-GANTT-ENTITY-GUIDE.md` defines a dropdown only for the `Type` column. The show/hide property-columns control and any in-cell validation for `StyleKey`/`LabelPosition` are undefined there; each needs a product-owner decision before this row can complete. Until then the Validate command is the only guard for those columns.

Exit demonstration: a user can select Types from the worksheet dropdown, edit one expanded entity's colour and label position, and add a Start-only delineator. No edit changes the chart until Refresh is clicked; one Refresh then produces the representative delay visual with overlapping planned/actual/critical activity, multiple events on a lane, milestones, and labelled vertical lines.

## Phase 6 — editable composition and clipboard

Goal: copy the whole selected Gantt as one editable Office shape group.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R6.7 | Run the copy-after-delete/delayed-rendering compatibility spike — first in this phase, ahead of the composition rows it de-risks | Spike records deterministic format/shape assertions and proposed ADR outcome | Required: test copy-after-cleanup on every supported Excel/PowerPoint build |
| R6.1 | Add export-selection and bounds model without creating shapes | Core tests for chart-only and table-plus-chart bounds | None |
| R6.2 | Compose table/header cells as native rectangle/text shapes | Shape count, bounds, cell text, style, and z-order contract tests | Required: inspect a real staged table composition in Excel |
| R6.3 | Compose complete chart scene at a staging origin | Contract proves group bounds equal scene bounds within tolerance | Required: stage beside real content and confirm no cell or view corruption |
| R6.4 | Group all temporary shapes and preserve stable child order | Contract test inspects group membership and deterministic child order | Required: group and ungroup in Excel; all expected children remain editable |
| R6.5 | Add staging lifecycle with cleanup after each injected failure point | Fault-injection tests leave zero temporary shapes | Required: force a mid-composition failure and inspect cleanup/state restoration |
| R6.6 | Copy the group and verify expected clipboard drawing formats | Windows clipboard-format contract test | Required: copy in Excel and inspect actual Office drawing clipboard formats |
| R6.8 | Wire one-click Copy Editable and concise success/failure feedback | Command, result, and failure-message tests | Required: one click, manual paste to Excel and PowerPoint, then edit/ungroup children |

> **Phase 6 note (revision 10):** the exported **data panel contains the visible parent/top-level rows after collapse**, while the exported **Gantt still contains every render-visible child shape and label** — Critical overlays, child activities and child milestones all render. Data-panel inclusion and Gantt-shape inclusion are separate decisions, and a child graphic is never excluded merely because its source row is hidden. R6.2 composes the collapsed panel; R6.3 composes the full scene. `LiveExcel` is a distinct profile that supplies **no** data panel at all, because the live sheet's real cells are already the panel — the live renderer must never cover worksheet cells with panel shapes.

> **Phase 6 note (revision 4):** **R6.7** now runs **first** in this phase. A spike that can produce an ADR never follows the work it validates; its ID is unchanged so existing references keep resolving.

Exit demonstration: a user clicks once, pastes into PowerPoint manually, ungroups or edits individual elements, and no staging shapes or extra worksheets remain.

## Phase 7 — direct PowerPoint transfer

Goal: one click places an editable shape group on the intended slide without saving or closing user work.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R7.1 | Add PowerPoint application/presentation/slide adapters and ownership rules | Fake-adapter lifecycle and ownership tests | None |
| R7.2 | Resolve attach/start and active-slide cases with typed outcomes | Complete decision-matrix tests | Required: test existing/new PowerPoint and missing/active slide cases |
| R7.3 | Call `PasteSpecial(ppPasteShape)` and validate returned `ShapeRange` | Adapter tests reject null, empty, and zero-bound results | Required: real `ppPasteShape` integration test on each supported Office build |
| R7.4 | Fit/position pasted group within slide margins preserving ratio | Geometry tests for 4:3, 16:9, and custom slides | Required: inspect placement and editability on each slide size |
| R7.5 | Handle protected/no-slide/unsupported-format/COM-busy failures | Typed mapping, bounded retry, and cleanup tests | Required: reproduce safe supported failure cases without hanging or data loss |
| R7.6 | Wire Send to PowerPoint and retain user's presentation ownership | Command and application-ownership tests | Required: confirm no automatic save/close and no termination of existing PowerPoint |
| R7.7 | Repeat 25 transfers and inspect orphaned Office processes/proxies | Harness records counts, timings, and cleanup assertions | Required: run 25 real transfers and inspect processes, shapes, memory, and logs |

Exit demonstration: one click transfers an editable group to the active slide, verifies it, retains aspect ratio, and never saves, closes, or terminates user-owned PowerPoint.

## Phase 8 — 300-DPI PNG export

Goal: export an exact-size PNG from the scene without depending on the Excel clipboard.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R8.1 | Add width/unit parsing, aspect calculation, rounding, and safety limits | Table-driven cm/in/px, ratio, rounding, and extreme-size tests | None |
| R8.2 | Add SkiaSharp renderer for rectangles, lines, polygons, clipping, and z-order | Primitive pixel/golden tests and deterministic rerender comparison | None |
| R8.2a | Decide pinned-font provisioning for raster text and golden images (ADR) | The ADR records installed-versus-bundled-versus-fallback with licence evidence; a font-presence probe reports the exact resolved font on the reference machine; the golden-image policy states what a missing pinned font does — a clear error or approved fallback, never a silent substitution | None |
| R8.3 | Add deterministic text and hatch/pattern rendering | Pinned-font golden, clipping, and pattern tests | None |
| R8.4 | Render the complete representative Gantt at several widths | Exact dimensions and visual-diff tests at approved widths | None |
| R8.5 | Add PNG `pHYs` 300-DPI metadata writer | Chunk order, CRC, unit, and 11811-pixels-per-metre tests | None |
| R8.6 | Add atomic file writer and post-write validator | Corrupt, locked, invalid, and partial-file failure tests | None |
| R8.7 | Add width/unit Ribbon controls and calculated-height preview | Parser, Ribbon-state, validation, and preview-value tests | Required: enter cm/in/px widths and confirm displayed height/pixel preview |
| R8.8 | Wire Export PNG and ensure exact scene crop | Command plus independent dimensions, density, crop, and aspect tests | Required: export from Excel and inspect user flow, selected path, and resulting image |

> **Phase 8 notes (revision 10):** **R8.7** consumes the size-preset contract R4.7H owns rather than defining its own widths, and R8.1's `ExportSize` aspect calculation continues to govern the PNG's height. Raster remains an exporter of the scene as built: because the critical interval is a **filled rectangle** in the scene (ADR-0027), the raster renderer draws a filled rectangle with no critical-specific reinterpretation, exactly as Excel does.

> **Phase 8 notes (revision 4):**
> - **R8.1** is landed early: `ExportSize` — added as the R0.5 sample and later hardened with `MaxTotalPixels` — implements width/unit parsing, aspect calculation, rounding, and safety limits, with `tests/GanttCreator.Raster.Tests/ExportSizeTests.cs`. The row verifies only the residual, the R8.7 preview value that will feed `ExportSize`, and links its acceptance criteria to the existing tests.
> - **R8.2a** precedes R8.3 because deterministic text rendering and every later golden image depend on the font decision.
> - Phase 8 consumes only the Phase 3 scene and the Raster project; it is independent of Phases 6 and 7. Executing it directly after Phase 4 requires one ADR recording the order change under the scope-change protocol — an order change, not a design change.

Exit demonstration: enter a width in each supported unit, export, and independently verify the exact pixel dimensions, aspect ratio, 300-DPI metadata, and absence of surrounding whitespace.

## Phase 9 — resilience, accessibility, and migration support

Goal: make normal failures safe and the transition from the macro tool understandable.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R9.1 | Add cancellation/deadline model for long non-COM rendering operations | Cancellation/deadline tests leave no partial output | Required: cancel through Excel UI and confirm host remains responsive |
| R9.2 | Add structured local diagnostics and log rotation/redaction | Privacy, redaction, rotation, and operation-ID tests | Required: provoke a live error and inspect the user message and local log |
| R9.3 | Add corrupted-settings recovery without losing worksheet data | Missing/corrupt/version-mismatch recovery contract tests | Required: open a corrupted synthetic workbook and confirm data remains intact |
| R9.4 | Add duplicate/missing shape ownership repair | Repair is idempotent and preserves unowned sentinels | Required: damage owned shapes manually, run repair, and inspect the worksheet |
| R9.5 | Add keyboard, high-contrast, and screen-reader labels to Ribbon/errors | Resource and accessibility metadata tests | Required: keyboard, high-contrast, and screen-reader review in Excel |
| R9.6 | Document manual recreation/mapping from legacy workbook columns | Documentation checks and synthetic mapping example validation | Required: follow the mapping once using a copied synthetic legacy workbook |
| R9.7 | Add crash-recovery test matrix for each export stage | Complete fault-injection suite passes with cleanup assertions | Required: repeat host-specific failure points and inspect workbook/Office state |
| R9.8 | Complete offline operation audit | Static dependency/network-call audit and package inspection | Required: disconnected-machine Excel, clipboard, PowerPoint, and PNG acceptance run |

> **Phase 9 notes (revision 4):**
> - **R9.2** is landed early in part: `RollingLog`, `Redactor`, and `IRollingLog` shipped in R0.8 with privacy tests. The row adds structured diagnostics for real command flows and verifies redaction against them; it does not re-create the abstractions.
> - **R9.3** consumes the R2.10 integrity/safe-repair workflow. It verifies end-to-end recovery from synthetic damage and must not re-specify the repair policy.

Exit demonstration: forced failures at worksheet, clipboard, PowerPoint, and file stages leave the workbook usable, Office state restored, temporary artifacts removed, and diagnostics actionable.

## Phase 10 — packaging and release candidate

Goal: produce a signed, installable, reversible release supported by evidence.

| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R10.1 | Define supported Office/Windows matrix and reference test machines | Matrix/schema validation and explicit unsupported-case list | Required: record actual Windows/Office versions, channels, bitness, and machine roles |
| R10.2 | Add deterministic Release packaging and checksums | Compare two clean Release builds and explain every binary variance | Required: load the packaged XLL rather than a developer output |
| R10.3 | Add code-signing hook with secret-free local/CI configuration | Build works without secrets; controlled signing verification checks chain/timestamp | Required: Excel loads the signed candidate under production-like Trust Center policy |
| R10.4 | Add install, update, rollback, and uninstall instructions | Installer/package smoke automation where practical | Required: clean-VM install, update, rollback, and uninstall rehearsal |
| R10.5 | Run complete automated, Office integration, visual, performance, and offline suite | All CI, coverage, mutation, golden, security, and package gates pass | Required: complete supported Office matrix with signed evidence manifest |
| R10.6 | Run exploratory user workflow from blank workbook to PowerPoint/PNG | Record issues and link every fix/defer decision | Required: human end-to-end authoring, editable transfer, and PNG review |
| R10.7 | Freeze dependency versions, third-party notices, and SBOM | Locked restore, vulnerability scan, licence checks, and SBOM validation | Required: confirm packaged native/runtime files match the approved inventory |
| R10.8 | Tag release candidate only after human go/no-go review | Verify commit, version, checksums, evidence links, and clean tree | Required: human accepts all gates and known limitations before tagging |

> **Phase 10 note (revision 4):** **R10.3**'s code-signing certificate is an external procurement with lead time; start it no later than Phase 8 so R10.3 is not blocked on paperwork. **R10.1** finalises the supported Office/Windows matrix from the host-build records that phase-exit Office evidence has carried since Phase 4.

Exit demonstration: a non-developer installs on a clean supported machine, builds a Gantt, transfers editable shapes to PowerPoint, exports a verified 300-DPI PNG, and uninstalls without residual configuration.

## Cross-phase compatibility matrix

Test at phase exits and before release:

| Dimension | Minimum coverage |
| --- | --- |
| Office architecture | x64 primary; x86 only if approved as supported |
| Workbook structure | one visible Gantt sheet, one valid `_GanttCreatorConfig` VeryHidden sheet, no others |
| Excel zoom | 75%, 100%, 125%, 150% visual check |
| Windows display scale | 100%, 150%, 200% visual check |
| Workbook date system | 1900 required; 1904 either supported by tests or rejected clearly |
| Locale | one day-month-year and one month-day-year locale |
| Slide size | 4:3, 16:9, custom |
| Gantt size | empty, 1 event, representative, 1,000 events |
| Date boundaries | leap day, year boundary, one-day activity, same-date events |
| Failure | invalid row, protected sheet, clipboard unavailable, PowerPoint absent/busy, unwritable PNG path |
| Multiple open workbooks | commands act only on the active workbook; Ribbon state follows the active workbook (R1.5); no cross-workbook mutation |

Phase-exit subsets: the phase-exit evidence records which matrix rows ran, and a phase cannot exit with an applicable, unexercised row.

- Phase 2 exit: workbook structure, workbook date system, locale, Gantt size (empty, 1 event, representative), failure (invalid row, protected-sheet read paths).
- Phase 4 exit: adds Excel zoom, Windows display scale, Gantt size 1,000, failure (protected-sheet write paths).
- Phase 6/7 exit: adds slide size and failure (clipboard unavailable, PowerPoint absent/busy).
- Phase 8 exit: adds failure (unwritable PNG path) at the approved widths.
- Phase 9 and R10.5: the complete matrix on every reference machine with the signed evidence manifest.

## Scope-change protocol

If evidence invalidates a fixed design:

1. Stop the current implementation.
2. Preserve the minimal spike and exact results outside production code.
3. Write an ADR with options, consequences, and a recommendation.
4. Obtain human approval.
5. Update architecture, roadmap, tests, and agent rules together.
6. Resume with a new work item.

No agent may convert a discovered limitation into an unreviewed fallback.
