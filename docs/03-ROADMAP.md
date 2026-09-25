
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

> Created 2 September 2026 under a new filename. Revision 6 (25 September 2026) inserts one row, `R2.7c`, which gives the Core style registry's resolved formatting its own commit. The StageInspect audit landed `GanttStyleResolver` as prerequisite work, but that resolver could only read `GanttCatalogues.StylePresets`, so a user-authored style was invisible to resolution — a gap against R2.7b's save/reopen gate and R5.6a's Custom Activity Refresh gate. R3.14 requires resolved style tokens in its equivalence field list, but it is blocked behind unimplemented R3.6–R3.12, so the registry payload lands first as its own tracked row rather than being folded into a test row. Revision 5 (23 September 2026) applied the approved R2 findings implementation plan: eight rows were inserted (`R2.1a`, `R2.2a`, `R2.4a`, `R2.4b`, `R2.5a`, `R2.6a`, `R2.7b`, `R5.6a`), ADR-0009 defines named-style capability columns, and ADR-0010 records the first-live-slice sequencing exception that moves R2.10 and R3.13 after R4.9 without removing or weakening either row. This revision contains 109 commit-sized work items. No existing ID was renumbered, reused, or removed, so work-item files, STATUS, and ADR references keep resolving. Revision 4 (21 September 2026) remains the historical basis for all unchanged rows and gates.

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
| R4.8 | Preserve unowned shapes/cells and active worksheet | Sentinel shape/cell and active-sheet adapter tests | Required: refresh with manual content present and prove it is unchanged |
| R4.9 | Wire Validate and Refresh Ribbon commands with guarded errors | Command, validation, and error-boundary tests | Required: exercise Validate/Refresh on valid and invalid workbooks |
| R4.10 | Profile 1,000-event refresh and remove only measured bottlenecks | Automated benchmark and no-regression threshold | Required: measure real Excel refresh on the reference machine and record Office build |

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
