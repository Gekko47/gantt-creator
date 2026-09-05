---
name: 03-ROADMAP
description: Use this skill when the conversation is about which phase, which work item (R0.x .. R10.8), which automated gate, or which Office gate applies. Also use it for the cross-phase compatibility matrix, the scope-change protocol, and when the user asks for the next safe action.
---

# Implementation roadmap — revision 3
> Created 2 September 2026 under a new filename. This revision contains 96 commit-sized work items and adds creation, named-range catalogues, migration, safe repair, and verification for the approved VeryHidden configuration worksheet. When installed, use the path `docs/03-ROADMAP.md`.
## How to use this roadmap
- Complete items in order unless an approved ADR records why order changed.
- One row is normally one commit. Split a row if the diff becomes difficult to review; do not combine rows merely because they are related.
- Create a work-item file before implementation. Record the exact acceptance tests and evidence.
- Every work-item acceptance criterion that names a test count, a command, or an artifact must link to a specific test or script step. Drift between the doc and the code is a defect (see `docs/08-TEST-CHECKLIST.md` section I).
- Every visual/layout/style/export work item names the affected sections of `docs/07-GANTT-ENTITY-GUIDE.md` and tests the defined cross-renderer contract.
- Every code commit must pass `scripts/verify-quick.ps1`; every phase exit and pull request must pass `scripts/verify.ps1`.
- A phase exits only after its stated demonstration. A screenshot is supporting evidence, not a substitute for automated assertions.
- Use synthetic construction data. Do not add customer schedule data to tests or examples.
The **Automated gate** is run from the terminal and CI. A **Visual Studio / Office gate** marked `Required` must be demonstrated against desktop Excel or PowerPoint on Windows, normally launched or debugged through Visual Studio. `None` means the commit does not require Office; it does not waive the automated gate. A work item cannot be marked done when a required Office gate was not run.
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
| R2.2 | Implement `Initialise Sheet` to create the visible Gantt sheet and `_GanttCreatorConfig` | Adapter tests assert exact names, one visible sheet, one `xlSheetVeryHidden` sheet, and no others | Required: initialise a blank workbook and inspect table, plot anchor, helper visibility, and sheet count |
| R2.3 | Add stable ID generation and preservation independent of row number | Insert, sort, move, and delete contract tests | Required: sort and insert rows in Excel; IDs remain stable and unique |
| R2.4 | Read cell values into neutral row DTOs without locale display parsing | 1900 date-system and multiple-culture conversion tests | Required: read representative real Excel dates under both supported locale formats |
| R2.5 | Map DTOs into Core events with all-errors validation | Table-driven valid, invalid, and all-errors tests | None |
| R2.6 | Add row-level error reporting without mutating valid input | Fake-worksheet mutation and error-order tests | Required: invalid rows show actionable errors while original cells remain unchanged |
| R2.7 | Persist versioned workbook settings and style/metric tables on `_GanttCreatorConfig` | Settings/style/metric round-trip plus invalid/missing schema tests | Required: save/reopen and confirm one visible sheet, one VeryHidden helper, and retained settings |
| R2.8 | Add add-activity, add-milestone, and add-delineator row commands | Command/contract tests for defaults, IDs, and insertion position | Required: invoke all three Ribbon commands and inspect the resulting visible rows |
| R2.9 | Materialise the central Type catalogue and apply its named-range dropdown | Catalogue uniqueness/hash, type/style/date/capability mapping, defined-name, and validation tests | Required: full-name dropdown appears on existing/new rows and resolves only through `_GanttCreatorConfig` |
| R2.10 | Add VeryHidden configuration integrity, migration, and safe-repair workflow | Missing/corrupt/wrong-visibility/version/hash tests preserve valid custom styles or require confirmation | Required: damage/copy/remove configuration in synthetic workbooks and verify detection, repair, and one-visible-sheet rule |
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
Exit demonstration: a command-line/test fixture produces a deterministic scene snapshot containing planned/actual overlaps, three events on one lane, critical segments, milestones, and two labelled delineators.
## Phase 4 — live native Excel renderer
Goal: render and refresh an owned set of Excel shapes beside the source data.
| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R4.1 | Add narrow Excel application/workbook/worksheet/shape adapter interfaces | Architecture test keeps Core reference-clean; adapter contracts compile | None |

---

## Where to read more

- Canonical source: `docs/03-ROADMAP.md`
- Always-on rule: `.clinerules/03-ROADMAP.md`
- Full reference: `./references.md` in this directory
