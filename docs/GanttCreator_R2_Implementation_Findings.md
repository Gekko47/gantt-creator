# GanttCreator — R2 Implementation Findings and Required Changes

## Purpose

This document is an implementation brief for an LLM/coding agent working on the `stage-inspect` branch of the GanttCreator repository.

It consolidates the actionable findings from a code-level review of the R2 implementation through **R2.6**, and distinguishes:

- what is already implemented and should be preserved;
- what should be changed before the implementation hardens further;
- what should be deferred to later roadmap work;
- what constraints the agent must observe when extending the codebase.

The agent should treat the existing production code as authoritative for current behaviour and use the repository's architecture/roadmap documents as supporting intent, not as evidence that a feature exists.

---

# 1. Current R2 Architecture — Preserve This Direction

The current R2 implementation establishes a strong boundary:

```text
Excel workbook
    |
    v
ExcelGanttTableReader
    |
    v
GanttRowDto
    |
    v
GanttRowValidator
    |---------------------->
    |                       GanttValidationIssue[]
    v
GanttEvent[]
    |
    v
future R3 Scene Engine
```

Preserve the separation between:

- **Core**: domain types, parsing, validation, mapping, contracts;
- **Office**: Excel/COM integration and worksheet mutation;
- **AddIn**: commands, Ribbon orchestration and user-facing reporting.

Do **not** move Excel/COM dependencies into `GanttCreator.Core`.

Do **not** make the renderer reinterpret raw worksheet values independently of the R2 mapping/validation pipeline.

---

# 2. R2.1 — Entity Type Catalogue

## Current implementation

The repository has a central `EntityTypeCatalog` backed by a machine-readable `GanttEntityType` enum/value model. Each selectable type defines behaviour including:

- display name;
- entity kind (`Span`, `Point`, `Delineator`, etc.);
- date mode (`Start/Finish`, `StartOnly`, etc.);
- default style key;
- colour capabilities;
- style requirements;
- allowable label positions.

There are currently **16 user-selectable Type values**.

## Preserve

Keep the catalogue as the single source of truth for Type behaviour.

Do not introduce parallel Type lists into:

- Ribbon code;
- worksheet validation;
- scene rendering;
- export renderers;
- UI dropdown population.

Whenever possible, derive those surfaces from `EntityTypeCatalog`.

## Required change: separate stable Type identity from presentation labels

The underlying machine identity is already separated from the display string. Preserve that separation and strengthen it.

### Desired principle

```text
Stable machine identity
        !=
User-facing display label
```

The machine identity should be treated as the durable schema/domain identifier. The visible label may be changed in future versions without forcing the underlying entity identity to change solely because wording/capitalisation changed.

Do not implement full localisation now. The immediate requirement is to avoid making future UI wording changes unnecessarily equivalent to domain/schema changes.

## Required change: clarify 16 selectable Types vs guide sections

Documentation currently refers to a larger number of entity/visual sections than the number of actual dropdown values.

Use the terms:

- **16 selectable Types** for workbook `Type` values;
- **visual/entity guide sections** for the broader documented visual taxonomy.

Do not create additional dropdown values merely to reconcile documentation counts.

---

# 3. R2.1 — Custom Activity Capability Semantics

## Current issue

`Custom Activity` is currently represented as style-defined, but its catalogue entry does not expose normal label-position capabilities.

This effectively makes it a placeholder for a future style lookup rather than a fully specified custom activity type.

## Required change

Resolve the intended semantics before R3 depends on the capability matrix.

Preferred design:

```text
Custom Activity
    -> requires StyleKey
    -> StyleKey resolves a known style definition
    -> style definition determines legal visual capabilities
```

The style definition should be able to define the allowed label-position set.

Do not hard-code Custom Activity as having zero label-position capabilities unless that is explicitly the intended product behaviour.

Add tests covering:

1. Custom Activity with a valid StyleKey and valid label position.
2. Custom Activity with a valid StyleKey and invalid label position.
3. Custom Activity without required StyleKey.
4. Custom Activity with unsupported colour/label override combinations.

---

# 4. R2.2 — Workbook Initialisation

## Current implementation

`ExcelWorkbookInitialiser` performs defensive discovery before mutation.

It:

1. resolves the active workbook;
2. checks for existing `_GanttCreatorConfig`;
3. scans for existing `tblGanttData`;
4. assesses the active worksheet;
5. chooses the target worksheet;
6. checks protection;
7. creates/adopts the Gantt data sheet;
8. creates `tblGanttData`;
9. creates `_GanttCreatorConfig`;
10. sets configuration visibility to `xlSheetVeryHidden`;
11. creates the `GanttCreator.PlotAnchor` defined name;
12. performs rollback on relevant failure paths.

The AddIn layer receives a typed outcome rather than directly controlling Excel objects.

## Preserve

Preserve:

- read-before-write discovery;
- typed outcomes;
- COM isolation;
- rollback/cleanup behaviour;
- VeryHidden configuration sheet;
- defined-name based identity/anchors;
- command/Office boundary.

This is a good foundation for later commands and refresh logic.

## Required change: distinguish Initialisation from Repair/Open

The current initialiser is intentionally conservative and refuses certain existing-workbook conditions, including conflicts around the managed table/configuration.

Do not weaken those safety checks merely to make initialisation succeed.

Instead, plan for three distinct concepts:

```text
Initialise New Gantt
Open/Recognise Existing Gantt
Repair/Recover Gantt Configuration
```

The current `Initialise()` path should remain safe and non-destructive.

Future repair/migration functionality should not be smuggled into the new-workbook initialiser.

## Required change: strengthen adoption criteria

Current blank-sheet detection is based on Excel used-range/content checks.

Before finalising adoption behaviour, add explicit tests for worksheets that appear blank by cell contents but contain potentially significant non-cell state, such as:

- shapes;
- named ranges;
- comments/notes;
- other managed objects.

The adoption rule should mean:

> safe for GanttCreator to take over

not merely:

> CountA returns zero.

Do not over-engineer the complete solution yet; establish the boundary and tests now.

## Required change: preserve stable workbook identity

Visible worksheet names such as `Gantt Data`, `Gantt Data (2)`, etc. are presentation names.

The eventual authoritative identity should continue to come from:

- defined names;
- table identity;
- configuration metadata;
- stable row IDs.

Do not build future logic that depends on the visible worksheet name being exactly `Gantt Data`.

---

# 5. R2.3 — Stable Row Identity

## Current implementation

`GanttRowId` is an immutable value object backed by a UUID-like format, currently generated independently of worksheet row number.

The design intentionally means identity survives:

- sorting;
- insertion;
- deletion of other rows;
- moving rows.

## Preserve

This approach is correct and should remain the foundation for:

- owned shape identity;
- refresh reconciliation;
- selection state;
- parent/child links;
- exports;
- migration.

## Required UX constraint

Treat `GanttRowId` as infrastructure.

Users should not normally be expected to author or edit IDs manually.

The eventual worksheet implementation should consider:

- hidden/protected ID columns;
- automatic ID generation;
- safe repair of missing IDs;
- clear duplicate-ID diagnostics.

Do not redesign the ID as a row-number or description-derived value.

---

# 6. R2.4 — Excel Table Reading

## Current implementation

The reader uses the Excel table data body range and performs a **bulk `Value2` read** rather than repeatedly reading individual COM cells.

Columns are located by header name rather than physical position.

The reader maps Excel values to neutral `GanttRowDto` objects.

Dates are interpreted from underlying Excel numeric/date values rather than displayed text, avoiding locale-dependent parsing of strings such as `dd/MM/yyyy` and `MM/dd/yyyy`.

## Preserve

Preserve:

- bulk COM reads;
- header-name column mapping;
- neutral DTOs;
- no business semantics inside the raw reader;
- no dependence on formatted display text for date parsing.

This is an important performance and correctness property.

## Required change: preserve Excel error state

Current conversion can flatten Excel error values into `null`.

This loses useful diagnostic information. For example, there is a meaningful difference between:

```text
empty cell
```

and:

```text
cell contains #N/A
cell contains #VALUE!
cell contains #SPILL!
```

### Implement

Introduce a neutral cell-state representation, without coupling Core to COM-specific types.

Conceptually:

```text
CellValueState
    Empty
    Value
    ExcelError
    Unsupported
```

Optionally carry a stable neutral error token/code where available.

The important requirement is that validation can distinguish:

- genuinely blank;
- populated with valid value;
- populated with an Excel error;
- unsupported/unreadable value.

Do not expose `Range`, `Xl...` or COM objects in Core.

### Tests

Add reader tests covering representative Excel error payloads and verify that the information survives into the neutral DTO/model sufficiently for diagnostics.

---

# 7. R2.4 — 1904 Date System

## Current implementation

The Office reader currently refuses workbooks using Excel's 1904 date system.

## Required change

Do not redesign the entire date model now, but isolate date-system conversion behind a dedicated abstraction so that 1904 support can be introduced later without rewriting the table reader or domain model.

Desired shape:

```text
Excel serial value
      |
      v
ExcelDateSystemConverter
      |
      v
System.DateTime / DateOnly-equivalent domain value
```

The current product may continue to reject 1904 workbooks if that remains an explicit supported-scope decision.

However, the rejection should be deliberate, test-covered and localised to the date-system boundary.

Add a regression test ensuring 1904 refusal remains deterministic until support is intentionally enabled.

---

# 8. R2.4 — Raw DTO vs Domain Mapping

## Current implementation

`GanttRowDto` remains a neutral/raw representation while `GanttRowValidator` maps valid rows into typed `GanttEvent` objects.

## Preserve

Do not move business semantics back into the table reader.

Do not make DTOs responsible for deciding whether values are valid.

The intended responsibilities are:

```text
Reader
    -> get workbook data

DTO
    -> hold neutral row values

Validator/Mapper
    -> validate and normalise

GanttEvent
    -> represent valid domain semantics
```

This boundary should become the hand-off into R3.

---

# 9. R2.5 — Row Validation and Mapping

## Current implementation

`GanttRowValidator` performs validation and, for valid data, maps rows into `GanttEvent` objects.

It validates items including:

- row ID;
- duplicate IDs;
- Type;
- LaneId;
- StackIndex;
- date semantics;
- ParentId;
- StyleKey;
- LabelPosition;
- FillColour;
- StrokeColour;
- SortOrder;
- critical-interval parent references.

Validation is not first-error-only. Multiple issues can be returned per row and then globally ordered deterministically.

Rows with blocking errors are excluded from the resulting event collection; warning-only rows remain renderable.

## Preserve

Keep:

- all-errors validation;
- deterministic issue ordering;
- blocking vs warning distinction;
- separate issue collection and valid-event collection;
- second-pass cross-row checks.

This is highly suitable for Excel authoring.

## Required change: explicitly recognise validation as normalisation

The current validator does more than identify bad data. It also normalises/decides how irrelevant properties are treated.

Examples include:

- ignoring Finish for StartOnly entities;
- warning when LaneId is not used by a type;
- normalising supported colour/label values;
- resolving parent relationships.

Treat this as an intentional design choice.

Document and test the rule:

> R3 must consume the validated `GanttEvent` model and must not independently reinterpret the original DTO fields.

This prevents two different semantic interpretations from emerging between R2 and R3.

## Required change: distinguish irrelevant vs erroneous values

Where a field is not applicable to a Type, decide explicitly whether the product wants:

- warning + ignore;
- blocking error;
- automatic clearing/normalisation.

Do not let this vary casually from entity to entity.

The default should remain:

> warning + deterministic ignore

where the value is harmless and unambiguous.

Use blocking validation where the value could alter geometry, identity, ownership or interpretation materially.

## Required change: Critical Interval semantics

The current cross-row check correctly requires a valid `ParentId` and verifies that the parent is a span.

Preserve that structure.

Do not push geometry-dependent checks into basic row validation unless they are purely semantic.

Future R3 logic should determine whether the critical interval:

- falls within the parent span;
- requires clipping;
- extends beyond the parent;
- needs a warning or blocking error.

Keep data validity and geometry validity conceptually distinct.

---

# 10. R2.6 — Validation Reporting

## Current implementation

Validation is exposed through `GanttValidationIssue`/report objects and then written into classic Excel Notes attached to offending cells.

Owned notes are identified using a GanttCreator-specific prefix/text marker because the classic Note author's identity is not writable/reliably controllable for ownership purposes.

Multiple issues for the same cell are consolidated.

Existing user-created notes are preserved.

Stale GanttCreator-owned notes are removed during a subsequent validation run.

The underlying cell values are not changed.

## Preserve

Preserve:

- non-destructive reporting;
- cell-specific diagnostics;
- aggregation of multiple issues per cell;
- protection of user notes;
- explicit ownership of generated annotations;
- a pure validation result independent of the reporting surface.

## Required architectural change: Notes must not become the validation architecture

Classic Excel Notes are acceptable as the **current output mechanism**, but they should not be treated as the long-term validation UI.

The authoritative object is the validation result:

```text
GanttValidationIssue[]
```

The note writer is only a reporter.

Future reporting surfaces may include:

- validation pane;
- warning/error summary;
- row navigation;
- field filtering;
- click-to-select offending cell.

Do not redesign the validation model to depend on Notes.

## Required change: keep reporter replaceable

Maintain a boundary like:

```text
Validation result
       |
       +----> Excel Notes reporter
       |
       +----> future validation pane
       |
       +----> future Ribbon summary
```

No Core validation type should depend on Excel Notes or COM.

## Current partial-failure behaviour

The note writer can clear stale owned notes before writing the new set. If a later write fails, some new notes may be missing until the next validation run.

Do not add complicated transactional note staging purely to eliminate this minor presentation-layer inconsistency.

The current failure requirement should remain:

- schedule data remains unchanged;
- user notes remain protected;
- failure is surfaced deterministically;
- the next successful validation run repairs the displayed annotations.

---

# 11. Hidden Configuration Worksheet

## Current design

`_GanttCreatorConfig` is intended as a `xlSheetVeryHidden` configuration store.

It must not become a second schedule database.

## Preserve strict separation

Allowed examples:

- style definitions;
- geometry defaults;
- catalogue metadata;
- named-range support;
- configuration version/hash metadata.

Do NOT store there as a shadow copy:

- activity rows;
- source dates;
- scene geometry;
- export staging data;
- duplicate schedule entities.

The visible `tblGanttData` must remain the authoritative user data source.

---

# 12. Data / Style Responsibility Hierarchy

The emerging design should use the following hierarchy:

```text
Entity Type
    |
    v
StyleKey
    |
    v
Global Style Definition
    |
    +---- row-level supported override
```

Avoid creating a situation where every row independently defines a complete style object.

A normal row should inherit its visual properties from its Type/StyleKey.

A selected expanded entity may override specifically permitted values such as:

- FillColour;
- StrokeColour;
- LabelPosition.

The renderer should resolve the final style deterministically from this hierarchy.

---

# 13. User-facing Entity Controls

The original product requirement remains important and should be preserved in later R5 work.

When the user has selected a single expanded entity row, the UI must eventually support changing:

- shape fill colour;
- stroke/outline colour where applicable;
- label position where applicable;
- Type from the approved dropdown;
- other supported entity-specific controls.

Changes should update the worksheet/configuration state.

The generated Gantt should update on **Refresh**, not on every edit.

Do not implement direct shape-only formatting as the source of truth.

The source of truth remains workbook data/configuration.

---

# 14. Refresh-only Rendering Rule

This is a product invariant and must remain explicit.

Normal workbook edits must NOT automatically redraw the Gantt.

The intended flow is:

```text
User edits data/control
        |
        v
Workbook state changes only
        |
        v
User clicks Refresh
        |
        v
Read -> validate -> map -> scene -> render
```

Required regression tests should establish that these actions do not render:

- changing Type;
- changing Start/Finish;
- changing FillColour;
- changing StrokeColour;
- changing LabelPosition;
- selecting another row.

Refresh is the render trigger.

---

# 15. Upcoming R3 — Scene Engine Constraints

Before R3 is implemented, carry forward these rules.

## Rule 1 — Consume `GanttEvent`

The scene engine must consume the validated domain model rather than raw Excel rows.

## Rule 2 — Deterministic geometry

Given the same validated input and configuration, the scene must be identical.

## Rule 3 — No COM in Core

Scene calculation must be testable without Excel installed.

## Rule 4 — Geometry semantics belong in R3

Examples:

- date-to-X transformation;
- clipping;
- minimum visible widths;
- lane heights;
- stack placement;
- milestone geometry;
- delineator height;
- critical interval geometry;
- label placement.

Do not overload R2 validation with geometry decisions unnecessarily.

## Rule 5 — No renderer-specific interpretation

Excel, PowerPoint and PNG should consume the same scene semantics.

They should not each independently decide what an As-Built bar or Delineator means.

---

# 16. Construction Delay Analysis Reference Fixture

Before R3/R4 become large, establish one canonical reference workbook/fixture containing at least:

- As-Planned Activity;
- As-Built Activity;
- Baseline Activity;
- Critical Interval;
- Delay Event;
- at least one procurement entity;
- planned milestone;
- actual milestone;
- baseline milestone;
- critical milestone;
- two delineators;
- multiple events on one lane;
- overlapping planned/actual/baseline events;
- one example using row-level FillColour override;
- one example using row-level LabelPosition override;
- at least one parent/child relationship.

This fixture should become the principal acceptance case for R3 and R4.

The target is not a generic project-management Gantt. The target is a construction-delay analysis visualisation that exercises the actual GanttCreator data model.

---

# 17. Recommended Development Sequencing Change

Do not allow the project to remain dominated by governance/documentation while production rendering remains absent.

The architecture is sufficiently mature to move into a vertical product slice.

Recommended sequence:

## Phase A — Complete the essential R2 foundation

Finish the remaining R2 items required for:

- reliable IDs;
- reliable row reading;
- validation;
- essential configuration;
- Type dropdown/source.

Defer non-essential migration/recovery complexity unless it blocks the first working workflow.

## Phase B — Build the first real Gantt

Implement a minimal R3/R4 vertical slice supporting:

- span activity;
- milestone;
- delineator;
- label;
- lane layout;
- time scale;
- Excel native rendering;
- Refresh.

This should produce an actual usable Gantt before the entire advanced roadmap is completed.

## Phase C — Add construction-delay complexity

Then exercise:

- planned vs actual;
- baseline;
- critical intervals;
- delay events;
- procurement;
- stacked/overlapping events;
- parent/child;
- selected-row overrides.

## Phase D — Professional authoring

Complete the more advanced controls, configuration, UX and warning surfaces.

## Phase E — Editable exports

Add PowerPoint/native Office composition from the same scene.

## Phase F — PNG/export hardening

Add the high-DPI raster pipeline and its strict verification requirements.

## Phase G — Recovery, migration and release hardening

Finish:

- migration;
- configuration repair;
- resilience;
- accessibility;
- performance hardening;
- packaging;
- signing;
- release validation.

---

# 18. Features That Should NOT Be Pulled Forward Prematurely

Do not delay the first working Gantt for features that are valuable but downstream, including:

- extensive migration/recovery machinery;
- full mutation-testing programme;
- complete PowerPoint export;
- 300-DPI PNG implementation;
- final release packaging/signing;
- exhaustive configuration repair tooling.

They remain part of the product roadmap.

They should not dominate the engineering sequence before the core Excel visualisation is demonstrably working.

---

# 19. Testing Requirements for These Changes

Every change described in this document must include targeted tests appropriate to its layer.

## Core tests

Cover:

- Type parsing/formatting;
- capability resolution;
- Custom Activity semantics;
- stable IDs;
- Excel cell-state conversion;
- date-system abstraction;
- validation issue ordering;
- validation blocking/warning behaviour;
- cross-row parent validation.

## Office tests

Cover:

- bulk table reading;
- named-column resolution;
- blank/adopt/create decisions;
- protection refusal;
- rollback;
- 1904 behaviour;
- Excel error values;
- note ownership/preservation/cleanup.

## AddIn tests

Cover:

- command orchestration;
- typed outcomes;
- non-destructive behaviour;
- correct presenter invocation.

## Regression requirement

Existing R2 tests must remain green after changes.

Do not weaken tests merely to make new implementation pass.

---

# 20. Acceptance Criteria for the R2 Changes

The following should be true when the R2 hardening described here is complete:

### Type system

- One authoritative Type catalogue exists.
- Machine Type identity is independent from user-facing wording.
- Custom Activity has fully defined style/label capability semantics.
- Exactly 16 selectable Types remain unless intentionally changed by a product decision.

### Initialisation

- New workbook/sheet initialisation remains conservative and non-destructive.
- Adoption is tested against non-cell artefacts.
- Visible worksheet name is not used as the authoritative workbook identity.
- Existing conflicting managed state is refused rather than overwritten.

### Identity

- Row IDs are stable and independent of worksheet row number.
- Users are not required to manually author IDs in normal workflows.

### Reading

- Table data is read in bulk.
- Column mapping is header-driven.
- Dates are read from underlying Excel values, not display strings.
- Excel error values are distinguishable from empty cells.
- Date-system conversion is isolated and testable.

### Validation

- Multiple issues can be reported per row.
- Issue ordering is deterministic.
- Blocking issues prevent invalid events from entering the domain model.
- Warning-only issues do not unnecessarily prevent rendering.
- Cross-row references are validated in a separate deterministic pass.
- R3 consumes validated `GanttEvent` objects rather than reinterpreting raw rows.

### Reporting

- Validation results are independent from the reporting surface.
- Classic Notes remain non-destructive.
- User-authored notes are preserved.
- GanttCreator-owned notes can be removed/recreated safely.
- A future validation pane can be added without changing the validation engine.

---

# 21. Explicit Non-Goals for the Agent

Do not:

- rewrite the R2 architecture wholesale;
- collapse Core, Office and AddIn projects into one layer;
- make Excel COM types available in Core;
- make shapes the source of truth;
- implement live redraw on every cell edit;
- create a hidden duplicate schedule model in `_GanttCreatorConfig`;
- identify entities by worksheet row number;
- add ad-hoc Type lists outside `EntityTypeCatalog`;
- make Notes the underlying validation model;
- pull PowerPoint/PNG implementation forward just because the architecture already anticipates them.

---

# 22. Implementation Priority

Use this priority order:

### P0 — Correctness and architecture preservation

1. Preserve the R2 boundaries.
2. Resolve Custom Activity semantics.
3. Preserve stable IDs.
4. Preserve bulk/header-driven reading.
5. Preserve all-errors validation.
6. Preserve Refresh-only rendering.

### P1 — R2 hardening

7. Preserve Excel error state.
8. Isolate date-system conversion.
9. Strengthen safe worksheet adoption tests.
10. Keep validation reporting independent from Notes.
11. Clarify Type identity vs display label.

### P2 — First working product slice

12. Move into R3 scene calculation.
13. Render a minimal Gantt in Excel.
14. Exercise the construction-delay reference fixture.

### P3 — Product expansion

15. Advanced selected-entity controls.
16. Advanced delay-analysis entity behaviour.
17. Editable Office output.
18. PNG/export hardening.
19. Migration/recovery/release hardening.

---

# 23. Final Instruction to the Coding Agent

Implement the changes in this document **incrementally** and in the repository's existing architectural style.

Before modifying code:

1. Inspect the current implementation on `stage-inspect`.
2. Locate the exact classes/tests affected.
3. Reuse existing abstractions where possible.
4. Do not create duplicate concepts.
5. Add/modify tests with each behavioural change.
6. Preserve existing R2 behaviour unless this document explicitly changes it.
7. Do not claim an item is complete unless the relevant code and tests exist.

After each meaningful change:

1. build the affected projects;
2. run the targeted test set;
3. run the broader relevant test set;
4. inspect the diff for architectural regressions;
5. record any deliberate deviation from this brief.

The immediate engineering objective is **not** to rewrite R2.

It is to **harden the small number of weaknesses identified above and then use the R2 pipeline as the stable input boundary for the first real Gantt scene/rendering implementation**.
