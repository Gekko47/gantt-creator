# GanttCreator Revised Implementation Plan - Revision 6

> **Read this first (added 2026-09-29).** This document is a **tracked
> governance input**, not the binding contract. Several of its proposals were
> **changed or rejected by the product owner**. Implement
> [`PRE-R4.9-DECISION-REGISTER.md`](PRE-R4.9-DECISION-REGISTER.md) instead, and
> treat the following as settled:
>
> | This document proposes | Resolved as |
> | --- | --- |
> | §9, §16 — "Core must emit a truthful line-like scene primitive" | **Rejected.** The critical overlay stays a **filled rectangle** (`CriticalFill` = `#FF0000`), half the parent bar's resolved height, top-aligned, `CriticalLinePt` retired. The stated reason was ease of sizing. |
> | §16 — separate `HostShapeKey` from `PrimitiveId` | **Rejected.** R4.7's L18 refusal stands. |
> | §15 — remove `Above`/`Below` | **Accepted.** Retain `Auto`, `Left`, `Right`, `Inside` (UI: Centre), `None`, the four delineator corners, and the splitter positions. |
> | §15 — `Auto` is largest-space-that-fits | **Accepted, with the existing truncation retained.** ADR-0015's D4/D5/D6 stand; this document does not replace them. Nothing new needs to be built. |
> | §1, §11 — `Duration` and `SiblingOrder` columns | Accepted. Schema 3 → 4, edited in place; **no migration** (no active users). |
> | §3 — `MaxChildrenPerParent = 7` | Accepted. |
> | §17 — size presets | Accepted; roadmap row R4.7H. |
> | §22, §23 — Refresh orchestrator and request factory | Accepted; roadmap row R4.8A. |
>
> Everything in this document not listed above stands as written.

**Authority:** supersedes REV5 and earlier R3/R4 remediation plans.

This plan incorporates the settled pre-R4.9 product decisions: real
parent/child rows; Excel outline expand/collapse; child shapes projected
onto parent lanes; Critical Interval as a real child with independent
dates; maximum seven children; fixed managed row height; exact row/lane
alignment; visible locked Duration; hidden locked engine metadata;
Ribbon Row Styling; promote-on-delete; new identities for copied/cloned
data; date-only semantics; row-based selection; collapsed export data
panel with all child graphics; output-size presets; horizontal-only
label placement; explicit mutation ownership; and a production Refresh
orchestrator before R4.9.

Preserve the existing Core/Office/AddIn separation, scene-first
architecture, stable identity, typed validation, shape
ownership/reconciliation, origin translation and application-state
restoration.

# 1. Workbook column contract

Classify every managed table column.

**Visible + editable:** Type, Description/Activity Name, Start, Finish,
and only other genuine user inputs.

**Visible + locked:** Duration.

**Hidden + locked + engine-maintained:** Id, ParentId, LaneId,
StackIndex, SiblingOrder (or equivalent), StyleKey, fill/stroke/label
overrides, projection metadata and other internal relationship/version
fields.

Engine columns may remain physically in `tblGanttData`; they are not
user interface. Initialise them hidden/locked and repair that state if
altered.

# 2. Identity lifecycle

IDs are engine-only. Normal user copy/paste transfers visible
plain-text/user values only; a pasted entity receives a new ID.

Assign identity when the single blank entry row first becomes a
meaningful entity. A completely blank entry row is not an entity.

Duplicating a managed worksheet/table must trigger an atomic identity
reseed:

``` text
detect duplicate -> build old->new Id map -> rewrite all Ids
-> rewrite ParentIds/internal references -> derive lane/projection metadata
-> rebuild outline groups -> establish new ownership context
```

Do not repair a cloned hierarchy one duplicated row at a time.

Tests must prove pasted rows get new IDs, cloned parents/children retain
relationships using new IDs, and valid IDs remain stable across Refresh.

# 3. Hierarchy and ordering

`ParentId` is authoritative; Excel grouping is presentation.

Add engine-maintained `SiblingOrder` or equivalent so hierarchy order is
not defined solely by physical Excel row number.

Arbitrary native sorting must not silently redefine hierarchy. Prevent
it through the managed protection/UI model or detect inconsistency on
Refresh and refuse.

Set `MaxChildrenPerParent = 7`. Add Child refuses an eighth child.

# 4. Add Child workflow

Provide Ribbon **Add Child**.

1.  Resolve selected row to its hidden stable entity ID.
2.  Verify the entity may be a parent.
3.  Count children and enforce the seven-child maximum.
4.  Offer only permitted child Types from a Core capability matrix.
5.  Insert after the existing descendant block.
6.  Generate child ID.
7.  Write hidden ParentId.
8.  Assign SiblingOrder.
9.  Derive LaneId/projection metadata.
10. Set Type and managed row height.
11. rebuild/extend the Excel outline group.
12. Select the child row for user entry.
13. Do not render until Refresh.

Critical Interval uses this normal child workflow.

# 5. Delete workflow - promote children

Managed deletion occurs only through a dedicated Ribbon command. Native
managed-row deletion should be prevented by worksheet protection where
practical.

Deleting a parent **promotes**, rather than deletes, its children.

If a top-level parent is deleted, children become top-level. If the
deleted parent itself had a parent, its children inherit that parent.

Promoted children retain their IDs and schedule data. Recompute
ParentId, SiblingOrder, projection/lane ownership and Excel outline
grouping.

Use confirmation because custom Undo will not be implemented.

# 6. Source-row visibility versus render visibility

Implement separate concepts such as `SourceRowVisible` and
`RenderVisible`.

A collapsed child is:

``` text
SourceRowVisible = false
RenderVisible    = true
```

Do not map Excel `EntireRow.Hidden` directly to `GanttEvent.Visible`.

Collapse affects source-data presentation, not semantic Gantt inclusion.

# 7. Projection model

Introduce a Core projection contract equivalent to:

``` csharp
EntityProjection
{
    SourceEntityId
    RenderLaneOwnerId
    ProjectionMode
}
```

Modes should cover `OwnLane`, `OverlayParent`, `StackOnParent`,
`MilestoneOnParent`, `PlotGlobal` and structural lanes as required.

Typical mapping: top-level activity/milestone -\> OwnLane; Critical
child -\> OverlayParent; child activity -\> StackOnParent; child
milestone -\> MilestoneOnParent; Delineator -\> PlotGlobal;
Splitter/Spacer -\> structural own lane.

Projection is resolved in Core before scene construction. Excel must not
infer it.

# 8. LaneId, StackIndex and child overlap

LaneId and StackIndex are engine-maintained and hidden.

Derive render-lane ownership from hierarchy/projection where possible.
Prevent contradictory ParentId/LaneId states.

Children do **not** vertically expand or stack the parent row. They
generate on the same parent lane and may overlap. Distinguish them by
horizontal date geometry, Type, style, z-order and labels.

StackIndex may remain as an internal deterministic ordering concept but
must not drive lane-height growth.

# 9. Critical Interval

Critical Interval is a genuine user-authored child row with its own
Description, Start, Finish, Duration, ID and permitted styling. It does
not inherit dates from its parent.

Its visual is a line/line-span overlay centred on the parent activity,
approximately half the normal activity-bar thickness, above the parent
in z-order, and consuming no extra lane height.

Add a Core metric such as `CriticalOverlayHeightRatio = 0.5`.

Core must emit a truthful line-like scene primitive. Do not emit a thin
rectangle that each renderer reinterprets.

# 10. Uniform row height and exact alignment

Introduce `GanttRowHeightPt` as the normal managed-row height.

On Initialise, hierarchy mutation and Refresh, normalise managed row
heights. A manually changed managed row is restored on Refresh unless
protection/ownership makes correction unsafe.

Every visible lane-owning row must satisfy:

``` text
Excel Top == Scene lane Top
Excel Height == Scene lane Height
Excel Bottom == Scene lane Bottom
```

Expanded projected-child source rows consume worksheet height but create
no plot lane; blank plot space opposite them is intentional. Collapsed
child rows consume no visible height, so later parent lanes move upward.

Remove live lane auto-growth. All shapes fit within the fixed row
height. Splitter/Spacer must also match their corresponding Excel row
geometry.

# 11. Duration and date semantics

Duration is visible but locked and recalculated by GanttCreator on
Refresh.

Dates are **date-only**. Time-of-day has no semantic meaning; normalise
Excel date-time values consistently to their date component.

For inclusive span entities:

`DurationDays = (Finish.Date - Start.Date).Days + 1`

Milestones and delineators display the approved non-duration marker
(e.g. `-`). Splitter/Spacer use blank or the approved placeholder.
Invalid/missing dates must not produce misleading duration.

This is calendar-day duration unless/until a working-calendar feature is
explicitly added.

Prefer one bulk Duration write and GanttCreator-owned values rather than
editable formulas.

# 12. Single blank entry row

Maintain one blank entry row. A completely blank row is not an entity
and need not have an ID.

On first meaningful entry, create the entity, assign metadata, validate
it and ensure one blank entry row remains.

Define one Core `IsSemanticallyEmptyRow()` rule and reuse it rather than
scattering blank-row tests.

# 13. Row-based selection

The user selection model is the selected worksheet row. Commands resolve
that row to the hidden stable ID before mutation.

Row Styling styles only the selected entity. A collapsed parent may
provide group context for navigation/expand-collapse, but it is not a
styleable group and Row Styling must not implicitly style children.

# 14. Row Styling

Provide **Row Styling** and **Reset Row Styling** Ribbon commands.

Expose only capabilities permitted for the selected entity. Persist only
overrides, retaining:

`Type -> default StyleKey -> configured style -> row override`

Users never edit hidden style columns. Styling changes do not render
until Refresh.

# 15. Label placement

Remove `Above` and `Below` from supported/user-selectable label
positions. Retain horizontal choices such as Auto, Inside, Left, Right
and None.

`Auto` is Core-owned and deterministic. Measure label width and evaluate
horizontal candidate spaces (inside, left, right), considering plot
boundaries, entity geometry, other labels on the lane, milestones and
projected children. Select the largest valid non-overlapping region.

Define one deterministic final fallback if no region fully fits. Never
solve label collisions by increasing row height.

# 16. Entity-to-primitives and host identity

Formalise `Entity -> 0..N ScenePrimitives`. One entity is not one Excel
shape.

Use deterministic primitive roles such as `:bar`, `:label`,
`:critical-line`, `:milestone`.

Separate full logical `PrimitiveId` from bounded Excel `Shape.Name`.
Introduce deterministic collision-resistant `HostShapeKey` within Excel
limits and retain ownership mapping to the logical identity.

# 17. Output size and plot geometry

Introduce a first-class size preset contract with initial presets:

-   A4
-   Presentation 16:9
-   Presentation 4:3

Define exact dimensions/orientation centrally.

Text/data columns retain their measured widths. The selected preset
defines total composition width; after approved margins/borders and the
fixed text-panel width, all remaining width becomes the Gantt/time-grid
area.

`PlotWidth = PresetWidth - TextPanelWidth - Margins/Borders`

Title and time-header bands use the same resulting composition geometry.

Do not resize user text columns to make a preset fit. If insufficient
width remains for a valid plot, Refresh must return a clear layout
error.

The SceneBuildRequestFactory/orchestrator is the production authority
for deriving PlotBounds from preset, text-panel measurement, date range
and scale.

# 18. Mutation ownership boundary

Make ownership explicit.

**GanttCreator owns:** managed table schema; engine columns; Duration;
managed row heights; its hierarchy outline groups; GanttCreator defined
names; `_GanttCreatorConfig`; GanttCreator validation artefacts; owned
Gantt shapes; managed protection configuration.

**GanttCreator does not own:** unrelated worksheets, shapes, names,
cells, user formatting outside managed ranges or unrelated outline
groups.

If Refresh cannot perform a required operation without changing unowned
content, refuse and report the conflict.

# 19. Worksheet protection

Design protection deliberately.

Users may edit unlocked authoring cells, select rows/cells, use outline
+/- and invoke Ribbon commands.

Users may not edit Duration/engine columns, arbitrarily insert/delete
managed rows, arbitrarily sort hierarchy, alter IDs/relationships or
change managed row heights.

The add-in must retain controlled ability to insert/delete rows, write
locked engine state, write Duration, group/ungroup, normalise heights
and repair metadata.

Add a live Excel test for protected-sheet outlining, including
`EnableOutlining` and any protection flags that must be reapplied per
session.

# 20. Export contract

Before export, collapse hierarchy for the **data-panel presentation**.

The exported data panel contains visible parent/top-level rows after
collapse.

The exported Gantt still contains all render-visible child shapes and
labels, including Critical overlays, child activities and child
milestones.

Data-panel inclusion and Gantt-shape inclusion are separate decisions.
Never exclude a child graphic merely because its source row is hidden.

Define explicit composition profiles:

-   LiveExcel: existing Excel cells are the panel; no panel shapes.
-   EditableExport: collapsed visible panel + all Gantt graphics.
-   PowerPoint: approved export semantics.
-   Raster: approved export semantics.

# 21. Undo policy

Do not implement custom Undo. Ribbon mutations may clear Excel's native
Undo history. Mitigate with dedicated commands, confirmations for
destructive operations, promote-on-delete, deterministic metadata and
repairable Refresh.

# 22. Refresh orchestrator - mandatory before R4.9

Introduce `IGanttRefreshOrchestrator` or equivalent and keep
`RefreshSheetCommand` thin.

Pipeline:

``` text
resolve workbook/sheet/table
 -> validate mutation boundary
 -> enforce hidden/locked column state
 -> read user data + engine metadata
 -> process blank/new entities and IDs
 -> validate hierarchy/order
 -> maintain outline groups
 -> calculate/write Duration
 -> normalise row heights
 -> resolve source/render visibility
 -> resolve projection/lane ownership
 -> build style registry
 -> resolve size preset
 -> resolve plot date range/scale
 -> measure text panel + visible row geometry
 -> build SceneBuildRequest
 -> build LiveExcel scene
 -> translate
 -> ownership/preflight
 -> reconcile native shapes
 -> return outcome
```

All validation/layout/scene construction must succeed before shape
mutation.

# 23. SceneBuildRequestFactory

Add a dedicated testable component to assemble Core scene inputs from
workbook/configuration state.

It owns translation of visible row geometry, projections, plot bounds,
size preset, date range, scale, styles, metrics and LiveExcel
composition into `SceneBuildRequest`.

Do not scatter this logic through Ribbon commands or renderer classes.

# 24. Pre-R4.9 implementation work packages

### R4.7A - Identity and hierarchy

ID lifecycle; SiblingOrder; ParentId; copy/paste new identities;
duplicate reseed; seven-child limit; hierarchy validation;
promote-on-delete.

### R4.7B - Projection

EntityProjection; capability matrix; lane-owner resolution; projected
children; source/render visibility separation; derived LaneId/StackIndex
policy.

### R4.7C - Workbook presentation/protection

Column classification; hidden/locked engine state; visible locked
Duration; deliberate protection; protected outlining; one blank entry
row; repair of presentation state.

### R4.7D - Row geometry

GanttRowHeightPt; normalisation; measurement after outline state; exact
alignment; remove lane auto-growth; parent geometry for projected
children; structural-row alignment.

### R4.7E - Critical correction

Independent child dates; OverlayParent projection; line-like primitive;
\~50% thickness; parent-centred geometry; Excel line translation.

### R4.7F - Duration/date

Date-only normalisation; Duration calculator; inclusive spans;
point/structural rules; bulk write; invalid-date handling.

### R4.7G - Labels

Remove Above/Below; update capabilities; largest-valid-horizontal-space
Auto; deterministic fallback; collision tests.

### R4.7H - Size/layout

Preset model; A4/16:9/4:3 definitions; fixed text-panel measurement;
remaining-width plot; insufficient-space refusal; title/time-band
geometry.

# 25. R4.8 - preservation and safety

Extend R4.8 to prove preservation of user shapes, unrelated
cells/formatting/worksheets/names/groups, unowned same-name shapes,
hidden/locked engine state, Duration lock, outline state, collapsed
child source rows and projected-child rendering.

Force a mid-reconciliation failure and document the partial-refresh
state.

# 26. R4.8A - orchestration integration

Before R4.9 land the Refresh orchestrator, SceneBuildRequestFactory,
plot/date/scale resolver, size/layout resolver, workbook geometry
measurement, LiveExcel composition profile and mutation-boundary
preflight.

Add orchestration tests without Excel where possible plus focused live
Excel integration tests.

# 27. R4.9 - first complete product Refresh

R4.9 should integrate settled contracts, not invent architecture.

Acceptance:

``` text
Initialise
 -> enter parents
 -> Add Child entities
 -> enter independent child dates
 -> expand/collapse
 -> optional Row Styling
 -> Refresh
 -> Duration updates
 -> heights normalise
 -> hierarchy/projection resolves
 -> collapsed children still render
 -> selected output size fits
 -> native shapes reconcile
 -> lane-owning rows align with Excel
```

Use a representative construction-delay workbook.

# 28. R4.10 - performance

Instrument table/metadata read, identity maintenance, hierarchy/order,
outline maintenance, Duration write, row normalisation, geometry
measurement, validation, projection, scene build, translation, shape
enumeration, update/create/delete/z-order and total Refresh.

Test collapsed and expanded hierarchies, a parent with seven children
and a large realistic schedule.

If repeated shape searches dominate, introduce a scoped indexed render
session that enumerates shapes once and operates through an ownership
dictionary.

# 29. R5 roadmap changes

R5 should implement user workflows over the established hidden engine
model:

1.  Add Child UI and permitted Type selector.
2.  Delete/promote UI.
3.  Hierarchy-safe Move Up/Down.
4.  Expand/Collapse/Expand All/Collapse All.
5.  Row Styling.
6.  Reset Row Styling.
7.  Size selector (A4 / 16:9 / 4:3).
8.  Plot range and scale.
9.  Delineator authoring.
10. Validation navigation.

Never expose ParentId, LaneId, StackIndex, StyleKey or IDs as a
substitute for these workflows.

# 30. Required integration fixtures and live Excel tests

Create a non-Excel vertical fixture from worksheet-like DTOs/metadata
through identity, hierarchy, validation, projection, Duration,
SceneBuildRequest, SceneBuilder, SceneShapeRenderer and
OfficeShapeRequest.

Include planned/actual/baseline/delay/procurement/milestone, parent with
seven children, multiple Critical Intervals, child activity, child
milestone, Splitter, Spacer, Delineator, collapsed-but-render-visible
child, render-disabled entity, Row Styling and clipping.

Before R4.9 closure, live Excel tests must prove: protected input
editing; engine/Duration locks; outline +/- while protected and after
reopen; Add Child grouping; collapsed children still render; parent
delete promotes; row height restores; copied visible data gets new IDs;
duplicated managed structures reseed; size presets preserve text widths;
insufficient plot width errors; Row Styling persists hidden overrides;
Refresh is idempotent; unowned shapes are preserved.

# 31. Explicit Refresh errors

Return actionable errors rather than guessing for mutation-boundary
conflicts, protection failures, broken/ambiguous hierarchy, \>7
children, unsafe identity duplication, insufficient plot width, invalid
plot date range, invalid required dates, unsupported hierarchy depth,
ownership collision, unsafe duplicate managed structure and
configuration corruption.

# 32. Pre-R4.9 gate

R4.9 must not begin until implemented/tested:

**Identity/hierarchy:** ID lifecycle, copy/paste identity, clone reseed,
ParentId, SiblingOrder, seven-child maximum, promote-on-delete.

**Workbook:** column classification, hidden/locked engine fields, locked
Duration, blank entry row, protection, protected outlining,
GanttRowHeightPt.

**Domain/scene:** projection, source/render visibility, child overlap,
Critical line, entity-\>0..N primitives, date-only semantics, horizontal
label contract.

**Layout:** size presets, fixed text widths, remaining-width plot,
shared title/time geometry, insufficient-space refusal.

**Integration:** mutation boundary, Refresh orchestrator,
SceneBuildRequestFactory, LiveExcel composition, vertical integration
fixture.

# 33. Definition of done for R4

R4 is complete only when the user sees only appropriate input/calculated
fields; engine columns stay hidden/locked; Duration is correct/locked;
one blank row behaves correctly; hierarchy works without exposing IDs;
parent deletion promotes children; seven-child limit is enforced;
expanded children are editable; collapsed children remain rendered;
lane-owning rows align exactly with Excel; projected children overlap on
parent lanes without row growth; Critical uses independent dates and a
half-thickness line; labels use horizontal placement with deterministic
Auto; size presets preserve text widths and allocate remaining width to
the plot; unsafe conflicts raise errors; Refresh is explicit/idempotent;
validation/build failure preserves the last valid chart; unowned content
is preserved; and realistic performance is measured.

# 34. Instructions to the coding LLM

Treat this plan as the current pre-R4.9 product contract. Do not expose
engine metadata. Do not use row number as persistent identity. Keep row
selection as the user interaction. Keep ParentId authoritative and
grouping presentational. Promote children on parent deletion. Atomically
reseed cloned managed structures. Keep projected children on the parent
lane and allow overlap. Never grow live row height for children. Enforce
seven children. Keep Critical as a real child with independent dates and
emit a truthful line primitive. Calculate locked Duration from date-only
semantics. Remove Above/Below labels and keep Auto logic in Core.
Preserve text-column widths under size presets and derive plot width
from the remainder. Refuse Refresh if the mutation boundary cannot be
respected. Add the orchestrator before R4.9. Preserve the existing
strong R4 ownership/reconciliation and Office isolation. Use bulk COM
operations. Add top-level tests for hierarchy/projection changes. Keep
commits small and independently verifiable.
