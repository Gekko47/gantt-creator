# GanttCreator R3/R4 Architecture, Remediation and Roadmap Review - Revision 5

> **Read this first (added 2026-09-29).** This document is a **tracked
> governance input**, not the binding contract. Several of its proposals were
> **changed or rejected by the product owner**. Implement
> [`PRE-R4.9-DECISION-REGISTER.md`](PRE-R4.9-DECISION-REGISTER.md) instead, and
> treat the following as settled:
>
> | This document proposes | Resolved as |
> | --- | --- |
> | §7, §16 — emit a truthful **line** primitive for Critical Interval | **Rejected.** Stays a **filled rectangle** (`CriticalFill` = `#FF0000`), half the predetermined `ActivityHeightPt`, **centred on its own visual slot** and derived from its **own** dates (owner ruling 2026-09-30 — **not** top-aligned to the parent, and the parent governs lane membership only). `CriticalLinePt` retired. |
> | §9 — remove live lane auto-growth | **Accepted**, and it reverses `LaneLayoutBuilder` plus checklist §B. ADR-0026. |
> | §12 — Row Styling | Accepted; unchanged. |
> | §15 — label positions | **Changed.** `Above`/`Below` removed (as REV6 also requires); delineator corners retained. |
> | §16 — bounded `HostShapeKey` | **Rejected.** R4.7's L18 refusal stands. |
> | §16 — `CriticalOverlayHeightRatio = 0.5` | Not a configurable token; derived from the parent bar's resolved height. |
> | §17, §18 — work packages | Rebuilt as roadmap rows R4.7A-R4.7E and R4.8A; see [`03-ROADMAP.md`](03-ROADMAP.md) revision 10. |
>
> Everything in this document not listed above stands as written.

**Status:** This document supersedes REV4 and earlier R3/R4 remediation
briefs.

It incorporates the agreed product model for parent/child hierarchy,
projected child shapes, Critical Interval rows, uniform worksheet
geometry, hidden engine metadata, Ribbon Row Styling, derived Duration,
Refresh orchestration, and the revised R4/R5 sequence.

The existing Core / Office / AddIn separation, scene-first design,
stable IDs, validation architecture, explicit Refresh model,
ownership/reconciliation, origin translation and application-state
handling should be preserved. The work below corrects and extends
hierarchy, projection, row geometry, workbook presentation and
orchestration before R4.9.

# 1. Target product behaviour

Excel is both the schedule-authoring surface and the host for the live
Gantt. The Gantt must look embedded in the sheet.

The workflow is:

``` text
Initialise
 -> enter visible schedule data
 -> Add Child / expand / collapse as required
 -> Row Styling if required
 -> Refresh
 -> maintain engine metadata
 -> calculate Duration
 -> normalise row heights
 -> validate hierarchy/data
 -> resolve projection
 -> build scene
 -> reconcile native Excel shapes
```

Normal worksheet edits, hierarchy changes and Row Styling do not render.
Refresh remains explicit.

# 2. Source rows and render lanes are different concepts

A **lane-owning entity** has a source row and its own Gantt lane. Its
lane must align exactly with its Excel row.

A **parent-projected child** has its own source row for entering Type,
Description, Start, Finish and viewing Duration, but its visual shape
renders on the parent's lane.

Example:

``` text
Activity A          -> owns Lane A
    Critical A      -> source row exists; shape projects onto Lane A
Activity B          -> owns Lane B
Milestone C         -> owns its row/lane
```

Do not assume every source row creates a plot lane.

# 3. Expanded/collapsed hierarchy

When expanded, child source rows are visible and editable. They consume
worksheet height but projected children do not create independent plot
lanes. Blank plot space opposite an expanded child row is therefore
intentional.

When collapsed, Excel outlining hides the child source rows. The child's
plotted shape remains on the parent lane. Subsequent lane-owning rows
move upward because the hidden child rows no longer consume visible
worksheet height.

**Invariant:** expand/collapse may change source-row visibility and
later lane Y positions, but must not change whether a child renders, its
dates, style, identity or parent relationship.

Separate `RenderVisible` from source-row/outline visibility. Never map
an Excel row hidden by hierarchy collapse to
`GanttEvent.Visible = false`.

# 4. Authoritative hierarchy and Excel outlining

`ParentId -> Id` is the authoritative relationship. Excel row grouping
is only its user-facing presentation.

Requirements:

-   users never edit `ParentId`;
-   GanttCreator creates and maintains it;
-   outline groups can be reconstructed from metadata;
-   insert/move/delete operations preserve hierarchy deliberately;
-   broken/ambiguous relationships block or warn before rendering;
-   multi-level nesting must be explicitly supported or explicitly
    rejected.

Use native Excel `+/-` outline controls for expand/collapse.

# 5. Introduce an explicit projection model

Do not encode all placement indirectly through `LaneId` and
`StackIndex`.

Introduce a Core concept equivalent to:

``` csharp
EntityProjection
{
    SourceEntityId
    RenderLaneOwnerId
    ProjectionMode
}
```

Recommended modes include `OwnLane`, `OverlayParent`, `StackOnParent`,
`MilestoneOnParent`, plus plot-global/structural handling where
appropriate.

Projection is resolved in Core before scene construction. Excel must not
decide where an entity belongs.

Typical mapping:

  Entity                          Projection
  ------------------------------- -------------------
  top-level activity              OwnLane
  top-level milestone             OwnLane
  Critical Interval child         OverlayParent
  child activity sharing parent   StackOnParent
  child milestone                 MilestoneOnParent
  Delineator                      PlotGlobal
  Splitter / Spacer               OwnStructuralLane

# 6. LaneId and StackIndex are engine-maintained

`LaneId` and `StackIndex` remain hidden and locked if retained
physically in the table.

Prefer deriving render-lane ownership from hierarchy/projection:

``` text
top-level lane owner -> own Id
projected child      -> resolved parent lane owner
```

Do not permit contradictory `ParentId`/`LaneId` state.

`StackIndex` is calculated deterministically from Type, projection,
child order, z-order and available fixed-row geometry. Users do not edit
it.

# 7. Critical Interval is a genuine child row

Critical Interval remains a real user-authored entity. It has its own
Type, Description, Start, Finish, Duration, stable ID and permitted
styling. It does not inherit dates from the parent.

Example:

``` text
Activity A       01 Jan ----------------------------- 30 Jun
    Critical 1       10 Jan ---- 25 Jan
    Critical 2                         03 Apr --- 18 Apr
```

Both critical rows project onto Activity A.

The critical visual is a line/line-span overlay centred on the parent
bar, approximately **50% of the normal activity-bar thickness**, above
the parent and consuming no additional lane height.

Use a Core metric such as `CriticalOverlayHeightRatio = 0.5`.

Core should emit a truthful line-like primitive. Do not emit a thin
rectangle and teach every renderer that it secretly means a line.

# 8. Generic Add Child workflow

Provide Ribbon **Add Child**.

It must:

1.  validate the selected parent;
2.  offer only permitted child Types from a Core capability matrix;
3.  insert after the parent's existing descendant block;
4.  generate the child ID;
5.  write hidden `ParentId`;
6.  initialise derived engine metadata;
7.  set Type;
8.  apply `GanttRowHeightPt`;
9.  rebuild/extend Excel outline grouping;
10. select the new row for data entry;
11. not render until Refresh.

Critical Interval uses this workflow rather than a special date dialog.

Hierarchy-aware delete/move/promote operations must preserve IDs and
references.

# 9. Uniform row height and embedded alignment

Introduce one normal-row metric: `GanttRowHeightPt`.

All ordinary visible managed rows use it. A lane-owning entity's live
scene lane must have exactly the same Top, Height and Bottom as its
adjacent visible Excel row.

On Initialise, hierarchy mutation and Refresh, normalise managed row
heights to this value subject to protection rules. If a user changes a
managed row height, Refresh restores it.

All shapes must fit inside the fixed height. Critical overlays consume
no extra height. Multiple projected entities use bounded within-row
stacking; if they cannot remain legible, warn/refuse rather than grow
the lane.

**Remove live lane auto-growth.** Live Y geometry must flow from current
visible worksheet row geometry into Core lane geometry, not the reverse.

Splitter/Spacer must also match their corresponding Excel row heights;
special structural heights are acceptable only if Excel and Core use the
same metric.

# 10. Duration is visible, derived and locked

Duration is useful to the analyst, so keep it visible, but lock it
against user editing.

GanttCreator owns the value and recalculates it on Refresh from
Type/date semantics.

Recommended rules:

-   span activity: inclusive calendar days,
    `(Finish.Date - Start.Date).Days + 1`, unless the product contract
    deliberately chooses another convention;
-   milestone / point event: `-`;
-   delineator: `-`;
-   Splitter / Spacer: blank or approved placeholder;
-   invalid/missing dates: no misleading duration; leave
    blank/placeholder and report validation.

Do not call it working duration unless a working-calendar engine exists.
At this stage define it explicitly as calendar-day duration if that is
the implemented calculation.

Prefer GanttCreator-written values rather than user-editable Excel
formulas, and use bulk writes where practical.

Test one-day spans, multi-day spans, milestones, delineators,
missing/invalid dates, Critical Intervals, child activities and
collapsed children.

# 11. Column classification, hiding and locking

Audit every table column and classify it.

**User-visible editable:** Type, Description, Start, Finish and other
genuine inputs.

**User-visible locked:** Duration and future calculated outputs.

**Engine-hidden locked:** Id, ParentId, LaneId, StackIndex, StyleKey,
fill/stroke/label overrides, projection metadata and similar
implementation state.

Near term, keep engine columns physically in `tblGanttData` if this
avoids disruptive migration, but initialise them hidden/locked and
repair their state if users unhide them.

Protection must still permit add-in-controlled Add Child, grouping, Row
Styling, row mutations, Refresh and repair.

# 12. Row Styling

Provide Ribbon **Row Styling** and **Reset Row Styling**.

Row Styling operates on the selected managed row and exposes only
capabilities valid for its entity Type/style. Persist only overrides,
preserving:

``` text
Entity Type -> default StyleKey -> configured style -> row override
```

Examples: activities may expose fill/stroke/label position; delineators
expose stroke/label controls; Critical Interval exposes line-supported
controls; Spacer exposes none.

The user never edits hidden style columns. Row Styling does not redraw;
Refresh applies the changes.

# 13. Refresh geometry must account for outline state

Collapsed children change the vertical position of later visible rows.
Therefore live scene Y geometry must be resolved from the current
visible worksheet layout.

Required sequence:

``` text
read hierarchy
 -> repair/read outline state
 -> normalise row heights
 -> measure current visible lane-owning rows
 -> resolve child projections
 -> build scene
```

A collapsed child gets no independent live lane; its shape uses its
parent's measured lane geometry.

# 14. Production Refresh orchestrator - mandatory before R4.9

Add a production orchestration service rather than placing the whole
workflow in `RefreshSheetCommand`.

Conceptually:

``` text
RefreshSheetCommand
    -> IGanttRefreshOrchestrator
        -> resolve workbook/sheet/table
        -> enforce managed column visibility/protection
        -> read source rows + hidden metadata
        -> read configuration
        -> validate/repair hierarchy
        -> maintain outline groups
        -> calculate/write Duration
        -> normalise row heights
        -> resolve source visibility + projections
        -> build style registry
        -> resolve plot range/scale
        -> measure visible worksheet geometry
        -> assemble SceneBuildRequest
        -> build LiveExcel scene
        -> translate to Office requests
        -> preflight ownership
        -> reconcile shapes
        -> return outcome
```

Also add an `ISceneBuildRequestFactory` or equivalent testable assembly
layer.

`RefreshSheetCommand` should only check availability, invoke the
orchestrator and present the outcome.

# 15. Live versus export composition

Make composition explicit. The live Excel sheet already contains the
real data cells and must never receive a duplicate scene data panel.

``` text
LiveExcel      -> IncludeDataPanel = false
EditableExport -> IncludeDataPanel = true
PowerPoint     -> IncludeDataPanel = true
Raster         -> according to export contract
```

Do not rely on a caller merely remembering not to supply a panel theme.

# 16. Remaining R4 technical work

**Critical primitive:** change current thin-rectangle critical geometry
to the truthful line-like primitive described above.

**Shape performance:** R4.7 reconciliation is functionally strong, but
repeated linear scans of Excel `Shapes` are likely to dominate large
Refresh operations. R4.10 must profile this. If confirmed, introduce a
render session that resolves the worksheet/protection once, enumerates
shapes once, builds an owned-shape dictionary, then
updates/creates/deletes/z-orders through that index.

**Host identity:** separate full logical `PrimitiveId` from bounded
Excel `Shape.Name`. Introduce a deterministic, collision-resistant
bounded `HostShapeKey`.

**Procurement hatch:** define one renderer-independent hatch
specification (angle, pitch, line width, foreground and
background/transparent behaviour) so Excel, PowerPoint and raster share
semantics.

**Refresh failure:** preflight everything possible. Validation/build
failure leaves the previous valid chart untouched. A COM failure during
mutation should return an explicit partial-refresh outcome; reassess
transactions only after R4.10 evidence.

# 17. Required vertical integration fixture

Add a non-Excel integration test:

``` text
worksheet-like DTOs + hidden metadata
 -> validation
 -> hierarchy
 -> projection
 -> domain events
 -> SceneBuildRequest
 -> SceneBuilder
 -> SceneShapeRenderer
 -> OfficeShapeRequest[]
```

Include planned, actual, baseline, delay, procurement, milestone,
Critical Interval child, ordinary projected child, child milestone if
supported, Splitter, Spacer, Delineator, render-disabled entity,
collapsed-but-render-visible child, Row Styling override and clipped
event.

Assert lane ownership, projection, primitive kind, geometry, style,
z-order and stable identity.

# 18. Revised implementation roadmap

## R4.7A - hierarchy and projection contract

Implement before R4.9:

-   source-row versus render-lane distinction;
-   authoritative ParentId hierarchy;
-   EntityProjection or equivalent;
-   projection capability matrix;
-   derived lane ownership;
-   derived StackIndex;
-   source visibility separate from render visibility;
-   Critical Interval projection;
-   Core hierarchy/projection tests.

## R4.7B - workbook presentation contract

-   classify every column;
-   hide and lock engine columns;
-   keep Duration visible and locked;
-   repair hidden/locked state;
-   maintain existing schema-version/migration discipline.

## R4.7C - row geometry and outline integration

-   implement `GanttRowHeightPt`;
-   normalise managed row heights;
-   reconstruct Excel groups from hierarchy;
-   handle expanded/collapsed source rows;
-   measure current visible lane-owning rows;
-   derive fixed lane geometry from Excel row bounds;
-   remove live lane auto-growth;
-   project children onto parent geometry.

## R4.7D - derived Duration

-   Core Duration calculator;
-   Type/date-specific rules;
-   locked visible column;
-   bulk Refresh write;
-   milestone/delineator/structural display rules;
-   validation-safe invalid-date handling.

## R4.7E - Critical primitive correction

-   line-like scene representation;
-   50% thickness metric;
-   parent-centred geometry;
-   Excel line translation;
-   vertical contract tests.

## R4.8 - preservation and safety

Extend existing preservation tests to hidden/locked engine columns,
outline state, collapsed children, user-created shapes, unrelated
cells/formatting, ownership collisions and forced mid-reconciliation
failure.

## R4.8A - production orchestration

Before R4.9 implement the Refresh orchestrator, scene-request factory,
plot-anchor/bounds resolution and explicit LiveExcel composition
profile.

## R4.9 - first complete live Refresh

R4.9 becomes integration, not architecture invention. Acceptance flow:

``` text
Initialise
 -> enter visible data
 -> Add Child / expand / collapse
 -> optional Row Styling
 -> Refresh
 -> metadata maintained
 -> Duration recalculated
 -> heights normalised
 -> hierarchy/projection resolved
 -> scene built from current visible worksheet geometry
 -> native shapes reconciled
 -> Gantt aligned with visible lane-owning rows
```

## R4.10 - performance

Instrument table read, metadata read, hierarchy resolution, Duration
calculation/write, outline/row normalisation, geometry measurement,
validation, scene build, translation, shape enumeration, updates,
creates, deletes, z-order and total Refresh. Test both expanded and
collapsed realistic schedules.

# 19. Revised R5 roadmap

R5 should now prioritise user workflows rather than exposing engine
columns.

**Hierarchy:** Add Child, permitted child-Type selection, Delete
Child/Row, hierarchy-safe Move Up/Down, Expand, Collapse, Expand All,
Collapse All, hierarchy repair.

**Row styling:** Row Styling, Reset Row Styling, capability-gated
controls, invisible persistence, Refresh-only visual update.

**Normal Gantt controls:** plot range, scale, global styles, delineator
creation/editing, label controls, selection context and validation
navigation.

Do not expose ParentId, LaneId, StackIndex or StyleKey as a shortcut
around implementing these workflows.

# 20. Updated priority list

## P0 - before R4.9

1.  Source-row versus render-lane model.
2.  Parent/child hierarchy contract.
3.  Projection model.
4.  Critical Interval as user-authored projected child.
5.  Critical line primitive and \~50% thickness.
6.  Source visibility separate from render visibility.
7.  `GanttRowHeightPt`.
8.  Exact worksheet/lane alignment.
9.  Remove live lane auto-growth.
10. Excel outline/group integration.
11. Hidden and locked engine columns.
12. Visible locked derived Duration.
13. Duration calculator.
14. Refresh orchestrator.
15. Scene-build request factory.
16. Live composition profile.
17. Full vertical integration fixture.

## P1 - before R5/export hardens

18. Add Child Ribbon workflow.
19. Row Styling and Reset Row Styling.
20. Hierarchy-safe move/delete controls.
21. Bounded host-shape identity.
22. Indexed render session if R4.10 confirms need.
23. Procurement hatch fidelity.
24. Explicit partial-refresh UX.

## P2 - hardening

25. Hierarchy repair edge cases.
26. Transactional Refresh reassessment.
27. Configuration repair hardening.
28. Documentation consolidation after implementation stabilises.

# 21. Definition of done for R4

R4 is not complete until live Excel demonstrates:

**Worksheet** - uniform configured normal row height; - Duration visible
and locked; - engine columns hidden and locked; - only intended
authoring cells editable.

**Hierarchy** - parent can have multiple child rows; - child rows have
independent data/dates; - children group/indent beneath parent; -
collapse hides source rows; - expand reveals them; - collapse does not
stop child shapes rendering; - grouping can be rebuilt from metadata.

**Geometry** - every visible lane-owning row aligns exactly with its
Excel row; - projected children render on parent lane; - projected child
rows do not create independent lanes; - no lane auto-growth breaks
alignment; - all shapes fit within fixed row geometry.

**Criticality** - Critical Interval is a real child row; - it uses its
own Start/Finish; - it renders as a parent overlay line; - line
thickness is approximately half the normal bar; - multiple critical
intervals can project onto one parent; - collapsing the parent leaves
overlays rendered.

**Derived Duration** - recalculated on Refresh; - one-day and multi-day
spans correct; - milestones/delineators show approved non-duration
value; - invalid dates do not produce misleading values.

**Styling** - Row Styling persists invisible overrides; - Reset restores
inheritance; - styling does not render until Refresh.

**Refresh** - validation/build failure preserves last valid chart; -
user/unowned shapes are preserved; - application state restored; -
repeated Refresh idempotent; - live renderer never covers worksheet
cells with export-panel shapes.

**Performance** - representative large schedule profiled; - COM
bottlenecks quantified; - required indexed-session optimisation
completed before later features depend on poor scaling.

# 22. Prescriptive instructions to the coding LLM

1.  Treat this revision as the current product behaviour for R3/R4
    corrections.
2.  Do not remove Critical Interval as a user entity.
3.  Do not equate collapsed Excel rows with render-hidden entities.
4.  Do not expose engine IDs/layout fields to avoid building proper
    hierarchy UI.
5.  Resolve hierarchy and projection in Core.
6.  Treat Excel grouping as presentation, not authoritative data.
7.  Derive live Y geometry from current visible worksheet rows.
8.  Keep lane-owning rows exactly aligned with Excel.
9.  Do not allow live lane auto-growth.
10. Keep projected children inside the parent lane.
11. Calculate Duration from Type/date semantics and persist it as locked
    derived data.
12. Use bulk COM reads/writes where practical.
13. Add the Refresh orchestrator before R4.9.
14. Keep `RefreshSheetCommand` thin.
15. Keep Row Styling as hidden overrides.
16. Do not render on edits, outline changes or Row Styling; Refresh
    remains explicit.
17. Correct Critical Interval semantics in Core, not with Excel-only
    special cases.
18. Add top-level composition tests for hierarchy/projection changes.
19. Preserve the existing strong R4 ownership/reconciliation, origin
    translation and application-state work.
20. Keep commits small and mergeable.
21. Do not regenerate golden fixtures merely to silence failures;
    approve semantics first.

# 23. Target architecture

``` text
VISIBLE EXCEL AUTHORING
(Type / Description / Start / Finish / locked Duration)
                |
                v
HIDDEN + LOCKED ENGINE METADATA
(Id / ParentId / LaneId / StackIndex / StyleKey / overrides)
                |
                v
VALIDATION + HIERARCHY
                |
                v
PROJECTION RESOLUTION
       +--------+---------+
       |                  |
       v                  v
OWN-LANE ENTITY      PROJECTED CHILD
       |                  |
       |             parent lane owner
       +--------+---------+
                |
                v
CURRENT VISIBLE EXCEL ROW GEOMETRY
                |
                v
SCENE
                |
                v
NATIVE EXCEL SHAPES
```

Excel outline state controls whether child **source rows** are shown. It
does not control whether valid child **Gantt entities** render.

This preserves the worksheet-embedded appearance while supporting
independent child durations, Critical Interval overlays, collapsed
source data, hidden engine state, Ribbon styling and future editable
export from one deterministic scene model.
