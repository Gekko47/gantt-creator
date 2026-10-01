# GanttCreator - Remaining Fixes and Relevant Context

This document lists only the fixes and verification items remaining from
the latest code review and subsequent product clarifications.

## 1. Fix stale DTO state after ID repair - R4

**Context:** Refresh may read the table, repair IDs in Excel, and then
validate the pre-repair in-memory row snapshot.

**Required fix:** If identity repair changes workbook data, re-read the
managed table or return an updated authoritative snapshot before
validation.

**Required behaviour:** A newly assigned or repaired ID must be
recognised during the same Refresh.

------------------------------------------------------------------------

## 2. Verify and correct Critical Interval rectangle geometry - R4

**Context:** Critical Interval is a genuine child entity with its own
independently entered Start and Finish. It renders on the parent's Gantt
lane but does not inherit the parent's dates.

**Required fix/verification:**

-   Keep Critical Interval as a `SceneRect` / rectangular bar.
-   Horizontal start comes from `CriticalInterval.Start`.
-   Horizontal finish comes from `CriticalInterval.Finish`, using the
    normal inclusive finish-date geometry.
-   `ParentId` determines the render lane only.
-   The Critical Interval rectangle is approximately 50% of the normal
    activity-bar height.
-   **The Critical Interval rectangle uses the same top alignment as its
    parent activity. It is not vertically centred on the parent.**
-   Multiple Critical Interval children may independently overlay the
    same parent.
-   A Critical Interval may be shorter than its parent and must not be
    stretched to the parent's duration.
-   Plot clipping uses the Critical Interval's own date span.
-   Keep the rectangle representation consistently through Excel,
    PowerPoint and raster rendering.

Do **not** convert Critical Interval to a line.

------------------------------------------------------------------------

## 3. Provide an R5-compatible lane-anchor seam - R4 architecture

**Context:** R5 will implement expanded/collapsed hierarchy. Expanded
child source rows occupy Excel rows but projected children do not own
separate Gantt lanes.

**Required fix:** R4 must provide or preserve a clean mechanism by which
R5 can later supply the measured Excel Top and Height for each
lane-owning row.

Conceptually:

``` text
LaneOwnerId
TopPt
HeightPt
```

R4 may continue using sequential lane placement where explicit anchors
are not supplied.

Do not implement the full expanded/collapsed row-alignment workflow in
R4; that belongs in R5.

------------------------------------------------------------------------

## 4. Finalise LiveExcel vertical-extent behaviour - R4

**Context:** Horizontal output size is governed by A4 / 16:9 / 4:3
presets, while the live Gantt must remain embedded beside worksheet
rows.

**Required decision/fix:** Define centrally whether the live plot/grid:

-   ends with the managed schedule-row region; or
-   extends vertically to the complete selected preset canvas.

The decision must live in the scene/layout request contract, preferably
`SceneBuildRequestFactory`, rather than individual renderers.

------------------------------------------------------------------------

## 5. Complete and verify Refresh orchestration - R4.9

**Context:** `RefreshSheetCommand` should remain a thin entry point.

**Required pipeline:**

``` text
Read
 -> identity maintenance
 -> re-read if identity repair mutated workbook
 -> validate
 -> calculate/write Duration
 -> resolve styles
 -> resolve size/range/scale
 -> measure workbook geometry
 -> build SceneBuildRequest
 -> build scene
 -> translate to Office requests
 -> ownership/preflight
 -> reconcile shapes
```

Do not add R5 Add Child/Delete/Expand/Collapse/Row Styling UI
responsibilities to the R4 orchestrator.

------------------------------------------------------------------------

## 6. Make the workbook mutation boundary explicit and enforce it - R4

**GanttCreator may mutate:**

-   managed Gantt table/schema metadata;
-   engine metadata;
-   Duration;
-   GanttCreator defined names;
-   `_GanttCreatorConfig`;
-   GanttCreator validation artefacts;
-   GanttCreator-owned shapes;
-   other explicitly managed state belonging to the current development
    stage.

**GanttCreator must not mutate:**

-   unrelated worksheets;
-   unrelated cells;
-   unrelated shapes;
-   unrelated defined names;
-   unrelated formatting;
-   unrelated outline groups.

**Required behaviour:** If Refresh cannot safely continue without
changing unowned workbook content, return an explicit error rather than
overwriting it.

------------------------------------------------------------------------

## 7. Formalise entity-to-primitives identity - R4

**Context:** One entity can create multiple scene/Office objects.

**Required contract:**

``` text
Entity -> 0..N ScenePrimitives
```

Use deterministic primitive roles for:

-   activity bar;
-   label;
-   milestone;
-   Critical Interval rectangle;
-   other entity-specific graphics.

Do not assume one entity equals one Excel shape.

This contract will also support later selection, PowerPoint and editable
export.

------------------------------------------------------------------------

## 8. Separate logical PrimitiveId from bounded Excel Shape.Name - R4/R4.10

**Context:** Excel shape names are bounded while scene identities should
not be constrained by Excel host limitations.

**Required fix:** If full `PrimitiveId` is still used directly as
`Shape.Name`, introduce a deterministic bounded host key.

Requirements:

-   stable across Refresh;
-   collision-resistant;
-   within Excel limits;
-   ownership-verifiable;
-   safely associated with the full logical primitive identity.

------------------------------------------------------------------------

## 9. Verify date-only normalisation throughout Refresh - R4

**Required contract:**

-   Start and Finish are date-only values.
-   Time-of-day has no semantic meaning.
-   Finish remains inclusive for span entities.
-   Duration and scene geometry use the same normalised date values.

Verify consistency through:

``` text
Excel read
 -> validation
 -> Duration
 -> scene geometry
```

------------------------------------------------------------------------

## 10. Verify Duration is derived and non-user-authored - R4

**Context:** Duration remains visible but is calculated by GanttCreator.

**Required behaviour:**

-   Span activity: inclusive calendar-day duration.
-   Critical Interval: inclusive calendar-day duration from its own
    Start/Finish.
-   Milestone/delineator: approved non-duration value.
-   Invalid/missing dates: no misleading Duration.
-   Prefer bulk writes.
-   Duration is not an editable user field.

------------------------------------------------------------------------

## 11. Remove any remaining Above/Below label capability - R4

Supported label positions should be:

-   Auto;
-   Inside;
-   Left;
-   Right;
-   None.

Check:

-   enums;
-   Type capability definitions;
-   validation;
-   configuration;
-   persisted-value parsing;
-   tests.

Do not retain Above/Below through a legacy capability path.

------------------------------------------------------------------------

## 12. Verify Auto label placement uses horizontal free-space evaluation - R4

**Required behaviour:** Auto evaluates available horizontal regions and
chooses the largest valid non-overlapping space.

Candidates include:

-   inside;
-   left;
-   right.

Consider:

-   plot bounds;
-   entity geometry;
-   other labels;
-   projected child geometry;
-   milestones and other relevant scene primitives.

Use a deterministic fallback if no region fully fits.

Do not increase row height to solve label placement.

------------------------------------------------------------------------

## 13. Verify output-size layout preserves text-column widths - R4

Initial presets:

-   A4;
-   Presentation 16:9;
-   Presentation 4:3.

**Required behaviour:**

``` text
GanttPlotWidth =
    PresetWidth
    - measured text/data column width
    - approved margins/chrome
```

The text/data columns retain their existing widths.

Do not silently resize them.

If insufficient width remains for a valid Gantt area, return an explicit
layout error.

Title and time-header geometry must derive from the same composition.

------------------------------------------------------------------------

## 14. Ensure LiveExcel never renders export-panel shapes over worksheet cells - R4

**Context:** The live Excel worksheet already contains the actual data
cells.

**Required behaviour:** `LiveExcel` composition explicitly excludes
generated data-panel rectangles/text.

Export composition may generate the data panel later.

This must be an explicit composition contract rather than an informal
caller convention.

------------------------------------------------------------------------

## 15. Profile reconciliation before further optimisation - R4.10

Measure separately:

-   workbook read;
-   metadata/identity work;
-   validation;
-   Duration;
-   geometry;
-   scene construction;
-   scene-to-Office translation;
-   shape enumeration;
-   shape update;
-   shape creation;
-   shape deletion;
-   z-order;
-   total Refresh.

**Context:** Repeated linear scans of Excel's `Shapes` collection may
become the dominant cost on large schedules.

If profiling confirms this, introduce one indexed render session rather
than repeated COM collection scans.

------------------------------------------------------------------------

## 16. Finish ownership/reconciliation behaviour before export renderers depend on it - R4/R4.10

Create/update/delete/z-order operations must act only on positively
identified GanttCreator-owned shapes.

Required behaviour:

-   user-created shapes remain untouched;
-   unowned same-name shapes remain untouched;
-   malformed/foreign ownership tags do not grant ownership;
-   repeated Refresh is idempotent;
-   reconciliation remains deterministic.

------------------------------------------------------------------------

# Items intentionally not required before R4.9

These are planned later features and should not be treated as current R4
defects.

## R5

-   Add Child Ribbon UI;
-   Delete/Promote Ribbon UI;
-   Excel outline Expand/Collapse;
-   exact expanded-child Excel/Gantt row alignment;
-   hierarchy-aware Move Up/Down;
-   Row Styling / Reset Row Styling UI;
-   row-based hierarchy/group selection behaviour.

## Later protection/hardening

-   final managed worksheet protection;
-   protected outlining;
-   final enforcement of hidden/locked engine columns under managed
    protection;
-   complete duplicate worksheet/table reseed Office workflow where not
    already required;
-   deeper workbook/configuration repair;
-   transactional Refresh, unless performance/reliability evidence later
    justifies it.

## Not required

-   custom Undo system.
