# ADR-0052 — Verify the live shape before skipping a scheduled write

- **Status**: Accepted (implementation 2026-11-03; amended 2026-11-08 by D6–D10)
- **Date**: 2026-11-03
- **Supersedes**: none
- **Relates to**: R4.7 D5 (full property diff rejected), R5.11 (coalesced
  automatic refresh and delta reconciliation), ADR-0019 (ownership tag),
  ADR-0036 D2 (native `xlMoveAndSize`), ADR-0037 (anchor-row translation),
  ADR-0054 O3 (conditional placement write),
  `GanttCreator.Office.ShapeReconciler`, `GanttCreator.Office.ExcelShapeWriter`

## Context

R5.11 gives the scheduled refresh a payload-signature cache: a SHA-256 of each
desired `OfficeShapeRequest` is compared with the last successful snapshot, and
an unchanged payload is not rewritten. The manual Refresh bypasses the snapshot.

The signature proves what the **scene** wants. It says nothing about what the
**host** holds. Shapes are host objects, and Excel moves and resizes
`xlMoveAndSize` shapes whenever a worksheet row is inserted, deleted, moved,
hidden, shown, or re-heighted — none of which changes the scene, and all of
which can happen between two refreshes. Live evidence 2026-11-03: the scene
declared the splitter band at 18pt on every refresh while the live band was four
rows tall in one state and a thin line in another; the scheduled refresh skipped
the band on `sig=match`, and only a manual Refresh (which bypasses the snapshot)
repaired it.

R4.7 D5 rejected a *full property diff* as unsound: a `null` text colour, font,
or hatch means "leave the host alone", so intent is unobservable after a write.
That reasoning does not extend to geometry.

## Decision

A scheduled write is skipped only when **both** the payload signature matches
**and** the live shape is verified. The verification rule is per shape class:

- **Scene-owned (default).** The owned shape exists and carries its ownership tag;
  its `Placement` is `xlMove` (a legacy `xlMoveAndSize` shape is not trusted); and
  its `Left`, `Top`, `Width`, `Height` equal the scene's bounding box within
  `GeometryMatchTolerancePt` (0.01pt, above `float` rounding, below display
  resolution).
- **Host-owned (D6).** The owned shape exists and carries its ownership tag; its
  `Placement` is `xlMoveAndSize`; `Left`, `Top`, `Width` equal the scene's bounding
  box within `GeometryMatchTolerancePt`; and `Height` equals the scene's bounding
  box within `HostResizeTolerancePt` (0.5pt). The wider height tolerance absorbs
  Excel's proportional-resize rounding on collapse/expand cycles; a difference
  inside it is a match and no write is issued, a difference outside it is an
  update because the scene's measured height is the authority again.

`IShapeWritePort.GeometryMatches` carries the check without exposing an interop
type; `ShapeReconciler` consults it for every candidate the cache nominated.
The signature remains the cheap first filter: a hash comparison instead of five
COM reads per unchanged shape, and a rewrite costs far more than either.

Scene-owned shapes are placed with `xlMove` on create **and** update, so the host
may still move a shape with its cell (the behaviour the row inserter relies on)
but may not resize it. Host-owned shapes are placed with `xlMoveAndSize` so the
host may resize them with the rows. Placement is re-applied on update only when
the host reports the other value (ADR-0054 O3), so existing workbooks heal on
their next Refresh.

### D6 — Host-owned class (amendment 2026-11-08)

One shape class is host-owned: the primitives that span the plot's full vertical
extent. The class is identified from scene geometry alone, with no new schema
field and no per-type list:

- a vertical `SceneLine` whose Y interval **covers** `[PlotBounds.Top,
  PlotBounds.Bottom]` (delineator lines, vertical grid lines);
- a `SceneRect` whose `Top` is at or above `PlotBounds.Top` and whose `Bottom`
  is at or below `PlotBounds.Bottom` (alternate-period bands, the plot frame).

**Coverage, not equality — corrected 2026-11-08 after the first implementation
did nothing on the live chart.** The builders never emit plot-bounds geometry
for these primitives: `PlotSpanGeometry.LiftTopIntoHeader` moves a band's top
edge up into the header row by `PlotBandHeaderOverlapPt` (ADR-0037), and
`ClosingLineBottomPt` extends a delineator's bottom through the reserved anchor
row by `anchorRow + boundaryWidth/2` (ADR-0038), with `SceneBuilder` supplying
both to the delineator. An equality predicate therefore rejected every real band
and every real delineator and left them on `xlMove` — the carve-out was inert
while unit tests using idealised bounds passed. Containment is also what keeps
the class honest in the other direction: a bar occupies one lane, a label its
measured text height, and a horizontal rule has no vertical extent, so none can
cover the plot however the rows are sized.

Everything else keeps `xlMove`. A horizontal line is not in the class: it is
row-anchored and `xlMove` translation is already what it needs.

### D7 — Placement

`ExcelShapeWriter.ApplyGeometry` chooses the placement from the request's class
rather than always writing `xlMove`. `Update` re-applies it only when the host
reports the other value.

### D8 — Verification

For a host-owned shape, `GeometryMatches` requires `xlMoveAndSize`, compares
`Left`, `Top`, `Width` within `GeometryMatchTolerancePt`, and compares `Height`
within `HostResizeTolerancePt` (0.5pt). The tolerance is stated, not tuned: it
is above Excel's proportional-resize rounding and below any height a real row
change produces.

### D9 — Scene height

The scene's height for that class is the measured (visible) height. A hidden row
reports 0, so after a refresh the scene and the host agree by construction;
between a collapse and the next refresh the host alone has resized, which is the
behaviour this carve-out restores. No new scene rule, no new field.

### D10 — Export parity

PNG and PowerPoint consume the same scene, and a scene build measures the sheet
as it is at that moment, so an export taken while a group is collapsed shows the
collapsed chart — the same geometry the host is showing. The invariant to test is
"live shape bounds == export composition bounds for the host-owned class, before
and after a collapse", not a second geometry rule.

## Consequences

- A scheduled refresh converges to the manual-refresh state by construction, at
  a few property reads per unchanged shape instead of a rewrite.
- The interim rule "any structural write forces a full reconcile" is retired: it
  made the cache stop paying on exactly the refreshes users perform most.
- The check is not a full property diff. Geometry has no "leave alone" member,
  so the R4.7 D5 objection does not apply; text and style remain unverified and
  are still reconciled by payload signature.
- A host that moves a shape between the verification read and the write can
  still leave one pass stale; the next refresh verifies again. The window is the
  same one R4.7 D3 already accepts for non-transactional host writes.
- `xlMove` changes a host behaviour the row inserter documented as measured
  (`Placement = xlMoveAndSize`). Insert-tracking depends on the *move* half,
  which is preserved; the Office integration matrix must re-prove row-insert
  movement after a redeploy.
- Host-owned shapes regain live stretch on collapse/expand at the cost of
  accepting host geometry for that class. Excel resizes proportionally, so
  repeated cycles can drift by a fraction of a point; `HostResizeTolerancePt`
  absorbs it and the next refresh corrects anything larger.
- ADR-0037's anchor-row rule still holds: an insert at or above the shape's
  `TopLeftCell` row **translates** the shape instead of resizing it, so a row
  inserted above the plot leaves the verticals one row short until the next
  refresh. A row-height drag that does not raise a cell change is now visible
  immediately on the verticals and corrected by the next refresh.
