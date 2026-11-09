# Host-owned vertical elements — design for the ADR-0052 carve-out

## The report this answers

> "The charts vertical elements; bands and delineators, do not stretch with the expand and
> collapse groups like they used to."

They used to, and the reason is recorded. ADR-0036 D2 measured all four shape families as
natively `xlMoveAndSize` on creation, so Excel resized them with the rows. ADR-0052 then
forced `Placement = xlMove` on every owned shape, because the host also moves **and
resizes** `xlMoveAndSize` shapes on row insert/delete/move/hide/show/re-height, none of which
changes the scene — the 2026-11-03 evidence was a splitter band the scene called 18 pt while
the host showed four rows, then a thin line. `xlMove` removed the resize half and the stale
geometry with it, and removed the live stretch at the same time.

So this is not a lock defect and not a regression to chase: it is a deliberate trade that
needs to be narrowed, not reverted.

## Decision to record (ADR-0052 amendment, new D-numbers)

**D6 — One shape class is host-owned: the primitives that span the plot's full vertical
extent.** Identified from scene geometry alone, with no new schema field and no per-type
list:

- a vertical `SceneLine` whose Y interval **covers** `[PlotBounds.Top, PlotBounds.Bottom]`
  (delineator lines, vertical grid lines);
- a `SceneRect` whose `Top` is at or above `PlotBounds.Top` and whose `Bottom` is at or
  below `PlotBounds.Bottom` (alternate-period bands, the plot frame).

Coverage is not a tightening of "equal" for its own sake; the first implementation used
equality with the plot rectangle and **did nothing on the live chart**, because no real
primitive has plot-bounds geometry. `PlotSpanGeometry.LiftTopIntoHeader` moves a band's top
up into the header by `PlotBandHeaderOverlapPt` (ADR-0037) and `ClosingLineBottomPt` extends
a delineator's bottom through the anchor row by `anchorRow + boundaryWidth/2` (ADR-0038),
with `SceneBuilder` handing both to the delineator. The three regression tests that pinned
this — a lifted band, an extended band, an extended delineator — failed against the
equality predicate and pass against containment.

Everything else — bars, milestones, labels, horizontal band edges, splitter/spacer rules —
keeps `Placement = xlMove` exactly as ADR-0052 says. A horizontal line is not in the class:
it is row-anchored, and `xlMove` translation is already what it needs.

**D7 — The class is created and kept at `xlMoveAndSize`.** `ExcelShapeWriter.ApplyGeometry`
chooses the placement from the request's class rather than always writing `xlMove`, and
`Update` re-applies it only when the host reports the other value (ADR-0054 O3's conditional
write stays).

**D8 — Verification trusts the host for that class.** `GeometryMatches` today requires
`Placement == xlMove` and a 0.01 pt bounding-box equality. For a host-owned shape it
requires `xlMoveAndSize`, compares `Left`, `Top`, `Width` within the existing 0.01 pt, and
compares `Height` within a **new, stated tolerance** (`HostResizeTolerancePt`, 0.5 pt) that
absorbs Excel's proportional-resize rounding. A difference inside the tolerance is a match
and no write is issued; a difference outside it is an update, because the scene's own
measured height is the authority again at that moment.

**D9 — The scene's height for that class is the measured (visible) height.** It already is:
`ExcelPanelGridMeasurement` sums per-row `RowHeight`, and a hidden row reports 0. So after a
refresh the scene and the host agree by construction; between a collapse and the next
refresh the host alone has resized, which is the behaviour being restored. No new scene
rule, no new field.

**D10 — Export parity is by construction, and is pinned.** PNG and PowerPoint consume the
same scene, and a scene build measures the sheet as it is at that moment, so an export taken
while a group is collapsed shows the collapsed chart — the same geometry the host is
showing. The invariant to test is therefore "live shape bounds == export composition bounds
for the host-owned class, before and after a collapse", not a second geometry rule.

## Why the alternatives were rejected

- **Full revert to `xlMoveAndSize` everywhere** — reinstates exactly the stale-geometry class
  ADR-0052 exists to prevent (band at 18 pt while the host shows four rows). Rejected.
- **Scene-owned height with a token fallback for hidden rows** — restores the full height but
  only on the next refresh, and it deliberately makes the chart taller than the visible rows,
  so a collapsed group would show bands running past the last visible row. Rejected by this
  design because the owner asked for host-speed stretching; it stays available as a fallback
  if the carve-out churns.
- **Two-edge cell anchoring** — Excel has no such API; a shape has one anchor cell plus
  `Placement`. `xlMoveAndSize` over the covered range is the closest mechanism and is what
  ADR-0036 measured.

## Known limits to write into the ADR

- Excel resizes proportionally, so repeated collapse/expand cycles can drift by a fraction of
  a point; `HostResizeTolerancePt` absorbs it and the next refresh corrects anything larger.
- ADR-0037 measured that an insert at or above the shape's `TopLeftCell` row **translates**
  the shape instead of resizing it, so a row inserted above the plot leaves the verticals
  one row short until the next refresh. Unchanged by this decision, and worth restating.
- A row-height drag that does not raise a cell change is now visible immediately on the
  verticals and corrected by the next refresh; previously they stayed stale until then.

## Touch points

| Layer | File | Change |
|---|---|---|
| Core | `Scene/` (read-only model) | none — the predicate is computed from existing bounds |
| Core | new `Scene/HostOwnedGeometry.cs` (or a method on `Scene`) | `IsHostOwned(primitive, plotBounds)` predicate, pure and unit-tested |
| Office | `IShapeWritePort`/`OfficeShapeRequest` | the request carries whether the host owns its extent (one bool, set by the scene→Office mapper) |
| Office | `ExcelShapeWriter.ApplyGeometry` | placement chosen from that flag instead of the fixed `xlMove` |
| Office | `ExcelShapeWriter.GeometryMatches` | per-class placement check + `HostResizeTolerancePt` on height |
| Office | `ShapeReconciler` | passes the class through; no new policy there |
| Tests | Contract | placement per class; verification accept/refuse at the tolerance boundary; a legacy `xlMoveAndSize` bar still heals to `xlMove` |
| Tests | Core | predicate: full-height rect yes, full-height line yes, horizontal line no, bar no, label no |
| Tests | Office integration | collapse a group → delineator `Height` shrinks with **no refresh**; expand → returns; a Refresh then reports zero geometry writes for that shape |
| Tests | Export | live bounds == export composition bounds for a host-owned shape, expanded and collapsed |

## Gates that must prove it

- `verify-quick.ps1` per commit; `verify.ps1` at branch end.
- `verify-office.ps1` in full: this changes host behaviour that only the live gate can
  confirm, and it must re-prove ADR-0052's clause (row-insert tracking still works) and
  ADR-0037's (a band above the insertion point still translates).
- Mutation gate: the new tolerance branch and the predicate need positive tests that fail
  when inverted.
