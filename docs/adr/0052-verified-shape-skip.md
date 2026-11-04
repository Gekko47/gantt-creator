# ADR-0052 — Verify the live shape before skipping a scheduled write

- **Status**: Accepted (implementation 2026-11-03)
- **Date**: 2026-11-03
- **Supersedes**: none
- **Relates to**: R4.7 D5 (full property diff rejected), R5.11 (coalesced
  automatic refresh and delta reconciliation), ADR-0019 (ownership tag),
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
**and** the live shape is verified:

1. the owned shape for the identifier exists and carries its ownership tag;
2. its `Placement` is `xlMove` (a legacy `xlMoveAndSize` shape is not trusted);
3. its `Left`, `Top`, `Width`, `Height` equal the scene's bounding box within
   `GeometryMatchTolerancePt` (0.01pt, above `float` rounding, below display
   resolution).

`IShapeWritePort.GeometryMatches` carries the check without exposing an interop
type; `ShapeReconciler` consults it for every candidate the cache nominated.
The signature remains the cheap first filter: a hash comparison instead of five
COM reads per unchanged shape, and a rewrite costs far more than either.

Charts are also placed with `xlMove` on create **and** update, so the host may
still move a shape with its cell (the behaviour the row inserter relies on) but
may not resize it. Placement is re-applied on update, so existing workbooks heal
on their next Refresh.

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
