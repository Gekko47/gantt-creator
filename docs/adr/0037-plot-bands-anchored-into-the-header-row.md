# ADR-0037 — The plot-spanning shapes are anchored into the header row so they stretch

- **Status:** Accepted
- **Date:** 2026-10-03
- **Relates to:** ADR-0036 (the insert mechanism this depends on); ADR-0026 D3; ADR-0031 D1; ADR-0033
- **Decided by:** implementation, from live-host measurement. Owner approved 0.5pt as the overlap.
- **Supersedes:** nothing. **Amends:** nothing.

## Context

ADR-0036 made Add Activity shift the worksheet by one real row, which moved the lane bars correctly. A follow-up live report: the lane bars move, but the plot's **vertical structure** — the alternate period bands and the vertical grid lines — does not stretch when a row is added at the **top** or the **bottom** of the body, so the plot area is left unpainted and the bands no longer reach the chart's footer.

This is not a defect in the insert. It is precisely what Excel does.

## Evidence

Measured 2026-10-03 on Excel 16.0 x64 by `scripts/probe-frame-stretch.ps1`, against a table-shaped fixture (header row 1, body rows 2..8, reserved padding row 9) and one tall rectangle spanning the body rows — the plot-background analogue:

| Insert at | Top | Height | Verdict |
| --- | --- | --- | --- |
| **TOP (row 2)** | **+18** | **0** | shape **slides down** |
| **UPPER (row 3)** | **+18** | **0** | shape **slides down** |
| MIDDLE (row 5) | 0 | +18 | stretches |
| LOWER (row 7) | 0 | +18 | stretches |
| BOTTOM (row 8) | 0 | +18 | stretches |
| BELOW (row 9, padding) | 0 | +18 | stretches |

**The rule:** `xlMoveAndSize` resizes a shape when rows are inserted only if the insertion point is **strictly below** the shape's `TopLeftCell` row. When it is at or above, Excel **translates** the shape and leaves its height alone. The plot bands' top edge sits exactly on the header/body boundary, so their anchor resolves to the **first body row** — the one insertion point that slides. `Shape.Placement` was already correct and needs no write (ADR-0036 D2); there is no placement value that makes a shape always stretch downward.

`scripts/probe-frame-anchor.ps1` then measured the fix, moving the shape's top edge 0.5pt **inside the header row**:

| Case | Insert at | Top | Height | Anchor | Outcome |
| --- | --- | --- | --- | --- | --- |
| baseline | row 2 | +15.75 | **0** | D2→D3 | **SLIDES** |
| **candidate** | row 2 | **0** | **+15.75** | **D1→D1** | **STRETCHES** |

Middle and bottom inserts are unaffected in both cases. Lifting the top resolves the anchor to the header row, which makes every body insertion strictly below it.

## Decision

- **D1 — the plot-spanning shapes extend `PlotBandHeaderOverlapPt` (0.5pt) UP into the header row**, and their height grows by the same amount so their **bottom is unchanged**. Applied to the alternate period bands and the vertical grid lines, which share the plot's vertical extent.
- **D2 — the chart BACKGROUND is not lifted.** It spans the whole chart including the title and panel, so its top already sits above the header row and every body insert is strictly below its anchor. Shrinking it to the plot would leave the headers unpainted. An earlier draft of this change did exactly that and `The_background_still_spans_the_whole_chart_and_is_not_lifted` exists to catch it.
- **D3 — the overlap is a new catalogue token**, `PlotBandHeaderOverlapPt`, default 0.5, range 0–4. The ceiling is what makes "sub-row" a structural property rather than a convention: it cannot reach one body row (18pt) even if configured. The minimum is 0, which reproduces the previous geometry **exactly** — asserted by `A_zero_overlap_reproduces_the_plot_bounds_exactly`, so a builder that ignored the request and always applied the default could not pass.
- **D4 — the top edge is the one that moves, never the bottom.** The plot's bottom is the boundary several entity contracts are stated against; its top is the only edge with no such meaning.
- **D5 — workbook schema 10 → 11.** A metric token is a `tblGanttMetrics` row the add-in reads, so an existing workbook's catalogue hash moves. Same class as version 10's retirement, opposite direction. No migration (ADR-0029 D6); the remedy is Initialise.

## Consequences

- **No re-render and no event handler.** The invariant that normal worksheet changes never render is preserved: Excel performs the resize itself, exactly as it already performs the bar movements.
- **The overlap is invisible.** The header paints at `ZLayer.Frame` (80) over the bands (`AlternateBand` 10) and the background (`Background` 0), and 0.5pt cannot reach far enough to be seen if a user hides the header.
- **Export is unaffected in shape, not in value.** An export has no worksheet rows to anchor to, so the overlap serves no purpose there — but it is applied uniformly rather than only to live compositions, because diverging live and export geometry would break the renderer-equivalence contract. The 0.5pt lift is therefore a real, deliberate change to exported output.
- **Golden snapshot regenerated.** Exactly **7 primitives** changed — `chart:band:1`, `chart:band:3`, and the five `chart:grid:*` lines. `ChartBounds`, `PlotBounds`, every lane, bar, label, and milestone are byte-identical, verified by a per-primitive diff rather than by eye. Per the golden test's own note this is a scene baseline, not an image, so no human image review is required; it lands in its own commit with this ADR as the stated reason.
- **Non-vacuity proven by mutation.** Forcing the overlap to 0 fails `The_plot_spanning_shapes_lift_into_the_header_row_and_keep_their_bottom` with `Expected: 99.5, Actual: 100` and the golden snapshot with the same delta.

## Alternatives considered

- **Write `Shape.Placement` explicitly** (rejected: already `xlMoveAndSize` on all four families per ADR-0036 D2, and no placement value forces a downward resize — the behaviour is driven by the anchor row, not the placement mode).
- **Extend the bottom edge downward instead** (rejected: the plot's bottom is a contract boundary; D4).
- **Overlap by a full header row** (rejected: it would paint over whatever the user replaced that row with. Half a point is enough to resolve the anchor).
- **Re-render the chart after an insert** (rejected: breaks the never-render-on-change invariant and rebuilds the whole scene to fix two rectangles).
- **Apply the overlap only to the live composition** (rejected: live/export geometry divergence breaks renderer equivalence).