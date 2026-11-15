# ADR-0061 — Initialise presents the chart surface: no gridlines, engine columns hidden first, one margin, strokes inside the frame, and one weight hierarchy

> **Amendment, 2026-11-15 — border weights and the margin size.** Three further
> live reports turned the first cut into a system instead of a set of lines.
> **D5** the chart's four margins are **4pt** (was 6): `ChartOuterPaddingPt` 6 -> 4
> and `ChartPaddingRowHeightPt` 5.75 -> 3.75, so the anchor row plus the padding
> row still equals the outer padding (ADR-0038 D4) and the leading gutter, which
> reads that token, still matches all three. The catalogue hash moves and the pin
> advances; no schema version changes. **D6** the plot's border is a **complete
> rectangle on all four sides**: a new `chart:plot:top` line closes the top edge at
> the period band's own bottom, because the left and right edges plus the closing
> line had read as "a backwards L along the bottom and right". **D7** every year
> band and every period cell is **boxed** by a new `AddHeaderBorder`, so labels
> separate like the table's own header cells instead of running together on one
> neutral fill. **D8 — one weight, and it is a hairline.** The weight was settled by
> looking twice. The first pass drew the plot's rectangle at 0.5pt on the theory that
> only the frame should be heavier; the second gave every border 1pt; the rendered
> sheet showed **no visible difference**, because at 100 % display 0.5pt and 1pt both
> occupy one pixel. `GridLinePt` and `MajorBoundaryPt` are therefore both **0.25pt**,
> Excel's thinnest line, so the frame, the plot's rectangle and every year and period
> cell share one hairline weight and the grid between periods is the same hairline.
> Three facts follow. A **cell** border cannot express 0.25pt - Excel's
> `Border.Weight` is a closed enum - so `GanttTableAppearance.OuterBorderWeightPt`
> keeps reading the frame token and the adapter maps it to `xlThin`; the table's
> outside edge is therefore as thin as a cell border goes, not as thin as the frame.
> The live stroke inset is half of the hairline. And the catalogue hash moves again
> (the pin advances a third time for values, not schema), while no schema version
> changes. The scene baseline is regenerated for the band borders, the plot's top
> line and the hairline widths, with this note as the stated reason; no golden image
> is involved.
>
> **D9 — the live frame closes on the table's column edge, and the gutter is now an
> OUTER margin (2026-11-15).** The frame's left edge used to be the leading gutter's
> outer edge, which is the worksheet's own left edge at x = 0. Excel clips a shape at
> that boundary, so a 1pt stroke showed a sliver and a 0.25pt hairline showed nothing:
> the owner reported the left border missing. `LiveDataPanelBounds` now closes the frame
> on the table's first column (`OriginLeftPt`), which is a cell boundary the host can
> draw, and the gutter reads as a page-style margin **outside** the frame. Nothing else
> moves: the right edge is still the measured panel's right edge, which is also where
> the live plot starts, and the top and bottom are still the chart bounds' padding rows.
> The rejected alternative was to inset the frame by a fixed point inside the gutter,
> which would have made the left margin 3pt against 4pt on the other three sides - the
> asymmetry this amendment exists to remove. The frame stays at `ZLayer.Frame`: moving
> it above the title layer was tried and reverted, because the entity guide's z-order
> contract names that layer and `EquivalenceThinSliceTests` enforces it.
>
> **D10 — two live anchor experiments (2026-11-15).** The left frame line plotted but
> could not be found: it was anchored in the gutter, because a shape whose top-left sits
> exactly on a column boundary resolves to the column on its left. The live frame now
> starts `GanttSheetLayout.LiveFrameAnchorNudgePt` (**0.25pt**) to the right of the
> table's first column, so the stroke lives inside a visible column. Separately, the
> bottom frame line never appeared and no shape existed at that y, so the live frame's
> bottom is now the **last body row's** bottom (`PlotBounds.Bottom`) instead of the bottom
> padding row's — the working hypothesis being that a shape anchored below every populated
> row falls outside the host's range and is never painted. That costs the 4pt bottom
> margin while it is under test; the export compositions keep the full outline. Both
> changes are live-only and both are recorded as experiments to be confirmed on the host,
> not as settled geometry.

- **Status**: Accepted
- **Date**: 2026-11-15
- **Related**: ADR-0031 D2 (padding rows), ADR-0038 D4 (anchor + padding rows sum to the outer padding), ADR-0048 (the live frame encloses the table), ADR-0059 (the table starts in column B behind a leading gutter), ADR-0060 (the live table's header and outside border)

## Context

Five defects were reported from the live host against one freshly initialised
workbook. They share a cause: the sheet around the chart was never treated as part
of the chart's presentation, and geometry that was correct as *specification* was
not correct as *host behaviour*.

1. The whole schema flashed before the engine columns hid. Initialise wrote the
   header row first and hid the engine columns only after the table had been added
   and presented, so the host painted twelve column names at the default width.
2. Excel's own gridlines ran through the chart. Nothing in the add-in ever set
   them; the frame is now one clean border, so the remaining ruling inside it reads
   as a defect.
3. The left side of the frame was missing. The frame's left edge is the leading
   gutter's outer edge, which is the worksheet's own left edge (x = 0), so half of
   a centred 1 pt stroke fell outside the sheet.
4. Column A measured 12 pt while the frame's other three margins measured 6 pt. The
   gutter read the bottom padding row's height alone (5.75), and Excel's
   characters-to-points relation is affine, so one measured correction from a
   near-one-character column landed nowhere near the target.
5. The table-to-plot boundary had no line. The plot's left edge is the table's right
   edge, and the panel's last cell fill covers everything to the left of it, so the
   edge line lost its left half; the same line also sits under the period band's
   own fill.

## Decision

### D1 — Gridlines are switched off, on the window, on the success path only

`ExcelWorkbookInitialiser` reads `Application.ActiveWindow.DisplayGridlines` and
writes `false` only when it is `true`, after every other step has succeeded. It is
a window setting, not a workbook one, so it is not part of the save/restore
discipline in ADR-0020: the owner asked for it as the sheet's presentation. Running
it last means a refused Initialise leaves the user's view exactly as it was, and
read-before-write means a converged window is never dirtied.

### D2 — The engine columns are hidden before the header row is written

`HideSchemaHiddenColumns` runs first in the mutating block, addressing whole
worksheet columns from `GanttSheetLayout.TableOriginColumnIndex` with the same
schema-index arithmetic the presentation pass uses. The presentation pass keeps
owning the locked state and every other column's visibility. `RollBackHeaderRow`
now restores those columns' visibility first and unconditionally, because a failure
at the header write or at the table Add leaves no table for
`RestoreColumnVisibility` to find.

### D3 — The gutter reads the chart's outer padding, and converges in at most two measured passes

The leading gutter is the frame's **left** margin, so it reads the token the other
three margins read: `ChartOuterPaddingPt`, which ADR-0038 D4 keeps equal to the
anchor row plus the bottom padding row. It previously read the bottom padding row
alone. Because Excel's characters-to-points relation is a fixed cell padding plus a
per-character advance, a single ratio cannot land on the target; each pass
re-measures the width it wrote and corrects again, bounded at two passes so a host
that never reports the width it was given cannot loop.

### D4 — On the live sheet, a stroke is drawn inside the rectangle it closes

This applies to the frame's four sides and the plot's two edge lines. Each centre
line moves inwards by half its own stroke width, while its **extent** still spans
the whole outline, so the corners overlap and stay closed instead of being notched.
The export compositions are untouched: nothing there abuts the plot or the sheet
edge, so their geometry stays exactly as specified, and the scene baseline does not
move.

## Alternatives rejected

- **Hide the engine columns after the table is added** (the old order): the flash is
  produced by the header write itself, so no later hide removes it.
- **Zero-width engine columns instead of hidden ones**: leaves a visible sliver and
  breaks print layout (already recorded on `ApplyColumnPresentation`).
- **Set `DisplayGridlines` before the rest of Initialise**: a later refusal would
  leave the user's view changed after a no-mutation promise.
- **Iterate the gutter correction until it converges**: unbounded against a host
  that reports nothing useful; two measured passes are enough for the affine form.
- **Emit the frame as one rectangle shape instead of four lines**: could join the
  corners perfectly, but changes the entity contract, the scene baseline, and the
  renderer's primitive set for a defect that the stroke-inside rule fixes without
  either.
- **Inset the frame on export too**: would move the export fidelity contract by half
  a point with no reported defect behind it.

## D11 — the live plot starts a measured inset inside the anchor column

**Report.** "The plot is still plotting from column G even though the anchor is in S."
The anchor measurement (D10) was correct and changed nothing visible.

**Measured, not reasoned.** `scripts/probe-plot-origin.ps1` builds a workbook with the
real column layout (table in B..R, engine columns G..R 2 characters wide and hidden,
anchor column S) and reads each shape's `TopLeftCell` on Excel 16.0:

| plot start (S.Left + x) | anchor while G..R hidden | moves with S when un-hidden |
|---|---|---|
| 0pt / 0.25pt / 0.5pt | `G3` (first **hidden** engine column) | no — 0pt |
| 1pt / 2pt / 4pt / 8pt | `S3` (the anchor column) | yes — 171pt |

**Why.** A hidden column has zero width, so the twelve engine columns and the anchor
column all share one x coordinate. Excel resolves a shape's `TopLeftCell` to the first
column whose left edge is at or before the shape's `Left`, which with that run of
zero-width columns is the first hidden one. The nudge in D10 (0.25pt) therefore landed
inside column G: the plot was anchored in a hidden column (the owner's original
observation), did not follow the anchor, and its 0.25pt stroke fell inside the table's
own 1pt outside border and disappeared.

**Decision.** The live plot starts `GanttSheetLayout.LivePlotOriginInsetPt` (2pt) to the
right of the anchor column's left edge, applied as the plot's **left chrome** in
`ExcelSceneBuildRequestFactory` (and mirrored by `RibbonStateService.MeasurePlotWidth`, so
the displayed width stays the rendered width). `ReadPlotAnchorLeft` returns the anchor
column's left edge with no inset: the inset is a width-budget decision, not a measurement.

- 2pt, not 1pt: 1pt is the exact threshold measured above, and a display-scaling change
  must not be able to push the anchor back behind a hidden column.
- 2pt, not more: it is chrome taken out of the plot's width, so it stays inside the
  chart's own 4pt outer padding. A fact asserts both bounds.
- `LiveFrameAnchorNudgePt` (0.25pt) is unchanged and is now documented as the **frame's**
  anchor nudge only. The frame closes on the table's first column, which is visible, so
  the threshold above does not apply to it.
- ADR-0031 D2's "flush against the table" is superseded for the live composition only:
  an inset was rejected in 2026-11-14 as visibly off the edge, but that ruling assumed the
  table's right edge is where the engine columns end. It is not — they are hidden.
- Export compositions are unchanged: the inset is live-only chrome, not a scene rule.

## D12 — the interior cell grid is drawn, not inherited from the table style

**Report.** "There are some table rows where the vertical line between Start and Finish does
not appear."

**Measured.** Live Excel 16.0, four independent runs. Every body row, including rows created
by the real Add-row path, holds a continuous Start|Finish rule in **both** the direct border
API and `Range.DisplayFormat`, before and after `CalculateFull`, with and without the
Type-options port. A bare worksheet-row insert below the table, with either
`xlFormatFromRightOrBelow` or `xlFormatFromLeftOrAbove`, is equally clean. There is no
formatting defect in this code base.

**Cause.** A table-style border is one device pixel painted through a style consult, and the
host does not paint every rule: on rows whose height or position rounds to a sub-pixel,
individual rules drop out. That is why the report is per ROW rather than per column, and why it
moves with zoom and scroll.

**Decision.** Initialise and Repair draw the two interior indices (`xlInsideVertical`,
`xlInsideHorizontal`) as direct formatting — a continuous thin rule in the theme's border
colour — instead of leaving the grid to the style. A direct border is painted from the cell
rather than from the style consult, so the grid stops depending on that rounding.

- `GanttTableAppearance.BorderColour`'s doctrine changes: the theme now owns the colour of
  every border **and** the grid's shape. Its remark records the measured reason.
- `ExcelWorkbookInitialiser.ApplyBorderColour` and its `ExcelWorkbookRecovery` mirror both
  draw it, through a `SetInteriorBorder` seam, because both own cell presentation; a Refresh
  still writes no cell formatting. The weight is one `xlThin`, the same as the outside edge and
  every drawn chart border (D5).
- Facts: `Initialise_draws_the_interior_grid_as_direct_formatting` (initialiser) and the
  recovery mirror's assertion in
  `Recover_creates_the_gantt_sheet_with_the_contract_header_and_table`.
- Recorded, so it is not repeated: `Range.DisplayFormat.Borders[xlInsideVertical]` read over a
  **multi-cell** range is ambiguous and reported "all lines" for a row whose per-cell edges
  were `xlLineStyleNone`; only a per-cell or direct query is trustworthy.

## Consequences

- No catalogue value, schema version, or catalogue hash moves: every decision is a
  write policy or a live-only geometry rule.
- Existing facts were updated with the reason: `SceneBuilderTests`' two live-frame
  facts now assert the stroke centre half a stroke inside the table's edge, and
  `FrameBandsBuilderTests` gains a live-inset fact and an export-unchanged fact.
- New facts: engine columns are hidden **before** the header write (an ordering
  assertion, not just a final state); gridlines are switched off once; a window that
  already hides them is not rewritten; the gutter target is the outer-padding token
  and equals the bottom margin.
- `ExcelWorkbookRecovery` is unchanged: it rebuilds a table on a sheet whose engine
  columns are already hidden, and its own passes are re-run only on damage.
- The scene baseline is **not** regenerated; the change is live-only.
- Guide sections 1, 3 and 8 record the rules.
- Office host gate: **required and still owed** — none of the five reports is
  observable outside Excel.
