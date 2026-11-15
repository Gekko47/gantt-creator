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
> neutral fill. **One weight, measured twice:** the first pass drew the plot's
> rectangle at the minor weight (0.5pt) on the theory that only the frame should
> be heavy; the rendered sheet then showed the plot's border visibly lighter than
> the table's lines beside it, so the second pass gives **every border the major
> weight** — table lines, header cells, the plot's rectangle and the frame — and
> leaves only the interior grid between periods at the minor weight. The scene
> baseline is regenerated for the band borders, the plot's top line and the plot
> rectangle's weight, with this note as the stated reason; no golden image is
> involved.

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
