# ADR-0030 — The live chart is anchored to the worksheet rows, not to page coordinates

- **Status:** Accepted
- **Date:** 2026-10-01
- **Relates to:** entity guide §1, §2, §3, §4, §5, §6, §9; ADR-0023; ADR-0026; R4.7D; R4.7H; R4.8A
- **Decided by:** product owner, 2026-10-01
- **Supersedes:** nothing. **Amends:** the live vertical layout of ADR-0026 D3 and the
  page-coordinate vertical composition inherited from R4.7H.

## Context

Refresh renders a live Excel chart that is visibly **not aligned with the activity
rows**: a bar belonging to worksheet row 2 is drawn roughly 2.4 rows below that row.
This was reported from a live workbook (7 activity rows, chart correct in every
respect except its vertical position).

The cause is a composition model that is right for a page and wrong for a worksheet.
Verified against source, not inferred:

1. `ExcelSceneBuildRequestFactory.Create` sets the plot's top offset to
   `TitleBandHeightPt + YearBandHeightPt + PeriodBandHeightPt` — 24 + 18 + 16 =
   **58 pt** — and the plot height to `preset.HeightPt − top − chrome`, i.e. the
   A4/16:9 **page** height. Both are properties of the size preset, and neither
   reads the worksheet.
2. `GanttRefreshOrchestrator` renders with `ChartOriginDelta.Identity`, so scene
   coordinates reach Excel unchanged. "58 pt below the origin" therefore means
   58 pt below the top of row 1.
3. The first body row's real top is the header row's height below row 1, and the
   header row is never normalised — `ExcelRowHeightNormaliser` walks only
   `table.DataBodyRange`. That is Excel's default, about 15 pt.
4. Net offset ≈ 43 pt ≈ 2.4 rows at `GanttRowHeightPt` = 18. This matches the
   reported screenshot.

The model cannot self-correct. `PanelCellGrid` carries column widths, per-row
heights and the header height but **no absolute sheet position**; `LaneMetricsResolver`
derives a lane *height* only; `LaneLayoutBuilder` stacks lanes down from
`plotBounds.Top`; and `SceneBuilder` adds `plotBounds.Top` to every slot centre.
There is no input anywhere that says where row 1's bottom edge is, so the required
equality is not merely unmet — it is currently unrepresentable.

This violates two contracts that are already landed and tested:

- entity guide §4 — the panel header's height aligns exactly with the period
  header's bottom;
- entity guide §9 and ADR-0026 D3 — `Excel Top == Scene lane Top`,
  `Excel Height == Scene lane Height`, `Excel Bottom == Scene lane Bottom`.

Horizontal placement is **not** affected: the plot's left edge is the measured panel
width plus chrome, which lands on the first column past the table because the hidden
engine columns contribute zero width.

A second defect sits behind the first. The plot's height is the preset page height
(~531 pt for A4 landscape) regardless of row count, so for 7 rows the chart extends
roughly 400 pt past the last activity row. Fixing only the top offset would leave a
chart that starts in the right place and still runs off the bottom.

## Decision

- **D1 — The live plot's top is the first body row's measured top.** Not a page
  offset, not a token sum. The worksheet is the source of vertical geometry; the
  scene consumes it. This is ADR-0026 D5's data-flow direction applied to the
  chart's own vertical origin, which until now was the one quantity still derived
  from the preset.
- **D2 — The live plot's height is the measured total height of the visible body
  rows.** The preset governs horizontal budget and export composition only. A live
  chart that outruns its own table is the same class of defect as one that starts in
  the wrong place, and fixing D1 alone would leave it.
- **D3 — The measured sheet origin becomes an input.** `PanelCellGrid` gains the
  absolute top of the first body row and the absolute left of the panel, measured by
  the Office adapter. This is additive: no existing member changes meaning, and the
  measured grid remains an input-only model that never measures.
- **D4 — One inserted row above the table, sized to `YearBandHeightPt`.** It carries
  the table title across the panel columns and the year band across the plot
  columns. The title stops being a drawn chart entity in the live view and becomes
  the table's own title.
- **D5 — The period band occupies the table's header row.** The header row is
  normalised to `PeriodBandHeightPt` so the period band's bottom is the header's
  bottom, which is the first body row's top. This is entity guide §4's existing
  requirement; the header row is the correct host for the band rather than a second
  empty row.
- **D6 — There is no live title band.** `ShowTitle` keeps its meaning for export
  composition, where the title entity still exists and is still drawn. In the live
  profile the title is the table title cell and no title shape is emitted.
- **D7 — The size preset no longer bounds the live vertical extent.** `SizePreset`
  and `PlotGeometryResolver`'s width formula are unchanged; the resolver's
  page-height containment check applies to export composition, not to a live sheet
  whose height is however many rows the user has.

## Row count: two band rows, one inserted row

D4/D5 reserve two band rows in total, but only **one** is inserted. The period band
needs no row of its own because guide §4 already requires the panel header's height
to align with the period header's bottom — the header row *is* the period band's
row. Two physically inserted rows would leave the header row stranded between the
period band and the first body row, which is precisely the gap this ADR exists to
close.

The resulting live stack, top to bottom:

| Worksheet row | Content across panel columns | Content across plot columns |
| --- | --- | --- |
| inserted above table | table title (D4) | year band (D4) |
| table header row | column headings | period band (D5) |
| body row *n* | that row's cells | lane *n* (D1) |

## Product-owner ruling, and one recorded divergence

The owner ruled for option C: *"the top reserved row carries the table title across
the panel columns and the year band across the plot columns; the second reserved row
is the period band."*

That is implemented as D4/D5 above — one inserted row plus the header row. **The
divergence is deliberate and recorded rather than silently taken:** the instruction
says two inserted rows, and this ADR inserts one. The reason is D5, which is not a
new preference but entity guide §4's existing, already-landed requirement that the
header's height align with the period band's bottom. A second inserted row would
contradict it. If the owner intends two physically inserted rows instead, D5 is the
decision to revisit, and the guide must be amended in the same change — not this ADR
quietly reinterpreted.

## Consequences

- A lane-owning row's lane Top/Height/Bottom equal its Excel row's exactly, which is
  what ADR-0026 promised and could not previously deliver. The chart stops drifting
  from the table as the user scrolls and edits.
- The chart grows and shrinks with the table instead of with the paper size. A
  40-row schedule on A4 landscape now renders 40 rows tall, where it previously
  rendered 29 rows tall and pushed the rest off the page.
- Initialisation writes one row above the table and the header height, so the
  workbook layout is a schema contract: existing workbooks need the approved repair
  path before Refresh can anchor correctly, and the schema version is bumped in the
  same change per ADR-0029.
- The live title is a cell, so it is editable, printable, and survives save/reopen
  without a shape. A user who wants a drawn title band above the chart in the live
  sheet no longer has one; that is the owner's ruling.

## Alternatives considered

- **Translate the whole scene by the measured offset** (keep page composition, shift
  by the row-1-to-body delta). Rejected: it fixes the top and leaves the height
  wrong, so the chart still outruns the table — and it keeps a page-shaped chart in a
  worksheet, which is the root cause rather than a symptom of it.
- **Insert two rows, one per band, and leave the header alone** (the literal
  reading of the instruction). Rejected: contradicts guide §4, which requires the
  header's height to align with the period band's bottom, and reintroduces a gap
  between the period band and the first lane.
- **Drop the year and period bands from the live view entirely.** Rejected: the
  owner asked for them to be properly placed, and §5/§6 require them.
- **Scale the lane height to fit the preset** so N rows fill the page. Rejected: that
  is lane auto-growth by another name, reversed by ADR-0026 D2, and it reintroduces
  the disagreement between a row and its lane.
- **Keep the title as a drawn band in reserved space above the year band.** Rejected
  by the owner: the title becomes a table title, not a Gantt title.

## Stop conditions

- Exact row/lane equality proves unachievable on the live host → stop and report the
  measured values. A tolerance is a decision, not an implementation detail.
- Anchoring would require editing a row the add-in does not own → stop;
  ADR-0029's ownership boundary applies.