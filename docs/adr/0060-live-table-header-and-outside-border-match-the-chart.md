# ADR-0060 — The live table's header type and outside border match the chart

- **Status:** Accepted (owner ruling 2026-11-14)
- **Date:** 2026-11-14
- **Amends:** the 2026-11-13 ruling recorded in guide §3 that left the visible header's bold
  flag and every cell border's weight to Excel's table style
- **Related:** ADR-0007 (catalogue hash), ADR-0016 (date display format), ADR-0029 D6 (no
  migration), ADR-0030 D5 (the header row is the period band's row), ADR-0055 (sheet
  protection), R5.18d, R5.18n

## Context

The header row of `tblGanttData` *is* the chart's period band row (ADR-0030 D5). Two rows
therefore sit on one line, and they disagreed about how the line should read.

- The period band is drawn from `FrameTheme`'s `PeriodHeader` style: **9pt, regular**.
- The header row was written from `GanttTableAppearance.HeaderFontSizePt`, which resolved
  the `HeaderFontSizePt` typography token: **8pt**, and its bold flag was never written at
  all, so Excel's table style kept the header **bold**.

The result, seen on the live host, was one heading in two typefaces and two weights: a bold
8pt header under a regular 9pt band. The same 2026-11-13 ruling that deferred the bold flag
also deferred every cell border's weight, so the table's own outside edge stepped down from
the chart frame's 1pt `MajorBoundaryPt` to whatever weight the table style supplied — a
visible seam where the frame meets the table on all four sides.

## Decision

1. **The live header row takes the period band's type: 9pt, regular.**
   `GanttCatalogues.Typography`'s `HeaderFontSizePt` value becomes `9` (its `isBold` flag was
   already `false`), and the Office adapters **write the bold flag explicitly** rather than
   leaving it to the table style. Writing `false` is what makes the header regular; a flag
   never written is a flag Excel keeps.
2. **The table's outside edge takes the chart frame's weight: 1pt.**
   `GanttTableAppearance.OuterBorderWeightPt` reads the frame's structural
   `MajorBoundaryPt` token rather than restating `1`, and the adapters write Excel's
   `xlThin` — which is Excel's 1pt cell-border weight — on the four edge indices only.
   Interior borders keep Excel's table-style weight and line style, exactly as the
   2026-11-13 ruling requires: the add-in recolours the grid, it does not redraw it.
3. **Both writes live in one seam per adapter** (`ApplyHeaderType`, `SetBorderWeight`), so the
   initialiser and the recovery mirror cannot disagree about the header's type or the
   edge weight.

## Consequences

- **The catalogue hash moves.** `ComputeCatalogueHash` covers the typography rows by design
  (ADR-0007 D6), so the pinned hash advances to
  `b8845b938a243ef918c37a92bf845a8b4677bb08557d7431393184aac8e689e4`. **No column, row, or
  setting contract changes**, so `CurrentSchemaVersion` does not move; this is the first pin
  advance driven by a catalogue *value*. A workbook written before the change reports
  `CatalogueHashMismatch`, which the Repair command auto-classifies; the full remedy is
  Initialise (ADR-0029 D6, no migration).
- **Excel's border weights are a closed set, not a point width.** The 1pt frame default maps
  exactly to `xlThin`; a user-raised frame width (R5.6d) is not mirrored on a cell border.
  The coupling is guarded by a Core fact that pins `OuterBorderWeightPt` to 1.
- **The export composition is unchanged.** The panel header it draws is a `SceneStyle`
  supplied by the caller's panel theme; this ADR rules on the live worksheet's header row.
- **No new setting, command, or schema row.** The change is a catalogue value, two adapter
  seams, and the facts that pin them.
