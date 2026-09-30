# ADR-0023 — The panel measures per row, and the panel's bounds are derived

- **Status:** Accepted
- **Date:** 2026-09-28
- **Relates to:** entity guide §3, §4, §10; R3.11; R3.12; R3.15; ADR-0021
- **Depends on:** ADR-0021
- **Plan:** [`REMEDIATION-PLAN-B.md`](../REMEDIATION-PLAN-B.md)

## Context

Three defects in the Core panel contract, each confirmed against source before this
ADR was written.

1. **`PanelCellGrid` carried one `RowHeightPt` for the whole body.** `PanelBuilder`
   positioned cells at `bodyTop + (rowIndex * grid.RowHeightPt)`. §3 requires the
   panel to reproduce "the exact measured cell bounds in points", so a body whose rows
   differ in height could not be reproduced at all — and the live adapter therefore
   *refused* a mixed-height table rather than misplacing it.
2. **The panel header's height was the body row's height.** §4 says the header
   "follows live header-cell bounds"; it is a separate Excel row with its own height.
3. **The panel's bounds had two authorities.** `SceneBuildRequest.PanelBounds` fed
   `FrameBandsBuilder` while `PanelBuilder` independently derived the panel from the
   grid. Nothing checked they agreed, so the chart background could be sized from one
   rectangle while the panel was drawn in another.

A fourth defect was found while implementing the first three: **panel rows were derived
from `placements.Placements`**, so any row that is not an ordinary lane-placed event —
`Splitter`, `Spacer`, `Delineator`, or a hidden row — could not appear in the data
panel at all. ADR-0021 made the structural rows lane participants, so this coupling
became visibly wrong rather than merely latent.

## Decision

- **D1 — `SceneBuildRequest.PanelBounds` is removed.** The panel's bounds come from
  `PanelBuildResult.PanelBounds` and nowhere else, and the frame consumes that derived
  value. The panel is therefore built **before** the frame, which is possible because
  `PanelBuilder` needs only the grid, the projected rows, the plot bounds, and the
  period-header bottom. The alternative — keeping both inputs and adding a consistency
  validator — was rejected: it would preserve two numbers that must agree instead of
  removing one.
- **D2 — `PanelCellGrid` carries `RowHeightsPt` and `HeaderHeightPt`.** `RowHeightPt`
  is removed, not deprecated, so a later consumer cannot quietly reintroduce a sample
  height.
- **D3 — Heights are positional and the count is validated.** A new
  `PanelBuildRefusal.RowCountMismatch` fires when the projected row count and the
  measured height count disagree. This is the guard that stops a source row which
  failed validation from shifting every panel cell below it.
- **D4 — Panel rows come from a new pure `PanelRowProjection`.** Every validated source
  row gets a panel row, in `RowNumber` order — worksheet order, not the plot's
  `SortOrder`-then-row lane order. The two are deliberately different and the
  distinction is commented at the call site.
- **D5 — Hidden rows appear in the panel.** §3's "visible approved columns" qualifies
  the *columns*; the field table says `Visible` "suppresses entity and its label", not
  the data-panel row. This is a product-observable change, stated here rather than
  discovered in a diff.
- **D6 — No compatibility shim.** The single-height API is not preserved, and the
  Office adapter was changed only enough to compile against the new grid, replicating
  its one confirmed-uniform height. Per-row measurement is a separate change.

## Consequences

- A worksheet with mixed row heights becomes renderable, which was the point.
- `PanelBuildRefusal.NonPositiveHeaderHeight` was **removed** as unreachable: the grid
  validates the header height, so a second check could never fire. Its test moved to
  `PanelCellGridTests`.
- A zero-row panel is no longer constructible. The grid requires at least one measured
  body row, so "header with no body" — which reads as a data table that lost its data
  — is refused rather than rendered. `tblGanttData` always has a body.
- **The golden scene's chart bounds change** and are presented for approval: the
  reference scene is deliberately panel-free, and it previously carried an arbitrary
  `PanelBounds` solely to make the chart wide. With the input gone, a panel-free chart
  is the plot plus its header bands.
- The Office measurement adapter still refuses a mixed body. The Core contract can
  express one; the adapter cannot yet measure one. That gap is the next change, and
  stating it here is the point — the contract being right is what makes it visible.
