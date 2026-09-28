# Remediation implementation plan — Commit B (panel geometry contract)

Follows [`REMEDIATION-PLAN-A.md`](REMEDIATION-PLAN-A.md) (Commit A, `4537b86`).
**This plan is written before any code change**, per the traceability requirement.

## Problem

Four defects in the Core panel contract, all confirmed against source before this
plan was written.

1. **`PanelCellGrid` carries one `RowHeightPt` for every body row.** `PanelBuilder`
   lays out `bodyTop + (rowIndex * grid.RowHeightPt)` (`:259`) and
   `bodyTop + ((index + 1) * grid.RowHeightPt)` for borders (`:338`). §3 requires
   "the exact measured cell bounds in points", so a worksheet whose rows are not all
   the same height cannot be reproduced. `ExcelPanelGridMeasurement` therefore
   *refuses* a mixed-height table (`HasUniformRowHeight` → `InvalidMeasurement`), so
   a legitimate workbook is unrenderable.
2. **The panel header height is the body row height.** `SceneBuilder` passes
   `request.Grid!.RowHeightPt` as `HeaderHeightPt`. §4 says the header "follows live
   header-cell bounds" — a separate measurement. The two being equal is an accident.
3. **Two authorities for the panel's bounds.** `SceneBuildRequest.PanelBounds` is
   caller-supplied and consumed by `FrameBandsBuilder` (`:153`), while `PanelBuilder`
   independently derives the panel from `grid.TotalWidthPt` and the row heights.
   Nothing checks they agree, so the chart background can be sized from one rectangle
   while the panel is drawn in another.
4. **Panel membership is derived from lane placement.** `SceneBuilder` builds panel
   rows from `placements.Placements.Select(...)`, so a `Splitter`, `Spacer`,
   `Delineator`, or hidden row cannot appear in the data panel at all. After Commit A
   the structural rows *do* reach lane layout, so this coupling is now visibly wrong
   rather than merely latent.

## Decisions (recorded before implementation)

- **D-B1 — `SceneBuildRequest.PanelBounds` is removed.** The panel's bounds come from
  `PanelBuildResult.PanelBounds` and nowhere else. `FrameBandsBuilder` keeps its
  `PanelBounds` *builder input* — it consumes the derived value rather than measuring a
  second one. To make that possible the panel is built **before** the frame:
  `PanelBuilder` needs only the grid, the rows, the plot bounds, and the
  period-header bottom, all known before the frame exists. This is the "one
  authoritative source" the brief asks for, chosen over the alternative of keeping
  both inputs and adding a consistency validator.
- **D-B2 — `PanelCellGrid` carries ordered per-row heights and its own header
  height.** `RowHeightPt` is **removed**, not deprecated: a single sample height is
  the wrong abstraction, and keeping it would let a later consumer quietly
  reintroduce the bug. New members: `RowHeightsPt` (ordered, one per body row, each
  finite and positive) and `HeaderHeightPt` (finite and positive).
- **D-B3 — Ordinal alignment, validated.** The grid's ordered heights align to
  `PanelBuildRequest.Rows` **by position**. A count mismatch is a new typed refusal,
  `RowCountMismatch`, with a positive test. This is the guard that stops a dropped
  source row from silently shifting every panel row: §3 reproduces the table, so a
  grid measuring a different number of rows than the projection produces is a
  disagreement, not something to guess around.
- **D-B4 — Panel rows come from a new, explicit Core projection.** A new pure
  `PanelRowProjection.TryProject` owns the source-row → panel-row mapping. Panel rows
  are **every** validated source row in `RowNumber` (i.e. table) order — not lane
  order. `LaneOrdering` sorts by `SortOrder` then `RowNumber`, and the panel must
  follow the worksheet, so the two orders are deliberately different; the distinction
  is commented at the call site.
- **D-B5 — Hidden rows appear in the panel.** §3's source is "`tblGanttData`, visible
  approved columns, current Excel column widths, and row heights" — "visible"
  qualifies *columns*. The field table says `Visible` "suppresses entity and its
  label", not the data-panel row, and §3 requires the panel to reproduce each included
  cell. **This is a product-observable change**: a `Visible=false` row now renders a
  panel row where it previously did not. It is the guide-faithful reading and is
  stated here rather than slipped in.
- **D-B6 — No compatibility shim.** The old single-height API is not preserved. The
  brief is explicit that this is the cheap moment to correct the boundary, and an
  adapter mapping a per-row grid onto one height would defeat the change.

## Steps

1. Write this plan. **Done.**
2. `PanelCellGrid`: replace `RowHeightPt` with `RowHeightsPt` + `HeaderHeightPt`;
   validate each row height and the header height; add typed refusals (`NoRows`,
   per-row `NonPositiveRowHeight`, `NonPositiveHeaderHeight`).
   **Every new refusal ships its positive test in the same commit.**
3. `PanelBuilder`: per-row cumulative layout for bodies, borders, and `panelBottom`;
   read the header height from the grid and drop the request member; add
   `RowCountMismatch`.
4. New `PanelRowProjection` with a typed outcome, the source-row ordering rule, and
   the cell projection currently living in `SceneBuilder.Cells`.
5. `SceneBuilder`: build the panel first, pass the derived bounds to the frame, drop
   `SceneBuildRequest.PanelBounds`; move `BuildSplitters` after the panel so it reads
   the derived panel left edge.
6. Update every caller and fixture: `SceneBuilderTests`, `PanelBuilderTests`,
   `PanelCellGridTests`, `ReferenceSceneBuilder`, `SceneBenchmarkTests`, and the
   Office `ExcelPanelGridMeasurementTests` **contract** tests. The live adapter is
   Commit C, but the contract tests must compile against the new grid.
7. Tests for the named cases the brief requires: Splitter, Spacer, Delineator,
   Critical Interval, hidden row, several rows sharing one `LaneId`, several stacked
   events on one lane; 12pt/24pt/15pt producing exactly three body heights; panel
   bottom equal to the sum of measured heights plus header geometry; equal heights
   reproducing today's geometry exactly (no regression).
8. **Golden gate.** The reference fixture's rows are all uniform, so the golden may
   move only through the *membership* change (D-B5 / D-B4). Whether it does is
   **verified during implementation**, not assumed: if it moves, the delta is
   presented for approval and regenerated in its own commit, per D4.
9. Docs: ADR-0023, entity guide §3/§4, REPO-MAP, STATUS, and the R4.6 work item
   (whose style work is blocked on the panel contract being correct).

## Non-vacuity

Every new refusal gets a positive test. The per-row layout tests assert **exact**
cumulative positions for a mixed-height fixture, because a tolerance here would accept
a layout that quietly used a sample height — the precise defect being fixed. The
"equal heights reproduce today's geometry" test is the regression pin for the common
case.

## Explicitly not in this commit

No Office adapter change (that is Commit C). No `IShapeWritePort` change. No
reconciliation (R4.7 proper). The read-only protection policy is untouched — it is
Commit C and an ADR-0008 decision.
