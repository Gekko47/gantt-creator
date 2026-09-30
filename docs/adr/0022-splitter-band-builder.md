# ADR-0022 — The §10 splitter band is a new scene builder on the existing style pipeline

- **Status:** Accepted
- **Date:** 2026-09-28
- **Relates to:** entity guide §2, §10, z-order table row 30; ADR-0007; ADR-0018;
  R3.10; R3.11; R3.12; ADR-0021
- **Depends on:** ADR-0021

## Context

Entity guide §10 describes a `Splitter` as occupying "a complete lane across the
included data panel and plot" with `SplitterFill` and a "major top/bottom border",
and permits four label positions — `DataPanelLeft` (the default), `PlotCentre`,
`Both`, and `None` — where `Both` "deliberately creates two scene text entities with
stable role-derived IDs". No Core code implemented any of that: §10 geometry was
unreachable (ADR-0021), and the whole of §10's visible output had to be written.

Two candidate designs were available: special-case the band inside `SceneBuilder`, or
give it a builder like every other emitting entity. The guide's own §10 wording and
the presence of an unused `ZLayer.Section = 30` made the second the coherent choice.

## Decision

- **D1 — A new `SplitterBuilder`, not a `SceneBuilder` special case.** It follows the
  landed `PanelBuilder`/`DelineatorBuilder` shape: a typed request record, `TryBuild`
  returning a typed outcome, a `SplitterRefusal` enum, and a positive test for every
  refusal member.
- **D2 — The band spans panel-left to plot-right.** §10 says "across the included
  data panel and plot", so the band's X extent is the panel's left edge to the plot's
  right edge. A plot-only band would leave the section break visibly incomplete on the
  data-panel side.
- **D3 — Band at `ZLayer.Section` (30), borders at `ZLayer.Frame` (80).** The guide's
  z-order table already assigns 30 to "splitter/section backgrounds". The borders sit
  above it so they frame the band rather than being covered by it.
- **D4 — The border width is `MajorBoundaryPt`, not the row's outline width.** §10
  calls it a *major* boundary, and the `Splitter` preset carries no outline token, so
  borrowing a style-derived width would tie a structural measurement to a style the
  guide never defined it from.
- **D5 — Labels are placed by the builder, using the shared `ITextMetrics` seam.**
  `LabelPlanner`'s §22 cascade places a label beside an entity's own bounds, which is
  a different problem from "left edge of the data panel". This mirrors how
  `DelineatorBuilder` owns its corner labels, and keeps one measuring seam for the
  whole scene.
- **D6 — A `Splitter`/`Spacer`/`Delineator` style comes from the code-owned catalogue
  preset when the registry cannot resolve one.** `GanttStyleResolver` deliberately
  refuses a blank `StyleKey` rather than guessing, and these three types carry no
  named style in the workbook registry, so requiring registry resolution would refuse
  a valid workbook. A previous delineator fallback restated `#404040` inline; that
  literal is deleted and the `DefaultDelineator` preset is read instead, so the
  `DelineatorStroke` token remains the single source of truth.
- **D7 — The lane top is offset to chart coordinates exactly once, at the call site.**
  `LaneLayoutBuilder` is lane-relative (its first lane starts at `y=0`), so passing
  `lane.Top` raw placed the band in the header bands and put its label outside the
  chart — caught by `SceneValidator`'s `LabelOutsideChart` finding during this work.
  The offset is applied where the builder is called, matching `BuildSpans`.

## Consequences

- §10 has a real, tested scene representation for the first time.
- `ZLayer.Section` is no longer an unused enum member; the scene-validator and
  equivalence surfaces now have a splitter to account for.
- The `Splitter` row's `LabelPosition` cell is honoured, and the four permitted
  positions are validated against `EntityTypeCatalog` rather than accepted blindly.
- A blank description emits the band but no label, matching the rule `PanelBuilder`
  already applies to a blank cell: blank is legal data and an empty text primitive is
  noise a renderer would have to special-case.
- No new token was introduced. `SplitterFill`, `SplitterHeightPt`, and
  `MajorBoundaryPt` already existed in the code-owned catalogue; the gap was that
  nothing consumed them.
