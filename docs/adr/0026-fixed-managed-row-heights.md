# ADR-0026 — Fixed managed row heights; live lane auto-growth is removed

- **Status:** Accepted
- **Date:** 2026-09-29
- **Relates to:** entity guide §3, §4, §9; `08-TEST-CHECKLIST.md` §B
- **Decided by:** product owner, 2026-09-29
- **Implements:** roadmap row R4.7D; register D1

## Context

The Gantt is hosted in the worksheet, beside the data the user is editing. A
lane-owning row's Gantt lane must therefore coincide with the Excel row it
belongs to, or the chart visibly detaches from the table as the user scrolls and
edits.

The landed model does the opposite. `LaneLayoutBuilder` line 170 computes:

```csharp
var laneHeight = Math.Max(metrics.LaneHeightPt, contentHeight);
```

and `08-TEST-CHECKLIST.md` §B states the same rule as an invariant: *"Lane height
grows to fit content (`max(LaneHeightPt, contentHeight)`); events are never
compressed."* Entity guide §9 describes the content-height rule in the same terms.

So the worksheet row and the lane can disagree, and the documented intent is that
the lane wins. Once a child projects onto a parent lane (ADR-0027's sibling
concern, and the R4.7B projection model), several entities share one lane, and
"grow the lane to fit" would push the chart out of alignment with the row it is
supposed to sit in. Expanding or collapsing an Excel outline group changes which
rows are visible, and a content-driven lane height makes the geometry depend on
content as well as on position.

REV5 §9 and REV6 §10 both require this to be reversed: one normal row height,
`GanttRowHeightPt`, and lane geometry flowing from the worksheet into Core rather
than the reverse.

## Decision

- **D1 — One normal-row metric, `GanttRowHeightPt`, is the authority for both the
  Excel row height and the Core lane height.** It is added to the metric-token
  catalogue (`GanttCatalogues.Metrics`), so it is a schema contract and is
  materialised to the VeryHidden configuration sheet like every other token.
- **D2 — Lane auto-growth is removed.** `LaneLayoutBuilder` no longer computes
  `Math.Max(LaneHeightPt, contentHeight)`. A lane's height is the measured
  worksheet row height for the row that owns it. The removal is explicit rather
  than a ceiling: there is no "grow if the content needs it" path to re-enable.
- **D3 — `08-TEST-CHECKLIST.md` §B's auto-growth item is replaced**, not
  deleted silently. It becomes: every visible lane-owning row's lane Top, Height
  and Bottom equal its Excel row's Top, Height and Bottom.
- **D4 — Normalisation is the add-in's job.** On Initialise, on hierarchy
  mutation, and on Refresh, managed row heights are normalised to
  `GanttRowHeightPt`. A user who drags a row taller has it restored on the next
  explicit Refresh, because the alignment is not negotiable.
- **D5 — The orientation of the data flow is inverted and stated.** The worksheet
  is the source of vertical geometry; Core consumes measured row bounds. The
  landed code derives lane geometry from content, and the scene builder is
  already documented as taking measured panel geometry as an *input* it never
  measures or defaults (ADR-0023/0024). This decision extends that existing
  discipline upward to lanes.
- **D6 — Structural rows use the same metric.** `Splitter` and `Spacer` rows
  must match their corresponding Excel row geometry. `SplitterHeightPt` and
  `SpacerHeightPt` remain, but the Excel row they describe is normalised to them,
  so Core and the worksheet cannot disagree.
- **D7 — Projected children create no lane and consume no lane height.** A child
  renders on its parent's lane. If several children cannot be kept legible
  without overlapping, the scene emits a warning or refuses; it never grows the
  lane. Growing the lane is the defect this ADR exists to remove.

## Consequences

- A lane-owning row aligns exactly with its Excel row, and that alignment holds
  through expand/collapse, because the geometry comes from the rows that are
  currently visible.
- **This deletes landed, tested behaviour** and invalidates a checklist
  invariant. Both are intended, and the checklist is updated in the same change
  rather than left to fail later.
- Content that genuinely cannot fit in the fixed height is now a **reported
  condition** instead of silently accommodated. That is a visible behaviour
  change and is why D7 specifies warn-or-refuse rather than grow.
- R4.7D depends on R4.7B: a child must have a resolved render lane before lane
  geometry can be derived from the worksheet rather than from content.
- `GanttRowHeightPt` is a new metric token, so it joins ADR-0029's schema bump.

## Alternatives considered

- **Cap growth instead of removing it** (lane grows to a maximum, then refuses).
  Rejected: a capped lane is still a lane that disagrees with its row, so the
  alignment defect survives in the exact case the cap was added to handle.
- **Keep growth and make expanded child rows own lanes.** Rejected: REV5 §3 and
  REV6 §6 both require a projected child to render on the parent's lane, and a
  child that owns a lane cannot be collapsed without removing its visual.
- **Let Excel row height be user-owned and read it live.** Rejected: then the
  chart's vertical rhythm is user-controlled, and `Duration`-and-spacing
  consistency across a schedule is lost. D4 normalises instead.
- **Keep `LaneHeightPt` as a separate minimum.** Rejected: two height authorities
  is the ambiguity ADR-0023 removed for panel bounds. `GanttRowHeightPt` is the
  single value.
