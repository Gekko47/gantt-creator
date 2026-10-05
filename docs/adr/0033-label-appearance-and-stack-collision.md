# ADR-0033 — Labels: no fill, no border, no margin, one row tall, and stacks are not collisions

- **Status:** Accepted
- **Date:** 2026-10-02
- **Relates to:** entity guide §5, §6, §9, §22, §23; ADR-0015; ADR-0018; ADR-0030
- **Decided by:** product owner, 2026-10-02
- **Supersedes:** nothing. **Amends:** entity guide §22's candidate-acceptance
  sentence, and the `LabelHeightPt` token's role as the label box height.

## Context

Four defects were reported together from one live workbook.

**1. Every label was a black box.** `ExcelSceneBuildRequestFactory` built the label
style as `new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000"))`.
The colour belongs on the TEXT, and the comment directly above already said "the
code-owned `DefaultText` token" — only the named argument was wrong.
`OfficeStyleMapper.CarriesFill` reports a `TextBox` as fill-capable, so the token
reached Excel as an opaque black rectangle with default black text on it.

**2. Labels were too short for their text, and stacks were truncated.** With two
events on one lane, `LaneLayoutBuilder` gives each a slot. An 18 pt row split by
`LanePaddingTop/BottomPt` = 3 and `StackGapPt` = 2 yields roughly 5 pt per slot,
while the label box was `LabelHeightPt` = 10 and is centred on its slot. The two
boxes therefore always intersected, `LabelPlanner.Blocked` rejected every candidate,
and both labels fell through to the ADR-0015 widest-gap fallback and were ellipsised
to a few characters. A stack is a deliberate layout, not an overlap.

**3. The year and period bands were inverted.** `FrameBandsBuilder.TryBuild` builds
the period band directly above the plot and the year band above it, and returns that
in `ChartFrameGeometry`. Its own `AddHeaders` then recomputed both bands from the
plot's top edge **in the opposite order**. `SceneBuilder` supplied the panel header
bottom as `plotBounds.Y - YearBandHeightPt`, a third answer. Three sources, one
question. On screen: month labels in the upper row, `2025` in the lower one.

**4. The plot range was day-granular.** The extent was
`earliest.AddDays(-padding)` … `latest.AddDays(padding)` with `RangePaddingDays` = 7,
so a 10 Jan – 10 Aug schedule began on 3 Jan and a bar could start part-way through a
month column.

## Decision

**D1 — A label is never painted.** Fill and stroke are suppressed by shape KIND in
`OfficeStyleMapper`, not by the token being null. A named style may legitimately
carry a stroke (`SplitterBuilder`'s label style does); whether a label box is ever
painted is a property of the entity. The rule is structural so it cannot regress
through a token. The four `TextFrame2` margins are zeroed in `ApplyText`, so the text
starts exactly where the scene placed the box rather than at Excel's default inset.

**D2 — A stack sibling does not block.** `Blocked` skips an occupant that names the
same lane and a *different* stack index. The exemption requires BOTH: a same-slot
collision, a cross-lane neighbour, and any occupant with unknown identity all still
block exactly as before. This is a deliberate, owner-approved amendment to §22's
"no intersection with … higher-priority labels".

The boxes still overlap vertically — measured at **11 pt** for a two-stack lane — and
that is now expected: the box is a full row tall, so it necessarily spans both slots.
Horizontal separation is what keeps the labels legible, and
`Stacked_labels_overlap_vertically_by_design_but_are_both_emitted_whole` pins that
both are emitted whole and never share horizontal space.

**D3 — The label box is one worksheet row, centred on the row.** `LabelHeightPt`
(10 pt, its own token) no longer feeds it; `GanttRowHeightPt` (18 pt) does. The
member is renamed `RowHeightPt` rather than left misnamed. Centring is on the
SHAPE, which is itself centred on its slot, so a stacked event still centres on its
own slot rather than the whole lane.

**D4 — The plot extent is snapped to whole months.** `plotStart` is the month start
containing `earliest − padding`; `plotFinish` is the month end containing
`latest + padding`. The order is load-bearing and the owner's two examples pin it:

- 10 Jan − 3 = 7 Jan, still January → **1 Jan**
- 3 Jan − 3 = 31 Dec, which is December → **1 Dec**

Padding-then-snap and snap-then-pad disagree on the first case, and only this order
satisfies both. A consequence worth recording: a range starting exactly on the 1st
also escapes to the previous month, by the same rule as the owner's 3 Jan case.

`RangePaddingDays` is **retained** and its default reduced 7 → 3. At 7 the pad was
wider than the gap the snap resolves and would have pushed 10 Jan back to 27 Dec,
contradicting D4's first example. It remains observable: a 31-day pad on 15 March
widens the range to whole February and April.

## Consequences

- The label box is taller than its bar, so it overhangs by 5 pt top and bottom. Any
  collision test comparing against the owning shape's own rectangle must exempt it;
  `A_row_height_box_is_not_blocked_by_the_owning_shape_it_is_centred_on` pins that.
- A fixture that measures a row height SHORTER than `GanttRowHeightPt` cannot contain
  a row-height label box, and every label is suppressed with
  `LabelSuppressedNoSpace`. `ExcelSceneBuildRequestFactoryTests` now reads its row
  height from the catalogue for this reason.
- `AddHeaders` no longer recomputes band placement, so it cannot disagree with
  `TryBuild` again. `The_emitted_header_primitives_agree_with_the_returned_geometry`
  pins the emitted primitives, which the pre-existing geometry-only test could not see.
- The golden scene snapshot records the OLD inverted band arrangement and requires
  human-reviewed regeneration. The verified delta is eight primitives: the four
  period bands and their four labels move down by `PeriodBandHeightPt`, and the year
  band and its label move up. No bar, milestone, or row label changed.