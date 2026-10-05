# ADR-0034 — Lanes are anchored to their worksheet rows, the content block is centred, and the critical interval is top-aligned

- **Status:** Accepted
- **Date:** 2026-10-02
- **Relates to:** entity guide §3, §9, §12, §16; ADR-0026; ADR-0027; ADR-0030; R4.7D; R4.7H; R4.8A
- **Decided by:** product owner, 2026-10-02
- **Supersedes:** nothing. **Amends:** the stacked lane composition inherited from R3.4 (D1), the slot-packing rule that left `LanePaddingBottomPt` geometrically inert (D2), and the critical interval's slot-centring from ADR-0027 (D3).

## Context

Two independent Core geometry defects were confirmed from a live workbook and from arithmetic against shipped token defaults (`GanttRowHeightPt` 18, `LanePaddingTopPt` 3, `LanePaddingBottomPt` 3, `ActivityHeightPt` 8, `StackGapPt` 2):

**D1 — row anchoring (proved by live experiment).** `LaneLayoutBuilder` stacked lanes as `laneTop += lane.Height` from the plot top, so lane *k* coincided with body row *k* only while every row above it had opened a lane of its own. A row that owns no lane — a `Delineator` (§24 makes it plot-global), a projected child (ADR-0026 D7) — consumed no height, so every lane below it rendered one row too high and the leftover space was stranded at the bottom of the plot. ADR-0030 made the plot top the first body row's measured top but left this stacking in place, so the chart started in the right place and still drifted row by row.

**D2 — row centring (proved by arithmetic).** The placement loop packed the content block from `laneTop + LanePaddingTopPt`, so a single 8pt slot in an 18pt row occupied `laneTop+3..laneTop+11`: its centre was 7 rather than the row's 9, every shape and label sat 2pt high, and the visible space below was 7pt while `LanePaddingBottomPt` said 3. That token had no geometric effect at all — it was read only by the overflow warning's number. The pinning test asserted the produced literals (`Top = 3, Centre = 7`) rather than the relationship, so the suite was green through the defect. The defaults are self-consistent for two entities (2 × 8 + 2 = 18 = row height exactly); the single-entity case — the common case — was the one never centred.

## Decision

- **D1 — A lane's vertical geometry is its measured worksheet row's.** A LIVE composition anchors every lane to the row it renders on (`LaneRowAnchorResolver`, positional by body-row ordinal; `PanelCellGrid.RowTopsPt` derived once so anchor and panel cannot disagree). `lane Top/Height/Bottom` equal the owning row's measured bounds, and a body row that owns no lane leaves its own band empty instead of pulling every later lane upwards. An EXPORT composition has no worksheet rows and passes no anchors, keeping the stacking it has always had. A failed anchor resolution is a typed refusal (`InvalidRowAnchors` / `UnresolvableLaneAnchor`), never a silent fall back to stacking: that fallback is precisely how a lane ends up one row from its row with nothing reporting it.
- **D2 — The content block is centred in the lane; both paddings are real insets.** `slotTop = laneTop + LanePaddingTopPt + free/2` where `free = max(0, laneHeight − LanePaddingTopPt − LanePaddingBottomPt − blockHeight)` and `blockHeight` is the content plus inter-slot gaps. A single-slot row lands at `top = 5, centre = 9` — the row's own centre exactly. A full-height block (`free = 0`) is unmoved. `Splitter`/`Spacer` lanes fill their fixed height whole and are exempt: they carry no stack content to centre.
- **D3 — The critical interval is top-aligned; everything else stays centred (owner ruling 2026-10-02).** The critical rect (half the predetermined `ActivityHeightPt`) shares its activity shape's top edge — the top a full-height activity in its slot would have — rather than the slot centre, so the two read as one band when the child overlays its parent. Bars, diamonds, and labels are unchanged: centred on their slot.
- **Hidden rows are in scope.** Excel reports a hidden row's height as zero (a collapsed outline group hides its child rows), and zero is a real measurement: it is accepted by the grid, contributes no space to `RowTopsPt`, and is skipped when the lane height is derived (a body of entirely hidden rows is refused as `NoVisibleRows`). A lane-owning row that measures zero is refused (`DegenerateRowHeight`) — a visible lane cannot sit in a hidden row, and that disagreement is reported rather than laid out.

## Consequences

- Every visible lane-owning row's lane `Top/Height/Bottom` equal its Excel row's exactly (ADR-0026 D3 made representable by ADR-0030, made actual here). Non-lane rows leave a vacant band; the plot no longer drifts.
- Single-entity rows move down 2pt to their row centre; full two-slot rows are byte-identical. The golden snapshot moves and is regenerated in its own dedicated, human-reviewed commit.
- `LanePaddingBottomPt` becomes geometrically live for the first time.
- The critical interval moves from slot-centred to top-aligned (2pt up for the default 8pt activity); its height, fill, and lane membership are unchanged.

## Alternatives considered

- **Constant offset correction for D1** (shift the whole scene by the measured delta). Rejected: it fixes the top and leaves every internal gap wrong — a non-lane row in the middle would still pull later lanes up. The anchor must be per lane, not per chart.
- **Centring each slot independently rather than the block.** Rejected: it would spread stacked slots apart and move the full-height case that currently fits exactly. The block is the unit that is centred; inter-slot gaps are content.
- **Keeping the critical interval centred.** Rejected by the owner: a centred 4pt rect sits in the middle of its 8pt parent bar, reading as a stripe rather than a band. Top-alignment is the explicit exception, not the general rule.
