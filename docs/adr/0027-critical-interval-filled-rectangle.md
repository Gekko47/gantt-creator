# ADR-0027 — The critical interval is a filled rectangle at half the predetermined activity height

- **Status:** Accepted — **amended 2026-09-30** (see [Amendment](#amendment-2026-09-30) below; D3 is superseded, and two claims about the golden snapshot are corrected as false)
- **Date:** 2026-09-29; amended 2026-09-30
- **Relates to:** entity guide §16, §20, and the per-entity field table; `08-TEST-CHECKLIST.md` §C
- **Decided by:** product owner, 2026-09-29; geometry ruling 2026-09-30
- **Implements:** roadmap row R4.7E; register D2 and D5
- **Supersedes:** nothing. REV5 §7/§16 and REV6 §9/§16 are **rejected** — see [`PRE-R4.9-DECISION-REGISTER.md`](../PRE-R4.9-DECISION-REGISTER.md) §3.

## Amendment (2026-09-30)

The original D3 was written before the owner's geometry ruling and the code that
implemented it. Three statements below no longer match the built product. They are
corrected here rather than left standing, because a reader acting on them would
implement the rejected rule.

**1. D3 is superseded.** Height is `PredeterminedHeightPt / 2` — the resolved
style's `ActivityHeightPt` — **not** `ParentVisibleBounds.Height / 2`, and the rect
is **centred on its own visual slot**, not top-aligned to the parent.
`CriticalOverlayRequest` no longer carries `ParentVisibleBounds` at all and
`TryBuild` no longer takes a parent-bounds dictionary. The horizontal extent is the
row's **own** Start/Finish, clipped only to the plot: an interval may be shorter
than, equal to, or longer than its parent, and one lying wholly outside its parent
is drawn rather than deleted. `UnresolvedParent`, `ClippedToParentCode` and
`OutsideParentCode` were removed as meaningless; `ClippedToPlotCode` /
`OutsidePlotCode` replaced the last two. **The parent governs lane membership and
nothing else.**

**2. The Context's golden claim was false.** This ADR asserts that
`reference-scene.json` contains *no critical overlay primitive at all*. It did
contain one. The statement was inferred from the presence of the single
`CriticalIntervalOutsideParent` **warning**, which was mistaken for the absence of
a primitive — but the canonical fixture's second interval was *deleted* for falling
outside its parent, while the first still rendered. The snapshot was the evidence
that drift existed, not evidence that it was harmless.

**3. Therefore D7's "the golden scene does not change" was false too, and it did
change.** Removing the parent clip resurrected the deleted primitive, so the
snapshot went from one critical primitive to two and was regenerated in `806a632`.
R4.7E's remaining work adds a critical child to the **vertical integration
fixture** so the representation is pinned in a test that asserts on it directly,
but the golden was *not* left untouched as originally predicted.

## Context

Two input documents proposed that the critical interval stop being a rectangle.

REV5 §7: *"Core should emit a truthful line-like primitive. Do not emit a thin
rectangle and teach every renderer that it secretly means a line."* REV6 §9 is
the same instruction in stronger terms: *"Core must emit a truthful line-like
scene primitive. Do not emit a thin rectangle that each renderer reinterprets."*

The concern behind them is real. Entity guide line 736 says the critical interval
is `rect {row}:critical` in the scene but a **line** in every host renderer, and
states *"the scene rect is not the host object: draw a line along the rect's top
edge at its resolved thickness. Filling the rect is a defect."* That wording was
approved in guide revision 6 and re-transcribed into the R3.16 equivalence field
table. It is a genuine interpretation layer: a renderer that ignores the rule
draws the wrong object.

**The drift is not hypothetical.** `SceneShapeRenderer` has no critical branch;
the rect falls through to the generic `Rect` path and becomes an
`OfficeShapeKind.Rect`. The PowerPoint (Phase 7) and raster (Phase 8) renderers
do not exist yet. **Zero of three renderers implement the guide's rule.** The
guide is the only artefact asserting a line, and nothing produces one. That is
how the divergence survived: `tests/golden/scene/reference-scene.json` **did**
contain a critical overlay primitive — this ADR's original claim that it contained
none was false, and is corrected in the Amendment above.

## Decision

- **D1 — The scene primitive stays `SceneRect`.** The thin-rectangle-means-line
  proposal is rejected. The product owner's stated reason is that a rectangle is
  easier to manage the size of, and the settled rule becomes: **the rect is what
  is drawn.** No renderer reinterprets it.
- **D2 — The rect is filled.** A new colour token `CriticalFill` is added to the
  catalogue carrying **`#FF0000`**, the same red as the existing `CriticalStroke`
  token. The two tokens differ by role and share a value: a *Critical Milestone*
  keeps `CriticalStroke` as a stroke, and a *Critical Interval* fills with
  `CriticalFill`. The name matches the role, so the catalogue does not carry a
  stroke token used as a fill.
- **D3 — SUPERSEDED by the 2026-09-30 ruling; see the Amendment above. Originally:
  height is half the parent activity bar's resolved height, top-aligned.** As
  written it read: the rect's `Y` is the parent's post-clip top and its `Height` is
  `ParentVisibleBounds.Height / 2`, occupying the top half of the parent bar.
  `CriticalOverlayRequest` already carries `ParentVisibleBounds`, a `RectD` with
  the parent's resolved height, so this needs no new token and no new input.
- **D4 — `CriticalLinePt` is retired.** The separate absolute thickness token
  (default 2.25, range 0.5-12) is removed from `GanttCatalogues.Metrics`. The
  `CriticalInterval` preset keeps its `ActivityHeightPt` reference, which supplies
  the denominator for the half-height. Retiring a metric token is a schema
  contract change and joins ADR-0029.
- **D5 — `CriticalInterval` gains the `Fill` colour capability.**
  `EntityTypeCatalog` currently grants it `EntityColourCapability.Stroke` only.
  **The catalogue, the validator, and the checklist change together.** The
  checklist's current §C rule that a `FillColour` override on a critical interval
  is a blocking error is inverted, and `GanttRowValidator` stops rejecting it.
- **D6 — The three changes in D5 are one commit, not three.** A partial change is
  the failure mode: the entity renders filled while the validator rejects the
  user's own fill override, which is a defect that looks like a rendering choice.
- **D7 — The critical interval remains a real child row** with its own `Start`,
  `Finish`, `Description`, stable ID and permitted styling. It does not inherit
  dates from its parent. It projects onto the parent lane via R4.7B's
  `OverlayParent` mode and consumes no additional lane height.

## Consequences

- **The interpretation layer is gone.** Excel, PowerPoint and raster each draw
  the primitive they are handed, so the three renderers cannot drift apart on this
  entity. This is the substantive win: the rule that was unimplementable-as-written
  becomes a rule that needs no renderer cooperation.
- `SceneLine` remains in use and is unaffected — it is already the live path for
  grid lines (5 sites in `FrameBandsBuilder`), delineators, panel borders and
  splitter bands.
- The R3.16 equivalence field table's critical row must be re-transcribed: the
  fields a renderer reads change from *Style stroke + width* to *Style fill*. This
  is a cell correction, not a primitive change, and the table's keying
  (primitive kind, owner, lifetime) is unaffected.
- The visual changes for users: a filled 4pt red bar **centred on its own slot**,
  instead of a 2.25pt line drawn along the parent's top edge. With the 8pt default
  activity height that is exactly the specified half.
- **The golden scene DID change**, contrary to the original prediction here:
  removing the parent clip resurrected a previously-deleted primitive, so the
  snapshot went from one critical primitive to two and was regenerated in
  `806a632`. R4.7E adds a critical child to the **vertical integration fixture** so
  the representation is pinned in a test that asserts on it directly, rather than
  only through a snapshot diff.
- `CriticalLinePt` disappears from the materialised metrics table on the
  VeryHidden configuration sheet, so the schema bump must carry the config
  integrity and repair path with it.

## Alternatives considered

- **Emit `SceneLine` (REV5/REV6's proposal).** Rejected by the product owner on
  size-manageability grounds. Recorded here with its real cost, because it is the
  more architecturally pure option: a line carries `From`/`To` endpoints and
  needs a thickness path, a line-cap rule and an endpoint rule in every renderer,
  and would require a guide revision, an `EquivalenceFieldTable` re-transcription
  and a golden regeneration.
- **Keep the rect unfilled and require renderers to draw only its top edge (the
  status quo).** Rejected: this is the rule that produced the observed drift, and
  it requires three renderer special cases. It is the option most likely to
  regress silently, because a renderer that ignores it produces a plausible
  looking chart.
- **Fill the rect using the existing `CriticalStroke` token as the fill.** Rejected:
  it reuses a stroke token for a fill role, so the catalogue's own naming
  contradicts its use, and any later change to the outline colour would silently
  change the bar.
- **Add a configurable `CriticalOverlayHeightRatio` token.** Rejected as a
  user-facing setting: the half-height is fully determined by the parent bar, so a
  second authority could contradict it. The ratio is derived, not configured.
