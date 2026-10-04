# ADR-0038 — An anchor row below the body lets the plot's BOTTOM stretch, closed by its own line

- **Status:** Accepted (implemented 2026-10-03)
- **Date:** 2026-10-03
- **Relates to:** ADR-0037 (the top half); ADR-0031 D2 (the padding row); ADR-0036 (the insert this depends on); ADR-0026 D3
- **Decided by:** product owner, 2026-10-03, after live-host measurement.
- **Supersedes:** nothing. **Amends:** [ADR-0037](0037-plot-bands-anchored-into-the-header-row.md) D4 (the top edge is not the only one that moves). **Extends:** ADR-0037, which solved only the top.

## Context

ADR-0037 lifted the plot-spanning shapes into the header row so they **stretch** when a row is added at the top. The bottom did not follow. A live report confirmed it: the lane bars move, but the bands and grid lines do not extend to the footer.

A probe of the bottom case (`scripts/probe-frame-bottom-anchor.ps1`) first appeared to show the append path was already correct — stretching with zero extension. Reality contradicted it, which is what exposed the real mechanism below.

## Evidence

`scripts/probe-anchor-row-height.ps1`, Excel 16.0 x64, 2026-10-03.

**Q1 — the row-height clamp.** Requested heights, read back:

| Requested | Read back | |
| --- | --- | --- |
| 0.25 | 0.25 | honoured |
| **0.5** | **0.5** | **honoured** |
| 0.6 | 0.5 | quantised down |
| 0.75 / 1 / 1.5 / 2 | exact | honoured |

There is **no 0.75pt floor**. The 0.5pt anchor row is achievable, so the token default is 0.5 and no fallback is needed.

**Q2 — the layout stretches.** Body → anchor(0.5pt) → pad(5pt), inserting at the anchor row:

| Insert at | Anchor | Height delta | Result |
| --- | --- | --- | --- |
| **row 9 — the anchor, the append target** | 0.5pt | **+18** | **STRETCHES** |
| row 9 | 1pt | +18 | STRETCHES |
| row 5 — body, control | 0.5pt | +18 | STRETCHES |

## The finding this ADR exists to preserve

**Insert at a shape's `TopLeftCell` row SLIDES it (`TopDelta` +18, `HeightDelta` 0). Insert at its `BottomRightCell` row STRETCHES it (`HeightDelta` +18).** The two ends are *not* symmetric.

This is why ADR-0037's `LiftTopIntoHeader` cannot simply be mirrored for the bottom, and why a well-meaning "symmetrise the two edges" change would re-break the bottom. It is also why the bottom needs a *reserved row* rather than a lifted edge: the append targets one row past the **body**, so the shape's bottom anchor must resolve to a row **below** that point, and the anchor row is what creates one.

## Decision

- **D1 — a new reserved anchor row, `ChartAnchorRowHeightPt` (0.25pt), sits immediately below the last body row.** The bands paint *through* it, so the shape's bottom anchor resolves to the row beneath and every body insert lands strictly above it. The token's range is 0–4, and `0` reproduces the pre-ADR-0038 geometry exactly.

  **Amended 2026-10-03 by owner ruling: the default is 0.25pt, not 0.5pt**, and `ChartPaddingRowHeightPt` reduces **6 → 5.75** so the reserved strip below the body totals the same **6pt** it reserved as one row. The probe measured 0.25pt honoured exactly, so no fallback is needed; see "Amendments" below for why the pairing matters.
- **D2 — the bands extend `anchorHeightPt + MajorBoundaryPt / 2` below the plot's bottom edge**, which at the amended values is 0.25 + 0.5 = **0.75pt**. **This reverses ADR-0037 D4**, which held that "the top edge is the one that moves, never the bottom". D4 was right about the *method* — a sub-row overlap cannot create the row an edge anchors into — and wrong about the bottom being immovable; the reserved row is what creates one.
- **D3 — a new closing line at that boundary**, at `ZLayer.Frame`, terminates the stacks and covers the overhang. A 1pt line centred on the boundary covers `[plotBottom + 0.25, plotBottom + 1.25]`, and the band bottom is `plotBottom + 0.75` — covered, with nothing spare.

  **Implemented as a package with the anchor row**: a zero anchor row emits no extension *and* no closing line. Half a line would otherwise survive on its own, making "zero restores the previous geometry exactly" false by 0.5pt and drawing a rule across the plot's bottom edge with no band behind it. `PlotSpanGeometry.HasAnchorRow` is that decision, and it is why the counterweight test can assert exact prior geometry.
- **D4 — `chart:frame:bottom` is UNCHANGED**, at the bottom of the padding row, and `ChartPaddingRowHeightPt` is **5.75** so the total reserved height below the body is unchanged at 6pt. ADR-0031 D2's contract is untouched: the chart frame still closes on the padding row, and the bottom margin is now the **sum** of both reserved rows.
- **D5 — two lines is the intent, not an artefact.** The closing line closes the stacks at the data boundary; the frame line closes the chart at its margin. The rejected alternative — one frame line at `plotBottom + 0.75` — would have made the chart's bottom margin 0.75pt instead of 5.75pt and broken ADR-0031 D2.
- **D6 — the anchor row is resolved AND verified** like the padding row: an occupied anchor row is a typed refusal, never a silent default. It carries **its own** refusal enum (`AnchorRowRefusalReason`) and its own adapter refusal (`AnchorRowNotOwned`), so a diagnostic can name which row the user has to fix.
- **D7 — schema 11 → 12.** A new metric token *and* a new reserved row. Same bump class as ADR-0029 requires for either alone. No migration; the remedy is Initialise.

## Amendments

**2026-10-03, owner ruling (implementation review):** D1's default and D4's padding value are paired — **anchor 0.25pt, padding 5.75pt, total 6pt**. The ADR as first written said 0.5 and 5, which also "unchanged" only approximately: 0.5 + 5 = 5.5, a 0.5pt shrink of the chart's bottom margin. The owner chose the pairing that makes D4's rationale *literally* true — the reserved strip below the body is exactly as tall as the single row it replaces — and the probe had already measured 0.25pt honoured exactly, so nothing is lost by preferring it. D2's arithmetic moves with it (0.75pt, not 1.0pt) and D3's coverage window with that; both remain exactly closed.

A consequence worth recording: this makes `The_padding_row_default_matches_the_chrome_margin_it_replaced` false as originally written. It is **rewritten, not deleted**, as `The_reserved_below_the_body_strip_totals_the_chrome_margin_it_replaced`, asserting the **sum**. The owner's earlier correction of an 18pt default to 6pt was about the size of this margin, so dropping the assertion rather than moving it would have lost the invariant it protected.

## Consequences

- **Bands, vertical grid lines and delineators are fixed together.** They share the plot's vertical span and fail as one class; leaving delineators out would leave a full-height line that stops at the old footer.
- **The chart background is untouched**, for ADR-0037 D2's reason: it spans the whole chart and its top and bottom already sit outside the body.
- **The inserter is unchanged.** It already targets one row past the **body**, which is now the anchor row — so this decision makes the existing insert correct for the bottom rather than changing it.
- **The overhang must be asserted, not eyeballed.** A test computes that the band bottom never passes the closing line. The probe's own `covered=False` output is an **artefact**, not a finding: it measured line coverage with `Top + Height`, and an Excel `AddLine` has a degenerate bounding box. Coverage comes from `Line.Weight`, so the test must do the arithmetic.
- **Cost:** one more reserved row per sheet, one more token, one more primitive family (the closing line), and a schema bump that makes existing workbooks require Initialise.

## Alternatives considered

- **Mirror ADR-0037's lift at the bottom** (rejected: measured — an insert at the bottom anchor *stretches*, so the top's failure mode does not apply there; and a lifted edge has no anchor row to resolve into).
- **Extend the bands down past the padding row to grab an anchor** (rejected: paints band colour across the whole bottom margin).
- **Reuse `chart:frame:bottom` as the closing line** (rejected: that would move the chart's bottom margin to 1pt and break ADR-0031 D2 — see D5).
- **A whole-row anchor (18pt) instead of 0.5pt** (rejected: 0.5pt is measured to work and a full row would be a visible strip of the sheet).

## Implementation status

**Implemented 2026-10-03.** Tokens, schema 12, `GanttSheetLayout` (anchor row plus the padding row moved down one), the scene's bottom counterpart to `PlotSpanGeometry.LiftTopIntoHeader`, the `chart:plot-closing` line, both Office adapters, the tests, and the documentation updates are all in source. The golden scene snapshot was regenerated and verified per primitive: exactly `chart:band:1`, `chart:band:3`, five `chart:grid:*` lines, two delineator lines, and the new `chart:plot-closing` changed — every band delta being `400 → 400.75`. `ChartBounds`, every lane, bar, label, and milestone are byte-identical.

The two amendments above (anchor 0.25 / padding 5.75, and the closing line being a package with the anchor row) were found during implementation rather than before it, and both are recorded above because both change what the ADR claims.