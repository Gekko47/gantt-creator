# ADR-0038 — An anchor row below the body lets the plot's BOTTOM stretch, closed by its own line

- **Status:** Accepted (design settled; **implementation pending**)
- **Date:** 2026-10-03
- **Relates to:** ADR-0037 (the top half); ADR-0031 D2 (the padding row); ADR-0036 (the insert this depends on); ADR-0026 D3
- **Decided by:** product owner, 2026-10-03, after live-host measurement.
- **Supersedes:** nothing. **Amends:** nothing. **Extends:** ADR-0037 (which solved only the top).

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

- **D1 — a new reserved anchor row, 0.5pt, sits immediately below the last body row.** The bands paint *through* it, so the shape's bottom anchor resolves to the row beneath and every body insert lands strictly above it. `ChartAnchorRowHeightPt` is a new catalogue token, default **0.5**, range 0–4.
- **D2 — the bands extend `anchorHeightPt + MajorBoundaryPt / 2` below the plot's bottom edge**, which is 0.5 + 0.5 = **1.0pt**.
- **D3 — a new closing line at that boundary**, at `ZLayer.Frame`, terminates the stacks and covers the overhang. A 1pt line centred on the boundary covers `[plotBottom + 0.5, plotBottom + 1.5]`, and the band bottom is `plotBottom + 1.0` — covered exactly, with nothing spare.
- **D4 — `chart:frame:bottom` is UNCHANGED**, at the bottom of the padding row, and `ChartPaddingRowHeightPt` reduces **6 → 5** so the total reserved height below the body is unchanged. ADR-0031 D2's contract is untouched: the chart frame still closes on the padding row.
- **D5 — two lines is the intent, not an artefact.** The closing line closes the stacks at the data boundary; the frame line closes the chart at its margin. The rejected alternative — one frame line at `plotBottom + 1` — would have made the chart's bottom margin 1pt instead of 5pt and broken ADR-0031 D2.
- **D6 — the anchor row is resolved AND verified** like the padding row: an occupied anchor row is a typed refusal, never a silent default.
- **D7 — schema 11 → 12.** A new metric token *and* a new reserved row. Same bump class as ADR-0029 requires for either alone. No migration; the remedy is Initialise.

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

**Pending.** Steps 1–7 (tokens, schema, `GanttSheetLayout`, the scene bottom counterpart and the closing line, the Office adapters, tests, and the documentation updates) are recorded in the session decision record and are not yet in the source. This ADR is written first so the design is settled and reviewable before the layout authority moves.