# ADR-0046 — Every chart border is one uniform soft grey

- **Status**: Proposed (owner approved direction 2026-10-20)
- **Date**: 2026-10-20
- **Work item**: R5.6a (named style presets) / R5.6 (style and theme settings)
- **Supersedes/amends**: amends the entity guide's §6 colour-token table (`MajorGridStroke`); amends ADR-0045 D2's token-authority rule only in the value the authority carries

## Context

ADR-0045 made the chart's frame colours editable settings whose defaults are the entity guide's approved tokens. That is the right authority, and it removed the factory's own literals — but it also changed what the live chart draws, because the token and the old literal disagreed.

| Frame member | ADR-0045's factory literal (what the user saw) | Guide token value (what ADR-0045 wired in) |
|---|---|---|
| Minor grid | `#E0E0E0` | `MinorGridStroke` = `#D9D9D9` |
| Major grid / year boundary / **plot frame** | `#BFBFBF` | `MajorGridStroke` = `#000000` |

The plot frame's four edges, the year boundaries and the plot's closing line are all drawn from `Theme.MajorGrid` (`FrameBandsBuilder.AddOuterFrame` and the grid builder), so the single token change turned every one of them **black**. The chart's borders stopped being the uniform soft grey they had been, and the frame no longer matched the worksheet's own table borders.

The owner ruling (2026-10-20): **the chart's borders must be one uniform soft grey.**

## Decision

### D1 — `MajorGridStroke` becomes the minor grid's soft grey

`MajorGridStroke` changes from `#000000` to `#D9D9D9`, the value `MinorGridStroke` already carries. One colour, one authority, no second grey to reconcile.

The **year boundary keeps its own `MajorBoundaryPt` width** (1.0pt against the minor grid's 0.5pt), so it still reads as a boundary. The hierarchy is carried by weight rather than by a second colour — the same way a good chart does it, and the reason the earlier black value was never needed.

### D2 — The guide's token table is amended, not worked around

The entity guide's §6 colour-token table is the contract for what the chart looks like, so the row changes with the value. A code change that quietly disagreed with the guide would be the defect ADR-0045 was written to remove.

### D3 — The token remains the single authority for the default

Unchanged from ADR-0045 D2: each chrome setting's default **is** its token value, read from `GanttCatalogues.Colours`. A test pins that the two grid roles default to the **same** colour, so a future edit cannot reintroduce two greys.

### D4 — A workbook that stored the old value keeps it, by design

The chrome colours are **settings**, and a stored setting is the user's choice. A workbook initialised while `MajorGridStroke` was black stores `MajorGridColour = #000000`, and the stored value wins over the default. The remedy is the Chart tab's **Reset to default**, which restores the code-owned presets. There is no silent migration of a user-editable setting (ADR-0029 D6).

## Consequences

- A new workbook, or one whose chart colours are reset, draws every border in `#D9D9D9`.
- The catalogue hash moves, so an existing workbook reports a catalogue mismatch and the remedy is Initialise (ADR-0029 D6, no migration) — and then, for the border colour itself, Reset to default on the Chart tab.
- The export composition's panel border, which reads the same token, follows the same colour.

## Alternatives considered

- **Leave the token black and special-case the live frame.** Rejected: that keeps two authorities for the same colour, which is the defect ADR-0045 exists to remove.
- **Give the year boundary its own colour token.** Rejected: a second grey is the opposite of the ruling, and the width already carries the hierarchy.
- **Migrate the stored setting silently on read.** Rejected: a user-editable setting is never replaced behind the user's back; the guide's colour rule forbids it.
