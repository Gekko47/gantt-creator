# ADR-0076 — The plot's edges are unconditional, the seam is one rule, and each header edge has one owner

- **Status:** Accepted (owner ruling 2026-11-27)
- **Related:** entity guide §3, §4, §6, §8; ADR-0060; ADR-0061 D6/D7/D9; ADR-0068 D1/D11; R11.1

## Context

Three separate duplications had accumulated around the plot's boundary:

1. The plot's left and right edges were emitted inside the `ShowMajorGrid` gate, while its top and
   bottom were not — so a workbook with major grid off had an open-sided plot (ADR-0061 D6 says
   the border is a complete rectangle on all four sides).
2. In a table-plus-chart composition the panel drew its own right-hand vertical border at the same
   x as `chart:grid:edge:left`, putting two coincident primitives on the table/chart seam.
3. Every year and period cell was boxed individually, so each shared vertical edge was emitted
   twice, every period cell's top duplicated the year row's bottom, and every period cell's bottom
   duplicated `chart:plot:top`.

## Decision

- **D1 — The plot's four sides are unconditional.** `AddGrid` always emits both edge lines. The
  `ShowMajorGrid` setting that gated them is retired with the grid (ADR-0075).
- **D2 — One vertical rule at the seam.** `PanelBuilder` emits one vertical border per column's
  LEFT edge and no outer right border: a panel's right edge IS the plot's left edge, and the
  chart's edge line closes it. This is the same rule the builder already applies to a hidden
  row's coincident horizontal edge.
- **D3 — Horizontal panel rules still abut the plot's bottom.** The panel's bottom border spans the
  table and `chart:plot:bottom` spans the plot; they are end to end, not coincident, so there is
  nothing to remove there. (An earlier review claimed they overlapped; they do not.)
- **D4 — Each header edge has one owner.** A cell always draws its LEFT edge; only the last cell in
  a row draws a RIGHT edge. The year row owns its TOP and BOTTOM (the bottom being the period row's
  top); the period row owns NO horizontal edge, because `chart:plot:top` is its bottom.
- **D5 — No live inset for header cell edges.** The frame and the plot's edges are inset by half a
  stroke on the live sheet because a measured loss of half the stroke to a cell fill or the sheet
  edge made them disappear. The header boxes are shapes over a filled band, nothing clips them, and
  the live header row IS the table's header row, so its own cell borders already draw on the same
  x. Insetting the shapes would create two near-parallel hairlines where they coincide today. They
  stay coincident, deliberately.

## Consequences

- A header row of N cells emits N+1 vertical lines and no coincident pair; the period row emits no
  horizontal line at all.
- The scene golden loses twelve header lines and the seam loses one panel line (recorded in the
  R11.1 golden commit).
- §8's "no duplicate coincident minor/major lines" obligation is now true for the header, and its
  "year line wins at shared X" clause is retired with the grid.
