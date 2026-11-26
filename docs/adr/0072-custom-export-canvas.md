# ADR-0072 — Custom width: a workbook Size mode whose width is the live width budget

- **Status:** Accepted, revised 2026-11-26
- **Date:** 2026-11-26
- **Relates to:** ADR-0020 (application-state scope), ADR-0030 D1/D2/D7 (row-driven
  vertical extent), ADR-0031 D2 (flush-left plot), ADR-0065 D4 (zero visible
  columns), ADR-0073 (the schema-29 keys)
- **Implements:** R8.9
- **Decided by:** product owner, 2026-11-26

## Context

The 2026-11-23 ruling withdrew R8.7's per-export width controls: "The PNG is the chart
the user sees, so there is no width or unit input." The owner then asked for a
**Custom** Size choice in the ribbon, entered in centimetres, overriding the preset
page and margin maths.

Those two requests read as a contradiction unless their scopes are separated. The
2026-11-23 ruling governs the **per-export** path: a PNG is sized from its own
composition, and no command prompts for a separate width. It does not forbid a
**workbook-level** Size choice that the plot and every export consume.

Only preset width is consumed today: `plotWidth = preset.WidthPt − panelRight −
leftChrome − rightChrome`, and the vertical extent is the measured rows
(`boundVerticallyToPage: false`).

The owner also ruled, after the first implementation, that a custom **height** cannot
be offered: "it is not possible to back calculate width without forcing a fixed ratio."
A typed height can only be honoured by resizing the user's rows, scaling the
composition (which changes the type), or deriving a width from a ratio that stops
being uniform once that width is also the live budget. None of those is a size
control, so the height input was withdrawn and the export canvas with it.

## Decision

- **D1 — Custom is a workbook Size mode, not a per-export input.** One persisted
  setting (`SizePreset = Custom` plus `SizeWidthCm`) is read by the live plot and by
  every export. The chart's-own-size path stays the default and is byte-identical.
- **D2 — There is one custom dimension: the width.** It is the **whole composition
  width** — the data table plus the plot plus the frame chrome. The table keeps its
  measured column widths and the plot takes the remainder, through the same
  `PlotGeometryResolver` subtraction the presets use, with `marginPt = 0`.
- **D3 — There is no custom height.** The live height is the measured rows (ADR-0030).
  The stored `SizeHeightCm` row is a **mirror of the measured live plot height**,
  refreshed wherever the width is measured and written with the layout group. It is
  never an input and never a page dimension.
- **D4 — The export is driven by the live composition.** No canvas, no letterbox, no
  scaling: PNG, Copy Editable and PowerPoint read the composition's own bounds at 300
  DPI / its own envelope, exactly as they do for a preset. A wider plot is a wider
  export.
- **D5 — Page presets and margins are ignored only while Custom is active.** The margin
  controls stay visible but disabled, and their stored values survive.
- **D6 — The ribbon shows the live numbers for every mode.** The plot width and the
  measured plot height are re-measured on activation, workbook change, refresh, preset
  change, margin change, size change and column-visibility change, so a preset user
  sees the plot shrink as the measured table widens. The readouts are centimetres only;
  no point value is shown (owner ruling 2026-11-26).
- **D7 — Selecting Custom adopts the width already on screen** (the active preset's
  width, or the stored custom width), then applies the band and workable-minimum
  clamps. A 16:9 user does not jump to a fixed default.
- **D8 — The width is clamped, not refused, on read.** A stored value outside the
  1.0–200.0 cm band, or below the smallest width that leaves a readable plot, is
  corrected. Refusing at the read would leave the ribbon showing no preset while the
  workbook still says Custom, with no control able to correct it.

## Consequences

- The export model has **no custom-size branch at all**. `ExportBounds.Bounds`,
  `PlacementDelta`, the 1:1 staging envelope check and the PowerPoint expectation are
  untouched, so the withdrawal removed code rather than adding a second path.
- A custom width smaller than the measured table plus chrome plus
  `MinimumPlotWidthPt` is raised to the workable minimum by the ribbon; the factory's
  shortfall refusal remains the backstop for a hand-edited workbook.
- The `SizeHeightCm` row stays in the schema, because removing a key is not an
  append-only change and the in-place repair (ADR-0073) can only append. It carries the
  measured height instead of a dead default.
- A PNG too large at the custom width is refused by the existing 300-DPI dimension
  check; there is no separate pixel-budget clamp to maintain.

## Alternatives considered

- **A custom height defining a 1:1 export canvas (with letterbox or uniform scale).**
  Withdrawn by the owner. The ratio cannot hold while the width is also the live
  budget, and scaling changes the type the user sees.
- **Prompt for a width on every export.** Rejected: withdrawn on 2026-11-23, and it
  needs a dialog the Size control does not have.
- **Add the Custom preset to `SizePresets.All`.** Rejected: the catalogue is the four
  fixed presets and `ByKey` must keep returning only them; Custom is built at runtime.
- **A second stored height key as an input.** Rejected: nothing could honour it without
  resizing rows or scaling.
