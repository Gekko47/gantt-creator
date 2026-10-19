# ADR-0044 — A named style carries a stroke dash pattern

- **Status**: Proposed (owner approved direction 2026-10-19)
- **Date**: 2026-10-19
- **Work item**: R5.6a (named style presets)
- **Supersedes/amends**: nothing; extends ADR-0009 (style capability schema) and ADR-0007 D7 (catalogue token authority)

## Context

The entity guide's §16 makes a solid line the default and lets an approved named style specify a dash pattern, and §22 lists "dash pattern" among the properties a named style may define. The **workbook schema carried no column for it** and **no layer modelled it**: `tblGanttStyles` had no dash column, `GanttStylePreset` had no dash member, `SceneStyle` had no dash member, and `OfficeShapeStyle` had no dash member. The host, however, has supported it all along — `Excel.LineFormat.DashStyle` is a get/set property of type `Microsoft.Office.Core.MsoLineDashStyle` (verified on the installed interop assembly).

The same gap applied to the **fill type** in the editor: `tblGanttStyles` already carried `HatchPattern` and the renderer already honoured it, but the custom-style editor offered no control for it, and the writer hard-coded `None` for every user row.

## Decision

### D1 — A named style carries a stroke dash pattern, stored in a new appended column

`tblGanttStyles` gains a `StrokeDash` column **appended at index 14**, so every existing column index is unchanged and the change is a pure append. The column holds the **name** of a `GanttStrokeDash` member, never a number.

The append was chosen over inserting the column among the other style properties because an insert would renumber every later column and turn the change into a rewrite of every stored row rather than an addition.

### D2 — The dash vocabulary is a closed, curated subset of the host's

`GanttStrokeDash` has six members: `Solid`, `SquareDot`, `RoundDot`, `Dash`, `DashDot`, `LongDash`.

The host's `MsoLineDashStyle` has twelve. The six omitted are the `SysDash` family (`msoLineSysDash`, `msoLineSysDot`, `msoLineSysDashDot`) and the two long-dash-dot variants (`msoLineLongDashDot`, `msoLineLongDashDotDot`). The `SysDash` family renders against the operating system's UI theme rather than a fixed geometry, so a workbook carrying one would look different on another machine — the same reason the PNG export refuses to substitute a colour. The two long-dash-dot variants are omitted for the same reason the pattern set is: the picker offers the six a line-style control in mainstream product actually offers.

The mapping is **closed**: `ExcelShapeWriter.MapDashStyle` throws `ArgumentOutOfRangeException` for an unmapped member, exactly as `MapHatchPattern` does. A dash the scene resolved is written to the host **verbatim**; the renderer never chooses one.

### D3 — The dash is a resolved scene token, not a renderer choice

The dash travels `GanttStyleDefinition.StrokeDash` → `SceneStyle.StrokeDash` → `OfficeShapeRequest.StrokeDash` → `OfficeShapeStyle.StrokeDash` → `line.DashStyle`.

The scene snapshot gains the field and its version advances 2 → 3. The rule is the guide's: a renderer consumes the scene and never recalculates it, so a dash the scene drops can never be drawn.

**Solid is written explicitly**, not left alone: the host keeps whichever dash it last had, so a style that resolved to `Solid` must say so rather than inherit a previous dash.

### D4 — A style with no visible stroke has no dash

`OfficeStyleMapper` reports `Solid` for a shape whose stroke is not visible. A dash the host would never draw is not carried.

### D5 — A partial style edit preserves the stored dash

A null `StrokeDash` on a style edit means **not carried**, and the writer preserves the stored value — the same rule the fill pattern and the sizes follow. An edit that changes one colour must not reset the style's stroke type.

### D6 — Built-in styles keep a code-owned dash

Every built-in preset carries `Solid`. The reader exact-matches the built-in rows' dash column against the preset, so a built-in row's dash cannot drift from the code-owned value. Built-in dashes are **not** user-editable: making them editable would break that exact match for no user gain, and the built-in patterns are already code-owned.

## Consequences

- **Schema version 18 → 19.** A version-18 workbook's style table lacks the column and its stored catalogue hash no longer matches. There is **no migration** (ADR-0029 D6): the remedy is Initialise.
- **The catalogue hash moves**, because the style rows' contents move and the hash's first line seeds with the schema version.
- **The golden scene snapshot is regenerated** to version 3, with the dash field added. The regeneration is a dedicated commit with a stated reason, per the work item's change policy.
- **No export-parity surface is broken.** The PNG and PowerPoint renderers do not exist yet (`GanttCreator.Raster` holds only the export-width parser), so there is no rasteriser that could silently drop the dash. The deviation is recorded for the day they do.
- **The fill type needs no schema change.** The pattern column already existed; only the editor and the write path were missing it.

## Alternatives rejected

- **Insert the dash column among the other style properties.** Rejected: it renumbers every later column and turns an append into a rewrite of every stored row.
- **Expose the host's full twelve-member dash set.** Rejected: the `SysDash` family renders against the OS UI theme, so a workbook would not look the same everywhere.
- **Let the renderer pick a dash when the style names none.** Rejected: the guide's "renderers consume, never recalculate" rule; a renderer that chose would draw a line the scene never described.
- **A separate `StyleType` column binding a style to one entity kind.** Rejected: the owner confirmed a style is not bound to a kind — the same custom style serves a Custom Activity row and a Custom Milestone row, and the row's type decides the geometry.
