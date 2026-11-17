# ADR-0049 — Delineator line width resolves from its named style

- **Status**: Accepted
- **Date**: 2026-10-28
- **Work item**: R5.7 (delineator workflow)
- **Supersedes/amends**: nothing; extends ADR-0044 (style stroke dash) and ADR-0007 D7 (catalogue token authority)

## Context

The entity guide's §24 makes the delineator's line width `DelineatorLinePt` — a **chart-wide metric**. Every delineator on the chart shared the same weight, and there was no way to give one line a different thickness. Yet the dash pattern was already per-style (`GanttStyleDashPolicy.IsUserEditable("DefaultDelineator")`), and the custom-style editor already offered a Width control (`ThemeSettingsForm` → `StandardOutlinePt`) for other entities. The two line properties were inconsistent: dash per-style, width chart-wide.

## Decision

### D1 — The Delineator is a style-key type

`EntityTypeCatalog`: `Delineator` changes `requiresStyleKey: false` → `true`. Its `StyleKey` cell is now the row's style authority, exactly as `Custom Activity` and `Custom Milestone` already are. A deleted or unchosen style falls back to the code-owned `DefaultDelineator` — the `GanttStyleResolver` `RequiresStyleKey` branch already implements this; it simply has no Delineator to apply to until this change.

`GanttUserAuthoredStyles` is *derived* from `RequiresStyleKey` types, so `DefaultDelineator` moves from the Entity styles tab to the **Custom styles** tab — where the user edits the delineator's style, matching the owner's ruling that "delimiters get edited in the custom theme settings."

### D2 — Width comes from the style; the chart metric is the fallback

`SceneBuilder` resolves the delineator's width as:

- Style's own `StandardOutlinePt` when positive (a user-authored custom style that sets its own width)
- Otherwise the chart-wide `DelineatorLinePt` metric

The built-in `DefaultDelineator` preset resolves `StandardOutlinePt` from the `DelineatorLinePt` token (0.75 default), so the built-in path renders identically to before — the metric knob stays live for built-in delineators. `DelineatorBuilder.WithWidth` no longer overwrites the style-resolved width; resolution happens in `SceneBuilder` where both the style and the metric are in scope.

### D3 — Corner label set is the floor

A Delineator's `LabelPosition` set remains `Auto`/`TopLeft`/`TopRight`/`BottomLeft`/`BottomRight`/`None` (§24). A custom style's `AllowedLabelPositions` is **intersected** with that set, so a style may narrow within the corners but never offer `Left`/`Right`/`Inside` — which the corner-label builder cannot render for a plot-spanning line.

### D4 — No new ribbon control, no new dialog

The existing Custom style ribbon dropdown keeps its `RequiresStyleKey` gate, so it stays for Custom Activity and Custom Milestone only. The Delineator's style is edited in the Theme dialog's Custom styles tab. `btnAddDelineator` (R2.8) and `btnRemove` (ADR-0008 D3) are unchanged. R2.7a D6's live destructive-command exercise is discharged by the existing `RemoveCommand` on a delineator row — cross-referenced in both files.

## Consequences

- **Schema version 20 → 21.** The type catalogue's `RequiresStyleKey` for Delineator changes, moving the catalogue hash. No migration (ADR-0029 D6); Initialise is the remedy.
- **Golden scene snapshot regenerated.** `DefaultDelineator` now flows through `TryResolveStyle` instead of `CataloguePresetStyle`, so it carries `TextColour` from the style definition. Dedicated commit with stated reason.
- **The custom-style Width control now affects delineators.** A user-authored style applied to a Delineator row sets its line weight; a style that names no width inherits the chart-wide `DelineatorLinePt` metric.

## Alternatives rejected

- **Keep width chart-wide (status quo).** Rejected: inconsistent with the dash (already per-style) and with every other entity (width from style).
- **Per-row width columns.** Rejected: duplicates the style system and contradicts ADR-0044's "a named style carries a stroke dash pattern."
- **Let the renderer pick the width.** Rejected: the guide's "renderers consume, never recalculate" rule.
