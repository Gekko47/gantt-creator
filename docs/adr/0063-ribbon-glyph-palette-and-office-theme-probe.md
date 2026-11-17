# ADR-0063 — Ribbon glyphs use a two-palette Lucide set and a guarded Office theme probe

- **Status**: Accepted
- **Date**: 2026-11-16
- **Work item**: R5.19 (`docs/work-items/R5.19-ribbon-icon-system.md`)
- **Relates to**: ADR-0004 (package pins), ADR-0058 (ribbon tab context), R5.16a icon ruling, R5.18h column glyphs, R5.18f lock glyphs
- **Amends**: nothing. It replaces the `imageMso` policy of R5.16a D4/D5 for glyph *delivery* while keeping its one-glyph-per-control rule.

## Context

The tab mixed two icon languages. Twenty-one controls carried Office gallery `imageMso` ids, which Office recolours for its own themes, while five carried committed PNGs drawn with a fixed opaque black stroke. On a Dark Gray or Black ribbon those PNGs are effectively invisible, and the mixture reads as broken rather than styled. The owner ruling of 2026-11-14 (D7) already decided that Lucide is the icon language for every ribbon glyph, which makes the delivery mechanism — not the icon choice — the open question.

Three constraints shape the answer:

1. **No supported theme API exists for this add-in.** `Office.context.officeTheme` and `OfficeTheme.isDarkTheme` are Office.js surfaces; a COM/Excel-DNA ribbon has no equivalent. The supported way to *obtain* the theme is therefore closed to us.
2. **`imageMso` fails silently.** An unknown id renders nothing, with no error and no log, which is why the repository verifies ids against two local Office reference documents before pinning them. Any replacement mechanism must be equally checkable.
3. **The Office theme is not the Windows theme.** Only the `Use system setting` choice follows Windows, so a Windows-only probe would pick the wrong palette for a user who has said, explicitly, Office Black.

## Decision

**D1 — Two palettes, two sizes, one geometry.** Every glyph is Lucide geometry on its native 24 × 24 view box, stroke width 2, round caps and joins, rasterised at 16 px and 32 px in a `light` stroke and a `dark` stroke. Neutral values only, no accent colour, so a glyph never competes with an Office accent or a chart token.

**D2 — One callback, one generated catalogue.** A single `getImage` handler serves every static control. Glyph identity and pixel size live in a generated `RibbonGlyphCatalog` keyed by control id, emitted by the same script that emits the PNGs, so there is no second hand-maintained table and no reliance on a host channel that cannot be tested in-process.

**D3 — Resolution order, fail-safe.** High contrast first (`SystemInformation.HighContrast`, then `SystemColors.Window` luminance), then the Office UI theme, then the Windows app theme, then `light`. An unrecognised or unreadable Office value is logged once and falls through; it is never guessed.

**D4 — The Office theme value is observed, never assumed.** `HKCU\Software\Microsoft\Office\{version}\Common` → `UI Theme`, `REG_DWORD`. `{version}` is the highest Office version key that holds the value — `16.0`, then `15.0`, then `14.0` — rather than a read of `Application.Version`, so answering an image request never touches Excel's object model. Observed 2026-11-16: Colorful `7`, White `5`, Dark Gray `3`, Black `4`, Use system setting `6`. The dark set `{3, 4}` agrees with Microsoft's documented predicate that only Black and Dark Gray are dark. No other mapping may be coded, and a future value outside the set is treated as unknown.

**D5 — `getImage`, not `loadImage`.** A theme change can happen mid-session. `IRibbonUI.Invalidate()` re-queries `getImage`; `loadImage` is a load-time channel and is not guaranteed to reload.

**D6 — Repaint on change.** A coalescing timer re-reads the one registry value, compares, and invalidates the ribbon on the Excel thread only when it changed. No `Thread.Sleep`, no blocking wait, disposed with the add-in lifecycle.

**D7 — The chart stays theme-independent.** The palette is ribbon chrome. It is never written to `_GanttCreatorConfig`, never projected into the scene, and does not weaken R5.6 D1, which keeps chart tokens explicit and free of Office-theme inheritance.

## Consequences

- A user on Dark Gray or Black now sees glyphs rather than ghosts, and the icon language on the tab is one visual family.
- The add-in depends on an undocumented registry value. That dependency is one isolated class, guarded by a fallback chain, and recorded here with the evidence that produced the mapping. If a future Office build changes the value, the palette degrades to the Windows and high-contrast branches, which is a legible wrong answer rather than a silent one.
- High-contrast users keep legible glyphs, but through our own branch rather than Office's recolouring of `imageMso`. That branch is the first thing evaluated for exactly this reason.
- Roughly 62 PNG assets are generated rather than drawn by hand, and a script proves each one exists in both palettes at its declared size with an approved stroke.
- `getImage` is queried per control per invalidation. On this tab that is 44 controls where 2 existed before, so the invalidation path and the COM leak signal are measured before the change is accepted rather than assumed acceptable.
- `item` elements cannot participate: `CT_Item` has no `getImage`, so the four Size preset rows carry no glyph. The owner then ruled that the Size control itself should carry none either — it is a text dropdown beside an inline label, and a picture there only crowds the row — so the family gains no Size glyph at all.

## Alternatives considered

- **Keep `imageMso` and draw only the gaps.** Rejected: it preserves Office's theme adaptation for free but leaves two icon languages on one tab, which is the defect being fixed, and contradicts the owner's D7 ruling.
- **One neutral mid-grey palette, no probe.** Rejected as the primary path, though retained as the documented outcome if the probe is ever rejected: a mid-grey that clears 3:1 on both backgrounds is safe but visibly weak on white, and it cannot express the user's explicit Office Black choice.
- **`loadImage` with static `image` attributes.** Rejected: fewer callbacks, but a theme change mid-session would leave stale pictures until the ribbon reloads.
- **Glyph identity in the RibbonX `tag` attribute.** Rejected: it removes a table but depends on the host surfacing `Tag` on the control handed to `getImage`; if it does not, every glyph disappears silently — the exact failure class this work exists to remove.
- **Pixel-sampling the live ribbon window.** Rejected: it breaks on collapsed ribbons, DPI scaling and animations, and it would sample a repaint in progress.
- **UI Automation to read the ribbon background.** Rejected: no background colour is exposed.
- **Reading the Windows theme only.** Rejected: it is wrong for a user whose Office theme differs from Windows, which is a normal configuration.
