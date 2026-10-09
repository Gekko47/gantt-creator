# ADR-0040 — The Month period-label format is the user's explicit choice (MM or MMM)

- **Status:** Accepted (implemented 2026-10-09)
- **Date:** 2026-10-09
- **Relates to:** entity guide §6 (period header band); ADR-0039 (the superseded MMM→MM fit ladder); ADR-0035 (a label's width is bounded by available space); ADR-0014 (approved settings surfaces)
- **Decided by:** product owner, 2026-10-09.
- **Supersedes:** [ADR-0039](0039-period-label-fit-ladder.md). **Amends:** nothing in the guide's geometry — §6 already defines `PeriodLabelFormat` as a stored, closed-set setting; this ADR records that the selected form is emitted verbatim and that no automatic step-down exists.

## Context

`PeriodLabelFormat` is a stored, culture-independent setting with closed values `MM`, `MMM`, `Quarter`, and `Week` (default `MMM` on the Month scale). The config reader/writer already round-trip it, and the factory already reads it, but until R5.12 there was **no ribbon surface** to change it — the user's only month forms were whatever the scale implied. The only width rule the guide defined was suppression below `MinimumHeaderLabelWidthPt` (18pt).

The first attempt to bridge the gap was ADR-0039's automatic fit ladder: on the Month scale with `MMM`, a period label measuring wider than its interval stepped down to `MM`. The owner reviewed the implementation and rejected it. Two reasons: (1) a per-interval ladder makes the band **non-uniform** — some months at `MMM`, some at `MM` — which is not what a month header should look like; (2) the ladder is **effectively unobservable in live use**. A three-character month at the 8pt measuring seam measures narrower than the 18pt suppression floor, so the step-down window is empty: any interval wide enough to escape suppression already fits `MMM`, and anything narrower is suppressed before the ladder is consulted. The ladder was reverted (commit on the R5.12 branch); ADR-0039 is retained as history.

## Decision

- **D1 — the format is the user's explicit selection.** On the Month scale, the user chooses `MM` (`01`) or `MMM` (`Jan`) from a second dropdown in the Plot Time Scale ribbon group (`ddnPeriodLabelFormat`). The chosen form is emitted **verbatim** by every visible interval, so the band is uniform. The dropdown is enabled only on the Month scale; on Quarter/Week the format is the scale's canonical value and the control greys out.
- **D2 — no automatic step-down.** The selected form is used as-is. An interval at or above `MinimumHeaderLabelWidthPt` emits the selected form; below it the existing suppression rule removes the label and (when all are removed) the `AllPeriodLabelsSuppressed` warning fires. There is no width-driven text substitution of any kind. A selected `MMM` that does not fit clips with the existing engine behaviour — the host's job, exactly as before R5.12.
- **D3 — the persist path writes the user's choice.** `RibbonStateService.PersistPlotSettings` currently re-derives `PeriodLabelFormat` from the `TimeScale` on every persist, which would clobber a user-chosen `MM` the moment any other ribbon setting changes. It now writes the state's `PeriodLabelFormat` member. On a scale switch, an incompatible stored format is **normalised** to the new scale's canonical format (`Month` → `MMM`, `Quarter` → `Quarter`, `Week` → `Week`) so the stored pair is always valid; the user's `MM` is therefore lost on a Quarter round-trip, matching the existing one-canonical-format-per-scale derivation.
- **D4 — no Core, catalogue, schema, or format-enum change.** `PeriodLabelFormat`, its closed-set parser, `IsCompatible`, and the factory read path already exist. The row is purely a ribbon surface plus the persist fix. No workbook migration, no `Initialise` remedy, no settings-hash drift.

## Consequences

- The Month scale now exposes both of its permitted formats; `MM` is reachable without any code change to the scene model.
- Because the format is uniform, the label primitive's role-derived identity (`{parent}:label`, R3.17) and the R4.7/R4.8 reconciliation keys are unaffected — they never were, but now the text is uniform across the band too.
- Live, export, and raster renderers all consume the same Core scene, so the verbatim emission is identical in all three; no renderer-equivalence divergence.
- The 8pt measuring seam vs the 9pt live period header is now irrelevant to label text — the format is chosen, not measured. (It still matters for the title ellipsis, unchanged.)

## Alternatives considered

- **The automatic MMM→MM fit ladder** (ADR-0039, rejected — non-uniform per-interval output, and effectively unobservable live because the measuring seam sits below the suppression floor).
- **An all-four-formats dropdown** (rejected — `Quarter`/`Week` on a Month scale are invalid pairs the config reader refuses; the dropdown must offer only the current scale's permitted formats).
- **Per-scale format memory** (rejected — one stored setting with scale-switch normalisation is the existing contract; remembering MM per scale adds state and a second source of truth for no capability the roadmap asked for).
- **Keep `MMM` and let a too-narrow band clip** (rejected as the only option — the owner wants a uniform, legible month header, which `MM` provides when `MMM` crowds).

## Implementation status

**Accepted; implemented on the R5.12 branch (2026-10-09).** `ddnPeriodLabelFormat` in `grpPlotTimeScale`; `RibbonControlIds.PeriodLabelFormat`; a `RibbonState.PeriodLabelFormat` member with a truth-table row (workbook + initialised + Month scale); `RibbonStateService` getter/setter, the `PersistPlotSettings` fix, and scale-switch normalisation; `GanttRibbon` `GetPeriodLabelFormat`/`OnPeriodLabelFormatChange`. Guide §6 records the verbatim-emission ruling (revision 11). No Core change. Observed suites green at commit time.
