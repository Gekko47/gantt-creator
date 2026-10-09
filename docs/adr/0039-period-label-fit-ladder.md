# ADR-0039 — The month period label steps down MMM → MM when the band is too narrow

- **Status:** Superseded by [ADR-0040](0040-period-label-format-selection.md) 2026-10-09
- **Date:** 2026-10-09
- **Relates to:** entity guide §6 (period header band); ADR-0035 (a label's width is bounded by available space); ADR-0015 (the measured-fallback precedent for event labels); ADR-0033 (label appearance)
- **Decided by:** product owner, 2026-10-09.
- **Supersedes:** nothing. **Amends:** [07-GANTT-ENTITY-GUIDE.md](../07-GANTT-ENTITY-GUIDE.md) §6, which previously emitted the selected format verbatim for every interval at or above `MinimumHeaderLabelWidthPt`.

## Context

The period header band emits one label per visible interval in the user's selected `PeriodLabelFormat` (default `MMM` for the Month scale). The only width rule the guide defined was suppression: an interval narrower than `MinimumHeaderLabelWidthPt` (18pt) emits no label at all. Between that floor and a comfortably wide interval there is a middle band — visible, but too narrow for the three-character month name — where the label either clipped (the host cuts glyphs because header text boxes do not auto-size or wrap) or had to be suppressed. Neither is what a month header should do: a clipped "Jan" reads as a defect, and a missing month breaks the scale.

The format catalogue already defines a narrower month form — `MM`, the two-digit month — as the Month scale's other permitted format. The question was whether the renderer may choose it when the selected form does not fit.

## Decision

- **D1 — the fit ladder.** On the `Month` scale with the `MMM` format, a period label that measures wider than its interval's visible width is emitted in the two-digit `MM` form instead ("Jan" → "01"). The measurement uses the same injected `ITextWidthMeasurer` seam the title ellipsis uses. The comparison is strict: a label that measures exactly its interval's width fits.
- **D2 — the step is unconditional.** When even `MM` measures wider than the interval, `MM` is still emitted and clipped by the existing host behaviour (`AutoSize = msoAutoSizeNone`, wrapping off). The owner's ruling: the narrowest defined form is always used; clipping is the existing engine's job, not a reason to suppress or to keep the wider form. Suppression remains reserved for intervals below `MinimumHeaderLabelWidthPt` and is evaluated first — the ladder never resurrects a suppressed label.
- **D3 — the scope is Month + MMM only.** `Quarter` (`Q1`), `Week` (`W01`) and the `MM` format have no defined shorter form and are emitted verbatim. The year band's four-digit year is unaffected. Any abbreviation form for another scale is its own decision.
- **D4 — the ladder lives in the presentation seam.** `FrameBandsBuilder` resolves the emitted text (`ResolvePeriodLabel`), mirroring the title ellipsis. `BandSequence` stays pure date/geometry/formatting: its `Label` remains the selected-format label, and its suppression rule is untouched.
- **D5 — no new warning.** The step-down is the documented policy, not an anomaly, so it is silent. The existing `AllPeriodLabelsSuppressed` warning still fires when suppression removes every period label. This contrasts with the title truncation, which warns because it is unexpected.

## Fidelity note

The measuring seam is 8pt (`LabelBodyFontSizePt`) while the Office factory's `PeriodHeader` style renders at 9pt — the same approximation the title ellipsis already lives with (an 8pt seam against an 11pt title token). The ladder can therefore fail to fire on a borderline interval, keeping an `MMM` that clips at the rendered size. Per D2 that is cosmetic: the host clips, and the chart remains correct. Exactness would require a period-header font-size input to `FrameBandsRequest`; deferred, because it touches the request contract and the factory for a borderline-only improvement.

## Consequences

- Label primitive identities are role-derived (`{parent}:label`, R3.17), so the text change cannot drift the R4.7/R4.8 shape-reconciliation keys — a renderer matching on `chart:period:{yyyy-MM-dd}:label` is unaffected by which text the label carries.
- Live, export and raster renderers all consume the same Core scene, so the ladder applies identically in all three; no renderer-equivalence divergence.
- No catalogue, settings, schema or format-enum change: no workbook migration, no `Initialise` remedy, no settings-hash drift.
- The visible effect is bounded: only Month-scale charts whose months are narrower than the measured `MMM` but at least `MinimumHeaderLabelWidthPt`.

## Alternatives considered

- **Keep `MMM` and let it clip** (rejected — the owner's ruling: the narrowest defined form is always used; a clipped month name reads as a defect).
- **Suppress when neither form fits** (rejected — suppression is reserved for sub-minimum intervals; the ruling prefers a clipped `MM` to a missing month).
- **A new user setting for the fallback** (rejected — `PeriodLabelFormat` stays the user's choice; the ladder is rendering policy, not a preference).
- **Locale-driven abbreviations** (rejected — all label formatting is culture-invariant, by the guide's own rule).

## Implementation status

**Superseded 2026-10-09 — never shipped.** This ADR was implemented on the R5.12 branch (`FrameBandsBuilder.ResolvePeriodLabel`, the measurer threaded into `AddHeaders`, and the Core tests) and reverted before merge. The owner's follow-up ruling rejected the automatic step-down: a per-interval ladder makes the band non-uniform (some months at `MMM`, some at `MM`) and is not reliably observable in live use — a three-character month at the 8pt measuring seam is narrower than the 18pt `MinimumHeaderLabelWidthPt` suppression floor, so the step-down window is effectively empty and any interval wide enough to escape suppression already fits `MMM`. The selected `PeriodLabelFormat` is instead surfaced to the user directly as a ribbon dropdown (Month scale: `MM` or `MMM`) so the format is a uniform, explicit choice. See [ADR-0040](0040-period-label-format-selection.md). The decision record and its reasoning are retained here as history.
