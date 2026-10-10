# ADR-0041 — Period-label visibility is measured-fit, so MM shows where MMM does not

- **Status:** Accepted (implemented 2026-10-10)
- **Date:** 2026-10-10
- **Relates to:** entity guide §6 (period header band); ADR-0040 (the explicit-format ruling, amended not superseded); ADR-0039 (the reverted MMM→MM ladder, retained as history); ADR-0035 (a label's width is bounded by available space)
- **Decided by:** product owner, 2026-10-10.
- **Amends:** [ADR-0040](0040-period-label-format-selection.md) D2. The format is still the user's explicit, uniform, verbatim choice; what changes is the **visibility rule** for the period band. **Supersedes nothing** — ADR-0039's ladder (per-interval text substitution) remains rejected; this is measured-**fit**, not a ladder.

## Context

After ADR-0040 the owner could choose `MM` (`01`) or `MMM` (`Jan`), but observed that both forms appeared to occupy the same width and `MM` did not reveal more months on a cramped plot. The cause: period-label visibility was gated by the single scalar `MinimumHeaderLabelWidthPt` (18pt) floor, identical for both forms, so switching to `MM` changed the text but not which intervals showed a label. `MM` was, functionally, a cosmetic choice.

The owner's ruling (2026-10-10): the step must be **uniform** (no per-interval mix of `MM` and `MMM` — the ladder's flaw), and it is cleaner for the **user to decide** the form than for the engine to step it down automatically. Given both, option **B** was chosen: replace the fixed period floor with a **measured-fit** test — a period label is shown iff its interval is at least as wide as the label **measures**. Because `MM` measures shorter than `MMM`, it survives on narrower intervals, so on a cramped plot `MM` renders month headers where `MMM` suppresses them — the capability the owner wanted, delivered without any per-interval form change.

## Decision

- **D1 — period visibility is measured-fit; the year band keeps the floor.** In `BandSequence`, a **period** interval shows its label iff `intervalWidth >= measuredLabelWidth`; a **year** interval keeps `intervalWidth >= MinimumHeaderLabelWidthPt`, because a year interval always spans the whole plot and its four-digit label is never the constraint. The measured-fit decision stays in `BandSequence.AddInterval` (the single show/hide authority), so the existing `AllPeriodLabelsSuppressed` warning — which reads `period.ShowLabel` — is unchanged.
- **D2 — one seam, no second font-size authority.** The measurement uses the **same injected `ITextWidthMeasurer`** the chart title already uses (8pt). It does **not** introduce a period-header-sized measurer. The period header renders at 9pt with no pinned family — the identical approximation the `Title` style (12pt, no family) already ships with — so B inherits that residual rather than adding a new one, and is strictly tighter than the title's own gap. Recording the residual is deliberate; closing it exactly (measuring the period header at its own render size) is a named follow-up, not part of this change.
- **D3 — measurement failure falls back to the floor.** A period whose measurement is unavailable, non-finite, or negative falls back to `MinimumHeaderLabelWidthPt`. A missing or misbehaving seam can therefore neither open the gate (a zero/negative width would show every label) nor close it (an infinite width would suppress every label); fit-gating activates only when a measurement genuinely succeeds. (A seam that fails the **title** too is already a separate, existing refusal before the sequence is built.)
- **D4 — the band stays uniform.** This is measured-fit, not ADR-0039's ladder. The format remains one scene-wide choice emitted verbatim; only the show/hide decision varies per interval, and it varies by the same rule for all of them. There is no width-driven text substitution of any kind. `MM` is not "stepped down to" — it is *selected*, and being shorter it simply fits more intervals.
- **D5 — scope is Month-visible but not Month-specific.** The measured-fit rule applies to every period scale (`Month`, `Quarter`, `Week`); a quarter (`Q1`) or week (`W01`) label is shown only when it measures to fit, on the same rule. No abbreviation forms are introduced for any scale.

## Consequences

- `MM` now reveals month headers on intervals where `MMM` suppresses them — the functional difference the owner asked for, with a uniform band.
- The label primitive's role-derived identity (`{parent}:label`, R3.17) and the R4.7/R4.8 reconciliation keys are unaffected: the text is still the selected format, only visibility changed.
- Live, export, and raster renderers all consume the same Core scene, so measured-fit is identical in all three; no renderer-equivalence divergence.
- No catalogue token, schema version, settings-hash, or format-enum change — the reference scene's month intervals are wide enough that `MMM` still fits at the 8pt seam, so the golden snapshot is unchanged.
- **Recorded residual:** a borderline period label could clip by roughly the 8→9pt render/seam ratio. Same class as, and smaller than, the title ellipsis's existing residual.

## Alternatives considered

- **Keep the fixed 18pt floor (status quo)** — rejected: `MM` and `MMM` show on identical intervals, so `MM` is cosmetic, which is the defect the owner reported.
- **The automatic MMM→MM ladder** (ADR-0039, still rejected) — non-uniform per-interval output; the owner explicitly wants the user to decide, not the engine to step.
- **A floor of `2 × measured` (option C from the review)** — rejected in favour of B: it adds a legibility gutter and absorbs the seam residual, but the owner chose B for maximum label density; the residual is recorded and closable as a follow-up.
- **A period-header-sized (9pt) measurer** — rejected for now (D2): a second font-size authority; the shared-seam residual is acceptable and matches the title. Named as the follow-up if exactness is wanted.
- **Pin the period header font family to Aptos** — not taken: a visual/export change outside this task's scope; the host default already renders Aptos on current Excel, so the family is exact there.

## Implementation status

**Accepted; implemented on the R5.12 branch (2026-10-10).** `BandSequence.TryCreate` takes an optional `ITextWidthMeasurer periodMeasurer`; `CreatePeriods` resolves each label's required width via `RequiredPeriodWidth` (measured width, floor on failure/non-finite/negative) and passes it to `AddInterval`, whose `ShowLabel` is `intervalWidth >= requiredWidth`. Years unchanged. `FrameBandsBuilder.TryBuild` passes its `textMeasurer` into the sequence call. Core tests rewritten/added: `MM_is_shown_where_MMM_is_suppressed_on_the_same_plot`, `Narrow_periods_suppress_labels_and_warn` (now measurer-driven), `A_failed_measurement_falls_back_to_the_floor`, `Quarter_and_week_labels_are_also_fit_gated`. Guide §6 records the measured-fit rule (revision 12). Observed: Core 1480/1480, Architecture 117/117, Office contract 748/748, golden snapshot unchanged.
