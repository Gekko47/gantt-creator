# ADR-0050 — Empty live chart and coalesced automatic refresh

- **Status**: Accepted
- **Date**: 2026-10-31
- **Relates to**: ADR-0021, ADR-0020, R5.8, R5.11
- **Supersedes**: ADR-0021 D3 only for an explicitly empty live-chart request

## Context

ADR-0021 preserves `EmptyEvents` for a genuinely empty scene. The owner has since
requested that first-time initialization produce an empty Gantt frame once, and
that later chart-affecting edits refresh automatically. This requires an explicit
empty-live-scene mode while preserving the default refusal for export and other
callers.

## Decision

1. `SceneBuildRequest.AllowEmptyEvents` is opt-in and defaults to false. The live
   Excel factory opts in, allowing an empty table to render the frame and title.
   The measured panel grid likewise permits zero body rows only by explicit opt-in;
   the default `NoRows` refusal remains for existing callers.
2. Successful first-time initialization queues one debounced render. A repeated
   initialization refusal queues no render.
3. Chart-affecting table edits and successful setting/structural commands use one
   debounced scheduler. Selection, navigation, cancellation, and pre-mutation
   refusal do not render.
4. Every scheduled or explicit render still reads and validates all current data,
   measures the complete layout, builds the complete scene, and only then mutates
   owned shapes. Explicit Refresh remains the full reconciliation path.
5. R4.10 performance measurement and benchmark work remain explicitly excluded.

## Consequences

The live chart can represent an empty schedule without weakening the default
scene-builder contract. Automatic rendering is coalesced and scheduled refreshes
use an in-memory content signature to skip unchanged shape writes. Contract tests
cover the delta planner and fallback behavior. The general OfficeIntegration
suite passes, but feature-specific host proof for automatic triggers, empty
initialization, delta operation counts, and protection feasibility remains due.

## Amendment 2026-11-16 — D5: an all-invalid table with no chart renders the empty chart

**Report.** "Why does the first refresh after adding an activity with no dates create this error ...
yet once the first set of dates are populated the subsequent activities allow a plot with no errors
even when there are dates missing?"

**The asymmetry, stated exactly.** The orchestrator refused only when **no row could be drawn**
(`validation.Events.Count == 0`) *and* a blocking error existed. One scaffold row with no dates is
that state, so the first Add raised a modal "could not complete the last action" on a sheet with no
chart to protect; with one valid row present, the same dateless row is simply dropped from the scene
(partial render) and nothing is reported.

**Decision.** The refusal now depends on whether a chart EXISTS, and
`IShapeWritePort.ListOwned()` — a read, so no shape is mutated — decides it.

- **No owned shape:** the empty chart is rendered (the frame and its period bands, no bars) and the
  scene carries a `SceneWarning` whose code is `SceneWarningCodes.EmptyChart`, which the Warnings
  panel reports. A normal first step no longer presents an error dialog.
- **Owned shapes present:** the refresh still refuses with `BlockingValidationErrors`, so a
  transient all-invalid table (the user cleared a date) cannot wipe a chart they were looking at.

**Consequences.** `SceneBuilder` gains `SceneWarningCodes.EmptyChart` and emits it whenever the live
request allows empty events and no row can be drawn. The two blocked-refresh facts in
`GanttRefreshOrchestratorTests` were rewritten to the new contract: one asserts the empty chart, the
warning and zero row-owned shapes; the other seeds one owned shape and asserts the refusal and zero
mutations. The single `ListOwned` read on the refusal path is the decision itself, and the test says
so. `SceneBuilderTests.An_empty_live_chart_carries_the_reason_it_is_empty` pins the scene half.
