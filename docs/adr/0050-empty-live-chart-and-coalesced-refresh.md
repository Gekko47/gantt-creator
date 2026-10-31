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
