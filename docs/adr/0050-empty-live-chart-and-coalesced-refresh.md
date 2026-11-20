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

## Amendment 2026-11-16 — D6: the automatic refresh is switchable, and the switch is persisted

**Ruling.** The owner asked for a control that turns the coalesced automatic refresh off, leaving
the explicit **Refresh chart** command as the only rebuild, and that remembers the choice in the
file. D3's "chart-affecting edits … use one debounced scheduler" therefore becomes conditional on
the user's switch.

**Decision.**

- One persisted key, `tblGanttSettings.AutoRefreshEnabled` (schema version 25), written by exactly
  one command and re-read when a workbook is activated or switched. Parsing is fail-open: absent or
  unreadable means enabled, which is the behaviour every earlier release had.
- The switch suspends **edit- and command-driven** passes only, in the scheduler's three existing
  check sites, with `CancelPendingRefresh` discarding a pass already inside its window. It does not
  touch the explicit Refresh, which remains the full validation and repair pass D4 requires, and it
  does not touch the activation and workbook-switch reconciliation ADR-0058 depends on.
- Switching the automatic refresh back on schedules no pass. The next chart-affecting edit or an
  explicit Refresh reconciles, so the control itself adds no render.
- The surface is a `toggleButton` inside the Refresh split button's menu — the switch belongs to
  the action it governs — with its pressed state and label stating the setting.

**Consequences.** With the switch off, engine-derived `Duration`, row heights, outline levels,
hidden columns, protection re-arm and error fills are reconciled only by a pass, and the Warnings
list keeps the last render's findings; R5.20 records this and the control's supertip states it. A
version-24 workbook reports a catalogue-hash mismatch and is repaired by Initialise (ADR-0029 D6,
no migration).

## Amendment 2026-11-20 — D3: the coalescing window is sized by the burst

**Decision.** One cell window becomes two. The **first** accepted cell change of a burst waits
**100 ms**; a **second** change in the same burst escalates the window to **300 ms**, which is what
keeps a paste or a fill collapsing into one render. A command keeps its **50 ms**, and the burst
counter resets when a burst's timer fires, so the next lone edit is fast again.

**Why.** The 13:3x trace split click-to-plot into **300 ms of debounce + 10–115 ms of macro queue +
164–1072 ms of pipeline**, making the fixed window the largest single term for a one-cell edit, while
`events=9–10` on the same log showed the bursts that need coalescing are pastes whose events arrive
together (nine inside 3.4 ms). The escalation only has to win that ~0.4 ms gap, so 100 ms is ample
margin for a slow host delivering an event late.

**Why it needed a prerequisite, and what that was.** A first attempt at this amendment was
implemented, contract-tested and then **withdrawn**, because the live gate failed two different
`AddRowIntegrationTests`. Those facts asserted row arrangement and shape set immediately after an
insert and had always been racing the scheduler — the 300 ms window merely outlasted them; at 100 ms
the add-in's scheduled pass, triggered by `Initialise`'s own cell writes, landed mid-assertion. The
amendment therefore lands with that family made **scheduler-independent**: it no longer calls
`RegisterXll`, because nothing in it needs the add-in (it drives the Office adapters directly), which
is the same deliberate omission already documented on `ColumnVisibilityRoundTripIntegrationTests` and
`ProtectionArmingIntegrationTests`.

**Residual trade, recorded rather than hidden.** Two edits 100–300 ms apart now render twice instead
of once. That is one action per render, and each render is 130–250 ms of write work.

**Measured effect.** ~200 ms off the perceived latency of a one-cell edit on the automatic path, with
paste coalescing unchanged. Evidence and the per-burst trace columns are in
`plans/refresh-latency-trace-and-opportunities.md` §P.
