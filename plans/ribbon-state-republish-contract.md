# Contract ruling — when the ribbon snapshot is republished

> **Status**: Implemented 2026-11-27 (code, tests, docs); the Office host gate is owed.
> **Relates to**: R5.11, R5.20, ADR-0050 D3/D4/D6, ADR-0055, ADR-0058.
> **Settles**: the two open questions from the code review of the Type-change fix.

## The single rule

**The ribbon snapshot is a projection of the table, never of the render, and it is
never suspended. It is republished at the end of every pass (whatever the outcome),
and on every accepted table change whose render is suspended. Render facts inside it
stay success-only.**

## Question 1 — does a refused pass republish?

**Ruling: yes, every terminal outcome republishes the sheet facts.**

- The snapshot is a read-only projection of the table (`Type`, `StyleKey`, label
  cells, hierarchy), not of the render. A refused pass can still have written cells
  (type-change normaliser, `Duration`, validation notes), so the ribbon must describe
  the sheet as it now is.
- The postlude already makes this split: `UpdateCurrentValidationIssues` runs on both
  paths, `UpdateWarnings` and `MarkChartRenderCurrent` only on success
  ([`RefreshSheetCommand.cs`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs:295)).
  The republish joins the first channel.
- A Type edit that creates a blocking error is the case the user most needs explained;
  leaving the old Type's controls on screen is the defect this change removes.

**Guardrails.**
- Sheet facts only. `PlotStart`/`PlotFinish` stay success-only, so the AUTO date text
  never claims a range the chart did not render (R5.1a D1).
- No reload of plot settings and no date re-derivation in the republish: the outcome
  already carries the range, and the outcome fields exist to avoid exactly that second
  derivation ([`IGanttRefreshOrchestrator.cs`](../src/GanttCreator.Office/IGanttRefreshOrchestrator.cs:144)).
- Sticky capture policy unchanged: a failed read keeps the previous snapshot.

## Question 2 — does the snapshot republish while the render is suspended?

**Ruling: yes. The ribbon is corrected irrespective of the auto refresh and the chart
lock (owner ruling 2026-11-27).**

- The owner's requirement is stronger than the analogy: the ribbon must always be
  correct and reflective of the table. Coupling it to the pass would leave the ribbon
  describing a stale table for as long as the switch is off, which is ruled out.
- Neither suspension rule is about the ribbon. ADR-0050 D6 suspends edit- and
  command-driven *passes*; ADR-0055 clause 7 keeps a rebuild from undoing a manual
  drag. A snapshot republish is a read plus a UI invalidation: it mutates no shape and
  no cell, so it cannot undo a drag and is not a pass.
- R5.20's accepted trades list sheet reconciliation and the Warnings list. They do not
  list the ribbon, and must not be read to cover it. The work item gains one
  clarifying line; the switch keeps exactly the meaning it has for the sheet.

**Rejected alternative.** Coupling the republish to the pass, on the "same trigger as
the plot" reading. It needs no new scheduler action, but it fails the owner's ruling
and leaves the Type controls wrong in the switch-off mode.

**Consequences.**
- The scheduler gains a second outcome on the same debounce: `OnTimerElapsed` runs the
  pass when allowed, and a read-only snapshot republish when suspended. One timer, one
  coalescing window, two outcomes.
- The switch keeps its full meaning for the sheet: `Duration`, row heights, outline
  levels, hidden columns, protection re-arm, error fills and the Warnings list still
  wait for a pass. The ribbon does not.
- While the chart is unlocked, a shape drag raises no cell change, so this adds no
  republish storm and cannot fight a manual edit.
- Cost: one table read plus the four context reads per accepted burst while suspended.
  Mitigated by the `LastTableRows` reuse below; measure with `gantt.refresh post` and
  the scheduler's `detect` trace, and report the number rather than claiming one.

## Consequences (all accepted, none silent)

- A Type edit that refuses the pass still republishes: the chart keeps its last state,
  the ribbon describes the new Type, and the validation notes explain the error. That
  is the split Q1 rules.
- With the automatic refresh off or the chart unlocked, the ribbon still tracks the
  table. The plot does not, by the owner's earlier ruling, so the two deliberately
  disagree while the switch is off, and the explicit Refresh reconciles the plot.
- The republish adds one table read plus the four context reads on every pass and on
  every suspended burst. The `SnapshotSources.LastTableRows` field already exists for
  the mitigation: the contexts should consume the supplied rows instead of calling
  `reader.Read()`. That is the deferred change named below.

## Implementation outline (code mode, revised against the tree)

**1. `RibbonStateService` — probe the facts, then capture.**
Extract the facts block from `Refresh()`
([`RibbonStateService.cs`](../src/GanttCreator.AddIn/RibbonStateService.cs:2006)) into
one private helper returning the six `SnapshotProbeFacts` plus the rows. Both `Refresh()`
and the new method must use it. `Capture` resolves every fact as `facts.X ?? previous.X`
([`RibbonSnapshotFactory.cs`](../src/GanttCreator.AddIn/RibbonSnapshotFactory.cs:83)), so
passing only `SheetInitialised` would leave `CriticalIntervalParentAvailable` and
`HasGroups` untouched — and `CriticalIntervalMenuItem`, `ExpandAllGroups` and
`CollapseAllGroups` would stay stale, which is the defect this change exists to remove.

**2. `RibbonStateService.RefreshAfterTableChange(DateOnly? plotStart = null, DateOnly?
plotFinish = null, bool measurePlot = false)`.** The order is load-bearing:
1. probe the facts and rows; assign `_lastTableRows` when the read succeeded, because
   `CaptureRowActionContext` reads it ([`RibbonSnapshotFactory.cs`](../src/GanttCreator.AddIn/RibbonSnapshotFactory.cs:195));
2. `CaptureSnapshot`;
3. when `measurePlot`: `LoadColumnVisibilityFromWorksheet()` then
   `WithMeasurement(MeasurePlot(_hiddenColumns))` — exactly what the explicit Refresh gets
   today ([`RibbonStateService.cs`](../src/GanttCreator.AddIn/RibbonStateService.cs:1846),
   [`RibbonStateService.cs`](../src/GanttCreator.AddIn/RibbonStateService.cs:2050));
4. `UpdateEffectivePlotDates(plotStart, plotFinish)`;
5. `Invalidate()`.

Never `LoadPlotSettingsFromWorkbook` (it re-derives the AUTO dates from a second read),
never `ActivateGanttTab`, no cell or shape write.

**3. `RefreshSheetCommand`.**
- `FinishRefresh` loses both `updateEffectivePlotDates` and `invalidateSelectionGates`.
  Always call step 2 after `UpdateCurrentValidationIssues`, with
  `measurePlot: !automatic`. Success keeps warnings, chart-current and the status bar;
  a refusal keeps its presenter.
- `RunForExcel` now applies the outcome dates, which it discards today
  ([`RefreshSheetCommand.cs`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs:212)); the
  orchestrator populates them ([`GanttRefreshOrchestrator.cs`](../src/GanttCreator.Office/GanttRefreshOrchestrator.cs:944)).
  That is a labelled behaviour change: the explicit path's AUTO text comes from the pass,
  not from a second derivation.
- `RunSilentForExcel` keeps its intent; the comment at
  [`RefreshSheetCommand.cs`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs:224) becomes
  false and is rewritten.

**4. `GanttRefreshScheduler`.**
- Add an injectable `Action? stateRefresh`, the singleton method in production.
- `OnSheetChanged`: keep the `_commandRefreshPending` guard, delete the `IsSuspended`
  early return ([`GanttRefreshScheduler.cs`](../src/GanttCreator.AddIn/GanttRefreshScheduler.cs:270)),
  and always reset the timer for an accepted change.
- `OnTimerElapsed`: replace the suspended `return`
  ([`GanttRefreshScheduler.cs`](../src/GanttCreator.AddIn/GanttRefreshScheduler.cs:402))
  with a branch. The projection branch must queue a macro, re-check the generation, trace
  with `source=state`, clear `_commandRefreshPending` as
  [`GanttRefreshScheduler.cs`](../src/GanttCreator.AddIn/GanttRefreshScheduler.cs:432)
  does, and call `stateRefresh`. Skipping the clear swallows the next edit.
- `ScheduleCommandRefresh` stays suspension-aware: commands are projected by their own
  callback ([`GanttRibbon.cs`](../src/GanttCreator.AddIn/GanttRibbon.cs:2643)). A command
  that writes cells while suspended therefore also queues one projection, so it costs two
  captures. Accepted and recorded.

**5. Wiring and ribbon.**
- `RibbonEventWiring.CreateScheduler` must pass the delegate
  ([`RibbonEventWiring.cs`](../src/GanttCreator.AddIn/RibbonEventWiring.cs:156)); the
  production constructor cannot reach the singleton by itself.
- `GanttRibbon.OnRefreshSheetClick` drops `NotifyRibbonStateChanged`
  ([`GanttRibbon.cs`](../src/GanttCreator.AddIn/GanttRibbon.cs:243)). Every other command
  keeps its callback.

**6. The toggle must not strand a burst.**
`SetAutoRefresh(false)` cancels the pending timer
([`RibbonStateService.cs`](../src/GanttCreator.AddIn/RibbonStateService.cs:2543)) and only
invalidates the toggle
([`RibbonStateService.cs`](../src/GanttCreator.AddIn/RibbonStateService.cs:2558)) — it does
**not** recapture. The cancelled action may now be a projection, so add
`RefreshAfterTableChange()` after the settings write and before `ActivateGanttTab`.
`SeedPersistedSwitches` assigns the preference directly
([`RibbonStateService.cs`](../src/GanttCreator.AddIn/RibbonStateService.cs:1790)), so
activation stays free of this call.

**7. Tests to revise, not only add.**
- [`RefreshSheetCommandTests`](../tests/GanttCreator.AddIn.Tests/RefreshSheetCommandTests.cs:249)
  counts seven `InvalidateControl` calls. `Invalidate()` calls `ribbon.Invalidate()`, so
  the count becomes zero (plus the two date controls when the dates move). Assert one
  `Invalidate` and assert the snapshot's contexts instead.
- [`GanttRefreshSchedulerTests`](../tests/GanttCreator.AddIn.Tests/GanttRefreshSchedulerTests.cs:239)
  and the two facts below it assert that *no* macro is queued while suspended. Split the
  fact: the refresh delegate never runs, the state delegate runs once. `CreateScheduler`
  ([`GanttRefreshSchedulerTests.cs`](../tests/GanttCreator.AddIn.Tests/GanttRefreshSchedulerTests.cs:524))
  gains the delegate.
- [`AutoRefreshRibbonTests`](../tests/GanttCreator.AddIn.Tests/AutoRefreshRibbonTests.cs:42)
  and [`RibbonStateSnapshotTests`](../tests/GanttCreator.AddIn.Tests/RibbonStateSnapshotTests.cs:463)
  call `SetAutoRefresh(false)`; the added capture must leave their assertions intact
  (toggle invalidated once, tab restored once).
- [`GanttRibbonCommandTests`](../tests/GanttCreator.AddIn.Tests/GanttRibbonCommandTests.cs:236)
  injects its own callback, so it survives; add a source fact that the production method
  passes none.

**8. Docs.**
Name a work item (none exists for this defect), record the ruling in
`docs/DECISIONS.md` and `docs/STATUS.md`, add the one line to the R5.20 accepted trades,
and fix the stale comment named above.

## Acceptance tests (each must fail with the change reverted)

- A Type edit that the pass renders: the ribbon's description-label options, date-label
  vocabulary and custom-style enablement resolve for the NEW type.
- A Type edit that refuses the pass (`BlockingValidationErrors`): the same three resolve
  for the new type, and the AUTO plot-date text keeps its previous value.
- With `AutoRefreshEnabled` off, a Type edit republishes the snapshot and the refresh
  delegate is never invoked.
- With the chart unlocked, the same: republish yes, pass no.
- `OnTimerElapsed` picks exactly one of the two actions for a burst, so a suspended
  burst cannot both republish and render.
- A change that flips a data-gated control: `CriticalIntervalMenuItem` re-asks, so only
  a full invalidation passes it.
- The Refresh button refreshes the ribbon exactly once, not twice.
- Toggling the switch off inside a debounce window leaves the ribbon projected.

## Deferred

- `RibbonSnapshotFactory` contexts reading `SnapshotSources.LastTableRows` instead of
  calling `reader.Read()`: one read per capture instead of five. Separate change, with
  its own measurement.
