# Remediation implementation plan (Commit A)

Derived from `docs/GanttCreator_R3_R4_Remediation_Brief.md` after verifying every
"Confirmed code" claim against `feat/r4-1-adapter-ports` @ `41fa0cc`.

Scope: **Commit A only** (Core structural-row composition). Commits B–F are
recorded as deferred and are not started. Decisions taken 2026-09-28:

- **D-A1** — the §10 splitter representation is a new `SplitterBuilder` reusing the
  existing style pipeline (`EntityTypeCatalog` already names `Splitter` as the default
  key; `GanttCatalogues` already defines the `Splitter` preset), band at
  `ZLayer.Section` (30), border lines at `ZLayer.Frame`, splitter labels placed by the
  builder as `DelineatorBuilder` already does for its corner labels.
- **D-A2** — `LaneLayoutBuilder` accepts an empty input and returns an empty successful
  layout, rather than `SceneBuilder` skipping the lane phase. The emptiness question
  belongs to the orchestrator, not the geometry builder.
- **D-A3** — the behaviour commit lands with `GoldenSceneSnapshotTests` showing the
  primitive-level delta; the golden is regenerated in its own commit after human
  approval, per `AGENTS.md` and the D4 golden policy.

## Verification results (brief claim vs. repo)

| Brief | Verdict | Evidence |
| --- | --- | --- |
| §1.1–§1.5 preserve R4.1–R4.5 | Confirmed | `ChartOriginDelta`; `TextFrame2`; `AddDiamond`; `CarriesOwnershipTagFor`; `ResolveShapes` → `QueryTarget` |
| §1.6 "SendToBack + BringForward" | **Stale** | Current code is reverse-order `msoSendToBack` only, no `BringForward` (`ExcelShapeWriter.cs:324-327`). Intent (no `BringToFront`) holds |
| P0-1 Splitter/Spacer dropped | Confirmed, and worse than stated | `SceneBuilder.cs:260-266`; `LaneLayoutBuilder.BuildFixedLane` (`:117-131`) unreachable from the top. **The fixture already carries 1 Splitter + 1 Spacer and the golden has none** |
| P0-2 two panel authorities | Confirmed | `SceneBuildRequest.PanelBounds` → `FrameBandsBuilder.cs:153`; `PanelBuilder` derives independently (`:200-203`, `:259`, `:265`); no validator |
| P0-3 single row height | Confirmed, plus a second site | `PanelCellGrid.RowHeightPt`; `ExcelPanelGridMeasurement.HasUniformRowHeight`; **also** `SceneBuilder.cs:621` uses the body row height as the panel *header* height |
| P0-4 panel rows from placements | Confirmed | `SceneBuilder.cs:618` |

## Two findings the brief does not name, both fixed in A

1. **A Splitter sharing a `LaneId` with an activity would be mis-routed.**
   `LaneOrdering.LaneKey` returns `LaneId.Value` verbatim (`LaneOrdering.cs:20-24`), and
   `LaneLayoutBuilder.cs:103` picks the lane branch from `laneInputs[0].Event.Type`. A
   shared `LaneId` therefore groups a Splitter with activities and takes the wrong
   branch. §10 says a Splitter "occupies a complete lane", so its lane key is
   forced row-scoped regardless of `LaneId`.
2. **A zero resolved height would be supplied for Splitter/Spacer.**
   `SceneBuilder.cs:302-304` exempts them from style resolution and `:316` assigns
   `new ResolvedEventStyle(new SceneStyle("None"), 0)`. The fixed-lane branch does not
   read the input height today, so this is latent — but the input would not state the
   height it occupies. A supplies `laneMetrics.SplitterHeightPt` / `SpacerHeightPt`
   explicitly instead of relying on that.

## Phases

- **Phase 1** — split `renderable` into `laneParticipants` / `laneForegroundTypes` /
  `plotGlobalEntities`. `panelSourceRows` is **not** introduced here (that is B).
  Acceptance: `rg "renderable" src/` is empty; no behaviour change yet.
- **Phase 2** — feed `laneParticipants` into the `LaneEventInput` loop so
  `BuildFixedLane` is reachable; force a row-scoped lane key for Splitter/Spacer.
  Tests: `activity → Splitter → activity` displaces by exactly `SplitterHeightPt`;
  `activity → Spacer → activity` by exactly `SpacerHeightPt`; both together contribute
  exactly once; a Splitter sharing a `LaneId` still gets its own fixed lane.
- **Phase 3** — new `src/GanttCreator.Core/Scene/SplitterBuilder.cs` per D-A1: a band
  from panel-left to plot-right, `SplitterFill`, top/bottom border at
  `MajorBoundaryPt`, and labels at `DataPanelLeft` (default) / `PlotCentre` / `Both`
  (two texts) / `None`. Roles `splitter-band` / `splitter-label` /
  `splitter-label-plot`.
- **Phase 4** — empty-scene semantics at **both** sites: `SceneBuilder.cs:267-270`
  and `LaneLayoutBuilder.cs:56-59` (per D-A2). The existing `EmptyInput` positive test
  is rewritten — a second test this change deliberately breaks.
- **Phase 5** — delete the `SceneBuilder.cs:292` literal; route the delineator
  fallback through the `DefaultDelineator` preset.
- **Phase 6** — rewrite
  `A_splitter_and_spacer_alone_are_refused_because_neither_has_an_entity_primitive`.

## Deferred

| Commit | Work | Blocked on |
| --- | --- | --- |
| B | Per-row `PanelCellGrid`; one panel-bounds authority; panel rows from source rows not `SceneBuilder.cs:618` | Approval. **`:618` is inside A's blast radius**: a Splitter in a lane will now flow into the panel row list. A keeps panel behaviour byte-identical to today so B stays separable |
| C | `IPanelGridMeasurementPort` per-row measurement; mixed heights; the read-only protection question | B |
| D | `ApplyZOrder` ownership preflight (`ExcelShapeWriter.cs:305-315`) | Independent and small |
| E | R4.6 style completeness | — |
| F | R4.7 identity/reconciliation | E |

  `An_empty_event_list_is_refused` and `A_hidden_only_event_list_is_refused` stay.
- **Phase 7** — golden snapshot gate per D-A3.
- **Phase 8** — `REPO-MAP.md`, `07-GANTT-ENTITY-GUIDE.md` §10, `DECISIONS.md`,
  `STATUS.md`, `03-ROADMAP.md`.
- **Phase 9** — targeted tests, then `verify-quick.ps1`, then `verify.ps1` on a clean
  tree; `git diff --check`; `git status --short`; full diff review.
  `verify-office.ps1` is **Not run** by A and is recommended before R4.6.

Checklist self-certification for A: **A, B, C, E, G, I**. D, F, H, J, K do not apply
(no config sheet, no date handling, no suppression-scope change, no COM, no failure
injection) and are not claimed.

| P1-1 style staleness | Already sequenced in the R4.7 work item | No new code in A |
| P1-2 `ApplyZOrder` ownership | Confirmed | `ExcelShapeWriter.cs:305-315` resolves by name, then `:324` mutates. **Deferred to D** |
| P1-3 name ceiling | R4.7-only | `MaxShapeNameLength` (`:79`), `Create` refuses pre-host (`:94`) |
| P1-4 delineator hard-coded style | Confirmed | `SceneBuilder.cs:292` duplicates `GanttCatalogues.cs:177` |
| P1-5 delineator-only = empty | Confirmed, and **incomplete** | `SceneBuilder.cs:267-270`; the brief misses that `LaneLayoutBuilder.cs:56-59` then refuses `EmptyInput` → `InvalidLayoutSettings` |
| P2-1 read-only protection | Confirmed, but the brief overstates | `ProtectionGuardFirstTests` classifies by mutation discovery, not guard calls, so removing the guard needs **no** architecture-test change. Architecture/ADR-0008 decision — **deferred** |
| P2-2 `StackIndex` | Confirmed | Never read by `LaneOrdering`/`LaneLayoutBuilder`. Product-owner call — **deferred** |
| P2-3 catalogue non-transactional | Confirmed | 4+ sequential `WriteOrReplaceTable` under one `try`. Logged as a known limitation |
