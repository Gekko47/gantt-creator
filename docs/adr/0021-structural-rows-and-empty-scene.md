# ADR-0021 — Structural rows reach lane layout; "empty" means no visible entity

- **Status:** Accepted
- **Date:** 2026-09-28
- **Relates to:** entity guide §9, §10, §11, §24; ADR-0012; R3.4; R3.12
- **Supersedes:** the `LaneLayoutRefusal.EmptyInput` contract and the
  `SceneBuilder` "renderable" list

## Context

`SceneBuilder` filtered every visible event into one list that excluded `Splitter`,
`Spacer`, and `Delineator`, then fed only that list to `LaneLayoutBuilder`. Two
distinct defects followed.

1. `LaneLayoutBuilder.BuildFixedLane` — the §10/§11 fixed-height lane branch, with
   its own tests — was unreachable from the top-level composition. A `Splitter` never
   occupied a lane and never emitted anything, so §10 geometry was dead code that the
   test suite could not detect as unused.
2. `Delineator` is a §24 plot-global entity that renders a line and consumes no
   lane. Because it was excluded from the lane-bound list, the emptiness guard
   (`renderable.Count == 0`) classified a delineator-only scene as `EmptyEvents` —
   reporting a chart with a line to draw as having nothing to render.

The `EmptyInput` refusal in `LaneLayoutBuilder` was the mechanism that made defect 2
unfixable at the orchestrator: a delineator-only scene produces zero lane inputs, so
the builder refused, and `SceneBuilder` mapped that to `InvalidLayoutSettings`.

A third defect was found while implementing this and is named here because it shares
the same root cause. `LaneOrdering.LaneKey` returned `LaneId.Value` verbatim and
`LaneLayoutBuilder` chose its lane branch from the first input in a group. A `Splitter`
sharing a `LaneId` with an activity would therefore be grouped with the activity and
built as an ordinary event lane, contradicting §10's "a complete lane".

## Decision

- **D1 — Four named row categories, not one "renderable" list.** `SceneBuilder` now
  partitions the visible rows into `visible`, `laneParticipants` (everything that
  occupies a lane, including `Splitter` and `Spacer`), and `plotGlobalEntities`
  (`Delineator`). A fourth category, the panel source rows, is deliberately *not*
  introduced here: panel membership is still derived from lane placements, and
  correcting that is a separate breaking change to `PanelCellGrid`.
- **D2 — An empty lane input is a successful empty layout.** `LaneLayoutRefusal
  .EmptyInput` is **removed** rather than deprecated. Whether the scene has content
  at all is the orchestrator's question; a geometry builder refusing "no lanes"
  forced `SceneBuilder` to conflate a lane-less scene with an empty one.
  `LaneEventLayout.TryBuild` changed the same way, for the same reason, and its
  `MissingLayout` refusal now means *null* layout only — coverage of a non-empty
  input is still checked by the existing completeness count.
- **D3 — "Empty scene" means no visible scene-producing entity.** The guard is
  `laneParticipants.Count == 0 && plotGlobalEntities.Count == 0`. A hidden-only
  scene and a genuinely empty request still refuse `EmptyEvents`.
- **D4 — A `Splitter` or `Spacer` always owns its lane.** `LaneOrdering.LaneKey`
  returns a row-scoped key for those two types regardless of `LaneId`, so §10/§11
  cannot be defeated by a shared lane identifier.

## Consequences

- §10 and §11 lane geometry becomes reachable and is now covered by top-level
  `SceneBuilder` tests, not only by `LaneLayoutBuilder` unit tests.
- The committed golden scene snapshot **changes**: `reference-gantt.json` already
  contained one `Splitter` and one `Spacer` that the old code discarded, so lane Y
  positions shift and splitter primitives appear. Per the D4 golden policy the
  regeneration is a separate, human-reviewed commit — see `docs/STATUS.md`.
- A `Splitter` sharing a `LaneId` with an activity now renders as its own band. A
  workbook that relied on the old mis-routing will look different; that is the
  §10-correct rendering.
- Panel membership is now visibly coupled to lane placement, because structural rows
  reach lane layout. This is recorded as a known coupling owned by the next change,
  not treated as correct.
- Two tests that encoded the old behaviour are rewritten: the `EmptyInput` refusal
  case and `A_splitter_and_spacer_alone_are_refused_…`. They were evidence of a
  defect, not of a contract.
