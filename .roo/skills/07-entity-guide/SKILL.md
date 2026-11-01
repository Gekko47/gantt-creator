---
name: 07-entity-guide
description: Entity geometry, style tokens, label positions, z-order, the minimum visual reference fixture, and the entity-to-renderer equivalence table for all 26 Gantt entity types in the Gantt Creator add-in. Part of the gantt-creator index; the largest contract, so load docs/07-GANTT-ENTITY-GUIDE.md directly when the summary is not enough. Trigger on any visual, layout, style, label, or export change.
---

# Gantt entity contract

Mirror of [`.cline/skills/07-entity-guide/SKILL.md`](../../../.cline/skills/07-entity-guide/SKILL.md),
generated from [`docs/07-GANTT-ENTITY-GUIDE.md`](../../../docs/07-GANTT-ENTITY-GUIDE.md).
The canonical doc wins on any disagreement; it is the implementation contract, this is a pointer.

## Index

- Index first: `gantt-creator-contract`; `00-ponytail` before any code change.
- Pair with `02-architecture` (scene-first rendering, renderer equivalence) and
  `04-test-strategy` (golden images and the equivalence thin slice).

## Do not get wrong

- The numeric and colour values in the token tables are **initial design defaults, not visually
  approved facts** — an agent must not change them. Only an approved product-owner review
  updates the doc, the tests, and the golden images together.
- If a requested visual behaviour is not defined in the guide, it is `unknown` — stop for a
  decision. Never infer it from a screenshot.
- Milestones and delineators use `Start` only for geometry; `Finish` is never used for their
  positioning.
- `Type` values come only from `EntityTypeCatalog`; per-row `FillColour` / `StrokeColour` /
  `LabelPosition` overrides are validated against the selected type's capabilities, and
  `Custom Activity` / `Custom Milestone` require a valid `StyleKey`.
- Renderers consume the resolved scene and never reinterpret a primitive: the critical interval
  is a filled rectangle everywhere, z-order follows the layer table exactly, and every
  renderer's output matches the guide's per-entity field contract.

## Tools here

`read_file` (the guide, per-entity sections) · `codebase_search` / `search_files` (find a token,
z-order constant, or renderer-equivalence entry) · `execute_command` (`dotnet test` on Core,
Raster, and Office contract golden/equivalence tests).
