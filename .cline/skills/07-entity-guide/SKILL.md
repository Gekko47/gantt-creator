---
name: 07-GANTT-ENTITY-GUIDE
description: Entity geometry, style tokens, label positions, z-order, the minimum visual reference fixture, and the entity-to-renderer equivalence table. Largest skill; load docs/07-GANTT-ENTITY-GUIDE.md directly when the summary is not enough. Trigger on any visual, layout, style, or export question.
---

The implementation contract for every visible Gantt entity: shared
geometry/colour/typography tokens, z-order, and a per-entity
geometry/style/label/validation/renderer-equivalence spec for all 26
entity types (chart frame through legend and validation indicators).

Do not get wrong:
- The numeric and colour values in the token tables are **initial
  design defaults, not visually approved facts** — an agent must not
  change them; only an approved product-owner review updates this
  file, tests, and golden images together.
- If a requested visual behaviour is not defined here, it is
  `unknown` — stop for a decision. Never infer it from a screenshot.
- Milestones and delineators use `Start` only for geometry; `Finish`
  is never used for their positioning.

---

## Where to read more

- Full reference: `docs/07-GANTT-ENTITY-GUIDE.md` (this is the canonical source; there is no separate copy under .cline/skills/)
