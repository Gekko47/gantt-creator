---
name: 07-GANTT-ENTITY-GUIDE
description: Use this skill when the conversation is about entity geometry, style tokens, label positions, z-order, the minimum visual reference fixture, the entity-to-renderer equivalence table, or any visual / layout / style / export question. This is the largest skill; load references.md when the summary is not enough.
---

# Gantt visual entity guide — revision 3
> Created 2 September 2026 under a new filename. This revision uses full Type names backed by `GanttCreator.TypeOptions` on `_GanttCreatorConfig`, while keeping schedule rows and per-entity overrides on the visible worksheet. When installed, use the path `docs/07-GANTT-ENTITY-GUIDE.md`.
## Purpose and authority
This is the implementation contract for every visible Gantt entity. It prevents each renderer—or an LLM—from making independent decisions about geometry, colour, labels, z-order, clipping, or export behaviour.
The named tokens and rules are authoritative for implementation. The numeric and colour values in the initial token tables are **initial design defaults, not visually approved facts**. An agent must not change them. The product owner reviews them through the Phase 3/5 reference render; an approved change updates this file, tests, and golden images together.
If a requested behaviour is not defined here, classify it as `unknown` and stop for a decision. Do not infer it from a screenshot alone.
## Shared entity contract
Every entity has:
- a stable `EntityId` unrelated to worksheet row number;
- an `EntityType` from the supported catalogue;
- an optional `LaneId`, `ParentId`, and `StackIndex` where applicable;
- point-based geometry in the shared immutable scene;
- a named `StyleKey`, resolved before rendering;
- an explicit z-layer and deterministic order within that layer;
- clipping and label behaviour defined below;
- ownership metadata beginning `GanttCreator.` in Excel shapes;
- equivalent Excel, editable-export, PowerPoint, and PNG representations unless explicitly excluded.
Renderers consume resolved scene primitives. They may translate primitives to host objects but must not independently move labels, recalculate dates, substitute colours, change line widths, or reorder entities.
## Required worksheet fields
| Field | Applies to | Rule |
| --- | --- | --- |
| `Id` | all data entities | Stable unique text ID; generated once |
| `Type` | all rows | Supported entity type only |
| `Description` | visible entities | User-facing text; blank permitted only where stated |
| `Start` | span/point events | Required by type |

---

## Where to read more

- Canonical source: `docs/07-GANTT-ENTITY-GUIDE.md`
- Always-on rule: `.clinerules/07-GANTT-ENTITY-GUIDE.md`
- Full reference: `./references.md` in this directory
