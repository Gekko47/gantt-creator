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
| `Finish` | span events | Required; must not precede Start |
| `LaneId` | lane-bound entities | Events sharing a visual line use the same value |
| `StackIndex` | lane-bound entities | Non-negative vertical-band order; equal values deliberately share the same line |
| `ParentId` | critical/child events | Stable ID of the owning activity when required |
| `LabelPosition` | labelled entities | Supported position; blank resolves to the Type default or `Auto` |
| `StyleKey` | styleable entities | Blank uses type default; otherwise an approved named style |
| `FillColour` | filled rectangles/diamonds/splitters | Blank uses resolved style; override is `#RRGGBB` |
| `StrokeColour` | lines/outlines | Blank uses resolved style; override is `#RRGGBB` |
| `Visible` | optional | Blank/true shows; false suppresses entity and its label |
| `SortOrder` | optional | Stable explicit order before ID tie-break |
The first release does not add raw X, Y, width, or height columns. Workbook style controls set shared tokens; named `StyleKey` values provide controlled exceptions. `FillColour`, `StrokeColour`, and `LabelPosition` are row-level overrides because the product explicitly requires individual entity editing.
Property columns may be shown or hidden through the Ribbon's **Show Properties** control, but they remain columns on the visible user worksheet. They are never moved to `_GanttCreatorConfig`.
## Type catalogue and worksheet dropdown
The `Type` column uses an Excel in-cell dropdown backed by the workbook name `GanttCreator.TypeOptions`, which refers to the materialised built-in catalogue on `_GanttCreatorConfig`. The code-owned `EntityTypeCatalog` remains authoritative. Initialisation and migration write it deterministically to `tblGanttTypes`; the domain parser, worksheet validation, Ribbon controls, entity guide, and tests use the same definitions. No renderer maintains its own type list.
Initial catalogue:
| Dropdown value | Meaning | Dates read | Default style | Fill/colour capability | Allowed label positions |
| --- | --- | --- | --- | --- | --- |
| `Splitter` | Section header | none | Splitter | Fill | DataPanelLeft, PlotCentre, Both, None |
| `Spacer` | Blank vertical space | none | Spacer | none | None |
| `As-Built Activity` | As-built activity | Start + Finish | AsBuiltActivity | Fill + outline | Auto, Left, Right, Inside, Above, Below, None |
| `As-Planned Activity` | As-planned activity | Start + Finish | AsPlannedActivity | Fill + outline | Auto, Left, Right, Inside, Above, Below, None |
| `Baseline Activity` | Baseline activity | Start + Finish | BaselineActivity | Fill + outline | Auto, Left, Right, Inside, Above, Below, None |
| `Critical Interval` | Critical child interval | Start + Finish | CriticalInterval | Stroke | None |
| `Delay Event` | Explicit delay span | Start + Finish | DelayEvent | Fill + outline | Auto, Left, Right, Inside, Above, Below, None |
| `As-Built Procurement` | As-built procurement | Start + Finish | AsBuiltProcurement | Hatch + outline | Auto, Left, Right, Inside, Above, Below, None |
| `As-Planned Procurement` | As-planned procurement | Start + Finish | AsPlannedProcurement | Hatch + outline | Auto, Left, Right, Inside, Above, Below, None |
| `Baseline Procurement` | Baseline procurement | Start + Finish | BaselineProcurement | Hatch + outline | Auto, Left, Right, Inside, Above, Below, None |
| `Custom Activity` | Named custom span | Start + Finish | required `StyleKey` | style-defined | style-defined subset |
| `As-Built Milestone` | As-built milestone | Start only | AsBuiltMilestone | Fill + outline | Auto, Left, Right, Above, Below, None |
| `As-Planned Milestone` | As-planned milestone | Start only | AsPlannedMilestone | Fill + outline | Auto, Left, Right, Above, Below, None |
| `Baseline Milestone` | Baseline milestone | Start only | BaselineMilestone | Fill + outline | Auto, Left, Right, Above, Below, None |
| `Critical Milestone` | Critical milestone | Start only | CriticalMilestone | Fill + outline | Auto, Left, Right, Above, Below, None |
| `Delineator` | Full-height vertical date line | Start only | DefaultDelineator | Stroke | Auto, TopLeft, TopRight, BottomLeft, BottomRight, None |
Full display names are used because the VeryHidden range removes the direct-list length limitation. These exact values form part of the workbook schema; localisation or renaming requires a schema migration.
Dropdown implementation requirements:
- Apply validation to every current and newly added table body cell in the `Type` column.
- Set validation formula to the workbook-defined name `=GanttCreator.TypeOptions`; do not construct an inline comma-separated list.
- Keep the named range limited to active display-name cells in `tblGanttTypes`, with no blank tail cells.
- Maintain the single approved `_GanttCreatorConfig` sheet; do not create additional list sheets or duplicate catalogues.
- Validate the catalogue hash and named range before applying validation or parsing Type values.
- Unknown pasted text is a blocking validation error; do not guess the closest type.
- Changing Type does not delete dates or overrides. On Refresh, incompatible fields are reported and ignored only where the entity contract explicitly says so.
`Start Label` and `Finish Label` are label roles owned by a span event, not initial Type dropdown entries. Dependency arrows, progress bars, and current-date lines are also excluded until approved.
### VeryHidden catalogue and style rules
`_GanttCreatorConfig` materialises the built-in catalogue and stores workbook style/metric presets so data validation and user customisation remain portable and offline.
- `tblGanttTypes` is regenerated only from the code-owned catalogue and is not user-editable through ordinary UI.
- `tblGanttStyles` and `tblGanttMetrics` are changed through approved Ribbon dialogs; the add-in writes and validates the underlying rows.
- Built-in style keys cannot be deleted. A user may create a named custom style with a unique key through the style dialog.
- Per-row `FillColour`, `StrokeColour`, and `LabelPosition` remain on the visible event row and override the resolved helper-sheet style.
- The helper sheet contains no schedule/event data and is never used as an export or rendering surface.
- Sheet protection and `xlSheetVeryHidden` reduce accidental edits but provide no confidentiality or tamper-security guarantee.
- A catalogue/style schema version and hash are checked on initialise and Refresh. Unsafe repair requires user confirmation.
## Selected-row property editing
Per-entity Ribbon properties operate only when selection resolves to one expanded, visible entity row in `tblGanttData`.
A valid `SingleEntitySelection` requires:

---

## Where to read more

- Canonical source: `docs/07-GANTT-ENTITY-GUIDE.md`
- Always-on rule: `.clinerules/07-GANTT-ENTITY-GUIDE.md`
- Full reference: `./references.md` in this directory
