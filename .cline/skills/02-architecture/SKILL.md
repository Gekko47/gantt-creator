---
name: 02-ARCHITECTURE
description: Use this skill when the conversation is about the product boundary, the one-visible-sheet + one-VeryHidden-helper-sheet rule, the user worksheet contract (tblGanttData), the VeryHidden config worksheet contract, the scene-first rendering model, the three-renderer equivalence rule, the Ribbon command groups, COM ownership, error handling, or the performance budgets.
---

# Product and architecture — revision 3
> Created 2 September 2026 under a new filename. This revision defines the scope, schema, lifecycle, migration and non-security status of the approved `_GanttCreatorConfig` VeryHidden worksheet. When installed, use the path `docs/02-ARCHITECTURE.md`.
## Product boundary
The add-in creates fast, presentation-quality construction-delay visuals from one visible worksheet. It is a drawing tool backed by tabular schedule events, not a critical-path scheduling engine. It must not imply that it calculates contractual entitlement, CPM logic, or delay causation.
## User worksheet contract
The first supported schema is an Excel Table named `tblGanttData` on the active worksheet. The table is visible and adjacent to the Gantt plot.
Required columns:
| Column | Type | Meaning |
| --- | --- | --- |
| `Id` | text | Stable event identifier; generated once, never row-number based |
| `LaneId` | text | Stable visual-lane identifier shared by events on the same line |
| `StackIndex` | whole number | Non-negative vertical-band order; equal values deliberately share one line |
| `Type` | catalogue text | In-cell dropdown from the single `EntityTypeCatalog` defined by the entity guide |
| `Description` | text | User-facing label |
| `Start` | Excel date or blank | Inclusive start for span events; the single date for a milestone/delineator |
| `Finish` | Excel date or blank | Inclusive finish for span events; not read for milestones/delineators |
| `ParentId` | text or blank | Owning activity for critical/child entities |
| `StyleKey` | catalogue text or blank | Blank uses Type default; otherwise an approved named style |
| `LabelPosition` | catalogue text or blank | Explicit supported position; blank resolves to the Type default/Auto |
| `FillColour` | `#RRGGBB` or blank | Per-row rectangle/diamond/splitter fill override |
| `StrokeColour` | `#RRGGBB` or blank | Per-row outline, hatch, critical, or delineator colour override |
| `Visible` | Boolean or blank | Blank/true renders; false retains the row without rendering |
Optional columns can be added only by an approved schema ADR, initially `SortOrder` and user notes. Core data columns remain visible; property columns may be shown/hidden using the Ribbon but remain on the same worksheet. Do not introduce one column per critical interval. A critical interval is another event record tied to a lane or parent.
Workbook-level configuration uses one worksheet named `_GanttCreatorConfig` with `Visible = xlSheetVeryHidden`, plus workbook-defined names prefixed `GanttCreator.` that point to its validation ranges. Shape ownership uses a deterministic shape name plus tags/alternative text.
## VeryHidden configuration worksheet contract

---

## Where to read more

- Canonical source: `docs/02-ARCHITECTURE.md`
- Always-on rule: `.clinerules/02-ARCHITECTURE.md`
- Full reference: `./references.md` in this directory
