---
name: 02-ARCHITECTURE
description: Use this skill when the conversation is about the product boundary, the one-visible-sheet + one-VeryHidden-helper-sheet rule, the user worksheet contract (tblGanttData), the VeryHidden config worksheet contract, the scene-first rendering model, the three-renderer equivalence rule, the Ribbon command groups, COM ownership, error handling, or the performance budgets.
---

# Product and architecture — revision 3
> Created 2 September 2026 under a new filename. This revision defines the scope, schema, lifecycle, migration and non-security status of the approved `_GanttCreatorConfig` VeryHidden worksheet. When installed, use the path `docs/02-ARCHITECTURE.md`.
## Product boundary
The add-in creates fast, presentation-quality construction-delay visuals from one visible worksheet. It is a drawing tool backed by tabular schedule events, not a critical-path scheduling engine. It must not imply that it calculates contractual entitlement, CPM logic, or delay causation.
## Build pipeline artifact contract
The build pipeline produces artifacts in a deterministic order. Tests may
only consume artifacts that a verify-script step guarantees:
- `verify-quick.ps1` and `verify.ps1` are the authoritative producers of
  `bin/` and `publish/` artifacts; `verify.ps1` additionally produces
  `coverage/`.
- A test that reads from `bin/` or `publish/` must be traceable to a
  specific step in those scripts. The traceability lives in the work item.
- A test must not assume prior build state. It either creates its own
  inputs, or depends on a verify-script step whose existence is asserted
  by the work item.
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
The workbook contains exactly:
- one visible user worksheet containing `tblGanttData` and the live Gantt; and
- one add-in-managed `_GanttCreatorConfig` worksheet with `xlSheetVeryHidden` visibility.
The configuration worksheet contains versioned tables/ranges only:
| Item | Purpose |
| --- | --- |
| `tblGanttTypes` | Materialised built-in Type values, display names, entity kind, date mode, default style and capability flags |
| `tblGanttStyles` | Workbook style presets and user-approved custom named styles |
| `tblGanttMetrics` | Rectangle heights, milestone size, gaps, line widths and other named geometry tokens |
| `tblGanttLabelPositions` | Valid label choices grouped by entity capability |
| `tblGanttConfig` | Schema version, catalogue hash, workbook ID and add-in version last used |
| `GanttCreator.TypeOptions` | Workbook name referring to the active Type display-name range used by data validation |
Built-in Type definitions remain code-owned and are materialised deterministically. The helper worksheet is the workbook-scoped validation/configuration representation, not a second competing type catalogue.
The helper must never contain activity/event rows, descriptions, schedule dates, scene primitives, rendered shapes, formulas that determine chart geometry, logs, or export staging. It is excluded from all exports and normal navigation. `xlSheetVeryHidden` and worksheet protection prevent accidental edits but are not treated as security.
On initialise/open/Refresh, validate sheet name, visibility, schema version, table headers, defined names and catalogue hash. Rebuild missing built-in catalogue rows safely. Preserve valid user style presets during a schema migration. If repair could lose custom configuration, stop and request confirmation rather than silently recreating the sheet.
If a user copies only the visible worksheet into another workbook, the add-in detects the missing configuration sheet and offers to initialise it from code defaults. It does not refuse to read the visible schedule data or copy schedule content into the helper.
## Domain model
Use immutable Core types. Suggested concepts, not mandated class names:
- `GanttDocument`: validated data, plot range, lanes, delineators, and style theme.
- `Lane`: ordered visual row with one or more events.
- `SpanEvent`: start/finish event such as planned, actual, baseline, procurement, or critical.
- `PointEvent`: milestone, date label, or delineator.
- `EventStyle`: explicit fill, stroke, thickness, hatch, marker, font, and label placement.
- `TimeScale`: maps `DateOnly` values to plot-space point coordinates.
- `Scene`: immutable ordered primitives in points.
- `ScenePrimitive`: rectangle, line, polygon, text, and group metadata.
- `ExportSize`: width, calculated height, pixel dimensions, and DPI.
Validation returns all actionable issues in deterministic table/row order. It must distinguish blocking errors from warnings. Do not throw for routine bad user input.
Key rules:
- IDs are stable and unique.
- span start is not after finish;
- lane and stack ordering is deterministic;
- milestones and delineators require one date;
- Type values come only from the central entity catalogue; unknown pasted values are blocking errors;
- row colour and label-position overrides are validated against the selected Type's capabilities;
- text is trimmed but not silently rewritten;
- date parsing uses the Excel cell value and workbook date system, not locale-dependent display text;
- events outside the selected plot range are clipped or warned according to an explicit policy;
- no shape is emitted with NaN, Infinity, negative width, or negative height.
## Scene-first rendering
Every renderer consumes the same immutable scene. Worksheet reading and layout are not repeated independently by each renderer.
```mermaid
flowchart TD
    A["One worksheet table"] --> B["Validate domain"]

---

## Where to read more

- Canonical source: `docs/02-ARCHITECTURE.md`
- Always-on rule: `.clinerules/02-ARCHITECTURE.md`
- Full reference: `./references.md` in this directory
