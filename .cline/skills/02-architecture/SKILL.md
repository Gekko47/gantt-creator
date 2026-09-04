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
  `bin/`, `coverage/`, and `publish/` artifacts.
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

---

## Where to read more

- Canonical source: `docs/02-ARCHITECTURE.md`
- Always-on rule: `.clinerules/02-ARCHITECTURE.md`
- Full reference: `./references.md` in this directory
