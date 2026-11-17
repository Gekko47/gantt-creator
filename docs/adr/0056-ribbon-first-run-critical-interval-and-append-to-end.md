# ADR-0056 — Ribbon first-run surface, Critical Interval attachment, and the append-to-end insert

- **Status**: Accepted
- **Date**: 2026-11-10
- **Work item**: [R5.16](../work-items/R5.16-ribbon-first-run-and-data-group.md)
- **Plan**: [ribbon-ux-redesign-plan.md](../../plans/ribbon-ux-redesign-plan.md)

## Context

Three problems sit in the same surface and were ruled on together.

**First-run discoverability and recovery.** The tab currently shows every
action on a workbook that has never been initialised, so a new user sees a wall
of controls that cannot work. Hiding everything except Initialise fixes that,
but a workbook that carries `_GanttCreatorConfig` and has lost its table returns
`TableMissing`: a strict two-state rule would hide Repair while Initialise
refuses with "use Repair configuration", which leaves no way out from the
ribbon.

**Data-group sprawl.** Seven flat controls plus an Add child dropdown occupy the
Data group. The owner asked for five large controls, with the secondary actions
behind split-button arrows, and for Add child to disappear because Demote
already exists.

**Insertion semantics.** R5.13 assigns a parent from the insertion position: a
row landing inside a group's contiguous child run joins that group. That rule
is position-only and never consulted the hierarchy catalogue, so it could write
a parent the model forbids. Separately, a Critical Interval is child-only and
the ribbon offered no way to add one from a row that was not already its
parent.

## Decision

**Ribbon modes.** A workbook is in exactly one of three modes, computed from
workbook presence, table readability, and configuration presence:

- `FirstRun` — no configuration: a large **Initialise sheet** button in **Get
  started**, plus **Diagnostics**;
- `NeedsRepair` — configuration present but the table missing or unreadable:
  the full surface, Initialise hidden, Repair enabled, table-dependent controls
  disabled;
- `Ready` — the table is readable: the full surface, Initialise hidden.

**Amendment — 2026-11-11 (R5.17b–d).** `Ready` now also requires the
configuration sheet. The mode is computed from workbook presence,
`ConfigurationPresent`, and table readability:

- `Ready` — table readable **and** configuration present;
- `NeedsRepair` — configuration present with the table missing or unreadable,
  configuration missing with a readable table, **or** both containers missing
  when a standard-labelled Gantt worksheet still exists;
- `FirstRun` — no workbook, or no configuration and no standard-labelled Gantt
  worksheet.

The exception the original rule missed is the both-missing state: it now keeps
the user on `Fix`, because sending it to `FirstRun` let **Initialise** create
`Gantt Data (2)` beside the existing empty Gantt sheet. A present but corrupt
configuration still reports `Ready`, so Fix and Refresh surface the catalogue
error instead of the ribbon hiding the surface.

Group visibility is fail-closed; control enablement stays fail-open. Diagnostics
is visible in every mode, which is a deliberate exception to "only Initialise
on a first run".

**Data group.** **Fix** (primary Repair configuration, arrow Validate), **Add
activity**, **Add milestone**, **Additional elements**, and **Remove**, all
large. Menu labels are projected from `EntityTypeCatalog`; the union of the
three entity menus equals the catalogue partition. **Add child** is removed;
**Demote to child** is the child-making command.

**Critical Interval.** A Critical Interval is always a child. It attaches to the
nearest eligible parent — a row that may own children, is not itself a child,
and has fewer than seven children — searching upward from the selection first
and then downward, and it is placed after that parent's last child. The insert
refuses only when the table contains no eligible parent, and the ribbon item is
disabled in exactly that case. The command never creates a parent as a side
effect.

**Other inserts.** R5.13 is **retained**: a child-capable row inserted inside a
contiguous child run joins that group. The candidate parent is now validated
through `AddChildPlanner` (child-capable type, depth, seven-child cap); a type
that cannot be a child, and any candidate the planner rejects, is appended as
the last body row instead. The end of the table is always a valid position and
cannot split a run.

## Consequences

- A new user sees one action, and a damaged workbook still has a reachable
  repair path.
- The configuration sheet gains one read on the `TableMissing` path
  (`ConfigurationPresent`); no schema change.
- `SheetInitialised` now requires a successful table read. Every existing
  consumer also requires an active workbook, so the visible behaviour is
  unchanged apart from a correct false where it was previously true with no
  workbook.
- The hierarchy catalogue becomes the authority for the insert as well as the
  planner: `AddChildPlanner.IsEligibleParent` and `Evaluate` are shared by the
  ribbon gate, the inserter, and `Plan`.
- A Critical Interval added from a milestone or a structural row moves to its
  parent's child run rather than staying at the click point. That is the price
  of a groupable, depth-correct row.
- Spacer and splitter additions inside a group move to the end of the table.
  They cannot be children, so the alternative is a split child run.
- Removing Add child deletes `AddChildCommand`, `IAddChildMutator`, and
  `ExcelAddChildMutator`. The planner stays; deleting it would weaken
  validation.

## Alternatives considered

- **Literal two-state visibility.** Rejected: it hides Repair on a workbook that
  needs it, which is a recovery dead end.
- **Put Diagnostics in the first-run group.** Not required: a separate group
  satisfies the owner ruling and keeps "Get started contains one control" true.
- **Keep Add child alongside the new menus.** Rejected by owner ruling 3; it
  duplicates Demote and keeps a second child path with its own command chain.
- **Insert the Critical Interval at the click point and attach the nearest
  parent.** Rejected: a child above, or rows away from, its parent cannot be
  grouped, so the outline would be invalid until a repair pass.
- **Create an As-Planned Activity when no parent is selected.** Rejected by the
  owner (2026-11-10): adding a Critical Interval must not create schedule data.
- **Extend R5.13's neighbour rule to Critical Interval.** Rejected: the
  neighbour rule assumes a parent is adjacent, which is exactly what is not
  guaranteed when the user selects a leaf or a structural row.
