# ADR-0064 — A user-authored `Activity ID` column leads the visible columns

- **Status:** Accepted (owner-approved R5.21, header chosen 2026-11-17)
- **Date:** 2026-11-17
- **Amends:** ADR-0059 D1's ordered column list (which named five visible columns); ADR-0029
  D6 (no migration) is reaffirmed, not changed
- **Related:** ADR-0029 (schema contract and versioning), ADR-0030 (reserved rows and the
  plot anchor), ADR-0059 (visible columns lead the table), R5.18h (persisted column
  visibility), R5.21

## Context

The table exposes the fields the analyst reads and edits, and every one of them is either an
entity fact (`Type`, `Description`, `Start`, `Finish`) or a derived value (`Duration`). There
is no place to record the analyst's own reference for a row — the identifier a programme,
cost report, or client schedule uses. The engine `Id` is not that field: it is an add-in
identity, hidden, locked, and referenced by `ParentId`, scene ownership and outline planning,
so exposing it would invite hand edits the next Refresh would overwrite.

The owner asked for a reference identifier the user types, positioned between `Type` and
`Description`, with the same hide/show control as the other visible columns.

## Decision

1. **The header is `Activity ID`.** It is the owner's chosen display name. It is distinct from
   the engine `Id` in name, position, access class and semantics.
2. **It is the second schema column, immediately after `Type`.** The visible order becomes
   `Type`, `Activity ID`, `Description`, `Start`, `Finish`, `Duration`, then the twelve
   engine columns unchanged.
3. **It is classified `Authoring` and required.** The user types it, it is unlocked so the
   sheet stays authorable under protection, and every v26 table carries the column.
4. **It is free text with no added validation.** A blank is legal data, the value is trimmed,
   and the row is still not an entity if only this cell is filled. A format or duplicate rule
   would be a new validation code and its own work item.
5. **It participates in the panel and the export.** It is measured as a visible column and its
   text is projected into the data panel, so the chart's panel and the worksheet agree.
6. **The schema version goes 25 → 26, and the pinned catalogue hash moves with it.**

## No migration

There is no migration path and none is written. A version-25 workbook lacks the column and
reports the schema-version mismatch; the remedy is Initialise, exactly as every previous
schema bump (ADR-0029 D6). Adding the column to a version-26 table is Initialise's job, and
the structure-repair plan inserts a missing schema column when a v26 workbook is damaged.

## Consequences

- **The plot anchor moves one column right.** `GanttSheetLayout.PlotAnchorColumnIndex` is
  `TableOriginColumnIndex + column count`, so the anchor moves from column `S` to column `T`
  with no code change. A v25 workbook's stored `GanttCreator.PlotAnchor` names the old column
  and the integrity checker reports the disagreement until Initialise or Repair rewrites it.
- **The measured panel gains one column width, so the plot narrows by it.** The plot starts at
  the measured anchor column's left edge plus the live inset; a wider panel therefore moves
  the origin right and shrinks the plot. Hiding `Activity ID` from the ribbon removes it from
  the measured set and restores the previous geometry. At a narrow size preset a very wide
  column can reach the existing `InsufficientPlotWidth` refusal; that is the existing rule
  firing on a real input, not a new failure mode.
- **The positional header repair sees a v25 table as misaligned** at index 1, because
  `Description` now sits where the schema expects `Activity ID`. The version gate reports the
  mismatch first, so the user is sent to Initialise rather than to a header repair.
- **Everything else is name-keyed and needs no change**: the reader map, the row writer, the
  visibility port, the presentation restorer, the validation reporter, the change analyzer,
  the extra-column scan and the Type-column validation.
- **The ribbon row id is `colActivity ID`.** Item ids are `col` + the schema display name, and
  the header contains a space. RibbonX ids are opaque attribute strings and the prefix parser
  is ordinal, so the id is expected to work; a live gate confirms the row renders and toggles.
- **The catalogue hash moves through the schema seed only.** No catalogue row, style or
  setting changes, so the settings table's key set and the config anchor do not move.
