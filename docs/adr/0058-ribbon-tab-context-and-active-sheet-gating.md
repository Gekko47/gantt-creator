# ADR-0058 — Ribbon tab context: active-sheet gating and tab activation

- **Status**: Accepted
- **Date**: 2026-11-12
- **Work item**: R5.16b (roadmap row in [03-ROADMAP.md](../03-ROADMAP.md))
- **Plan**: [ribbon-compact-layout-and-context-plan.md](../../plans/ribbon-compact-layout-and-context-plan.md)

## Context

The add-in's data surface is a real Excel table (`tblGanttData`), so whenever the
selection enters that table Excel shows its own **Table Design** contextual tab and
takes the tab row. The owner asked whether the Gantt Creator tab can be *pinned*
the way Table Design is, whether the Table tab can be unpinned so it stops taking
over on a cell click, and whether the add-in's controls can be disabled when the
Gantt sheet is not the active sheet.

Three facts constrain the answer.

1. **A custom tab cannot be made contextual in RibbonX.** `<contextualTabs>` only
   customises a built-in tab set by `idMso`; it cannot register a new one. Custom
   contextual tabs are an Office JS add-in capability (RibbonApi 1.2 plus a shared
   runtime), which an Excel-DNA COM add-in does not have. *(Microsoft Learn,
   "Create custom contextual tabs in Office Add-ins", retrieved 2026-11-12.)*
2. **Built-in tab visibility is static.** Excel's table tab set can be hidden with
   `<tab idMso="…" visible="false"/>`, which removes native table formatting from
   **every workbook for the whole session**. There is no runtime toggle and no
   supported way to re-order Excel's contextual tab sets.
3. **`IRibbonUI.ActivateTab` exists in Excel-DNA 1.9** alongside `Invalidate`,
   `InvalidateControl` and `InvalidateControlMso` (verified in R1.5 against the
   installed package metadata), and returns `S_FALSE` when the ribbon is collapsed.
   It selects a custom tab; it does not pin it.

The add-in also reads the table from **anywhere in the active workbook**, so
`SheetInitialised` is workbook-wide: today every table button stays live while the
user is looking at an unrelated sheet.

## Decision

**The tab stays visible in every mode.** Hiding the tab breaks first run (no table
exists yet, so the tab would never appear) and makes the ribbon shift as the user
moves between sheets. Group visibility keeps following the three R5.16 modes.

**Table-scoped controls are gated on the active sheet.** A new reader fact records
whether the active worksheet owns `tblGanttData` (or carries the standard Gantt
label when the table is missing). Every table-scoped control multiplies its existing
gate by that fact.

**Recovery commands are exempt.** Initialise sheet, Repair, Validate, Diagnostics
and Open log are never gated on the active sheet. This is ADR-0056's own rule: a
recovery action must never be hidden behind the state it recovers from.

**The Gantt Creator tab is activated on sheet entry and after every completed
refresh.** `ActivateTab` is called (queued through the macro queue, best-effort,
skipped when the ribbon is collapsed) when the Gantt sheet becomes active, and at
the end of both refresh paths — `RunForExcel` and `RunSilentForExcel` — because the
refresh's own selection restore is the moment Excel pulls up Table Design. It is
**never** called on a selection change.

**Excel's Table Design tab is not hidden**, and the Quick Access Toolbar is not
modified.

## Consequences

- The takeover by Table Design is accepted as host behaviour. The user is returned
  to the Gantt Creator tab at the two moments the add-in owns: entering the sheet,
  and finishing a refresh.
- A user who deliberately opens another tab during a refresh is brought back to the
  Gantt Creator tab. That is the owner's ruling (2026-11-12) and is recorded here so
  a later reviewer reads it as a choice, not a defect. A preference to switch it off
  would be a workbook-settings change and is not part of this decision.
- `IExcelApplicationAdapter` gains `SubscribeSheetActivated`; the attach/detach
  discipline matches the existing workbook-state subscription.
- `GanttTableReadOutcome` gains the active-sheet fact, computed by the same sheet
  search that already locates the table.
- A selection change never activates a tab, so the user can work on another ribbon
  tab while staying on the Gantt sheet.

## Alternatives considered

- **Create a custom contextual tab.** Not available to a RibbonX add-in; rejected as
  impossible rather than difficult.
- **Hide Excel's Table Design tab statically.** Rejected: it removes native table
  formatting from every workbook and is not confined to this document contract.
- **Call `ActivateTab` on every selection change.** Rejected: it would yank the
  ribbon away from the user mid-edit, which is worse than the takeover.
- **Hide the Gantt Creator tab when the Gantt sheet is not active.** Rejected: it
  breaks first run and makes the ribbon shift on every sheet switch.
- **Put the key actions on the Quick Access Toolbar.** Rejected by the owner
  (2026-11-12): the QAT is the user's own surface and RibbonX additions persist
  across sessions.
