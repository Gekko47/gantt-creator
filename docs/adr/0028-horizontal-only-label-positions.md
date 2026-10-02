# ADR-0028 — Horizontal-only label positions; the existing truncation is retained

- **Status:** Accepted
- **Date:** 2026-09-29
- **Relates to:** entity guide §12, §20, §22; ADR-0015
- **Decided by:** product owner, 2026-09-29
- **Implements:** roadmap row R4.7G; register D4 and D6
- **Amends:** ADR-0015 (D1, D2 only). **ADR-0015 is not superseded.**

## Context

REV5 §15 and REV6 §15 propose removing `Above` and `Below` from the label
positions, on the grounds that with fixed row heights (ADR-0026) a label placed
above or below a bar cannot fit without either overlapping the adjacent row or
being clipped at the plot edge.

`GanttLabelPosition` carries `Above = 5` and `Below = 6`. Entity guide §616-617
defines their geometry, and `EntityTypeCatalog` includes them in the permitted
sets.

The product owner ruled on 2026-09-29: *"above and below are now to be removed,
only right, left and centre are to be used."* In the same ruling the cascade was
specified: *"Preference order is Right, then auto based on largest space
available that fits the text string. and then none for those that have none."*

**An earlier draft of the pre-R4.9 plan misread that second sentence** and
proposed replacing ADR-0015's widest-gap truncation fallback with a hard
suppression. The owner corrected it: *"there is a function already for this, it
will truncate if there is no space, nothing new needs to be built."* That reading
is **withdrawn**.

## What already exists

`LabelPlanner` already implements the whole cascade, and has since R3.6:

| Element | Location |
| --- | --- |
| `TruncatedToFitCode = "LabelTruncatedToFit"` | line 130 |
| `WasTruncated` on the result | lines 65, 71 |
| `_spanAutoOrder`, **already beginning at `Right`** | lines 139-141 |
| `FreeRight` / `FreeLeft` occupancy measurement | lines 658, 685 |

So the cascade the owner described is the cascade that is built, except that
`Above`/`Below` still participate in it.

## Decision

- **D1 — `Above` and `Below` are removed from `GanttLabelPosition`**, and from
  every permitted-position set in `EntityTypeCatalog`, from `LabelPlanner`'s
  milestone order, from entity guide §616-617, and from
  `08-TEST-CHECKLIST.md` §C. **All five sites change in one commit**; a partial
  change leaves a position that parses but cannot be placed, or a capability set
  advertising a removed position.
- **D2 — The retained horizontal set is `None`, `Auto`, `Left`, `Right`,
  `Inside`.** `Inside` is surfaced in the Ribbon as **Centre**, which is the
  owner's term for it. **CORRECTED 2026-09-30: the retained members were
  renumbered contiguously** — `Inside` moved 6 → 4 and the delineator corners and
  splitter positions each shifted down by two as `Above`/`Below` vacated 5/6. This
  ADR originally claimed *"the numeric values of the retained members are
  unchanged, so no stored ordinal is invalidated"*, which is **false as an
  arithmetic claim**. It is safe in effect, but for a different reason than stated:
  the enum is persisted **by name**
  (`GanttTypeCatalogueRow.AllowedLabelPositions` serialises
  `position.ToString()`), so **no workbook stores a number** and renumbering cannot
  orphan a stored value. Contiguity was chosen over a retained 5/6 gap because a
  gap requires a marker member nothing consumes, which CA1700 correctly rejects.
- **D3 — The delineator corner positions are retained**: `TopLeft`, `TopRight`,
  `BottomLeft`, `BottomRight`. They are not above/below positions; they are the
  four corners of a full-height line. The landed R3.10 `DelineatorBuilder`
  depends on them, and removing them would break that row. **Renumbered 5-8** as a
  consequence of D2's correction above.
- **D4 — The splitter positions are retained**: `DataPanelLeft`, `PlotCentre`,
  `Both`, which the splitter band and the landed R3.x builders depend on.
  **Renumbered 9-11** as a consequence of D2's correction above.
- **D5 — ADR-0015 D4, D5 and D6 stand unchanged.** The widest-gap truncation
  fallback, the single-character `…`, and the nothing-to-place outcome are
  **reused as built**. No new label machinery is written. This ADR amends ADR-0015
  on *which positions exist* and nothing else.
- **D6 — The stated preference order is recorded as the resolved order:**
  `Right` is the preferred position; `Auto` selects the largest available space
  that fits the text; the existing truncation handles a space that is too small;
  `None` applies to entity types that have no label. `CriticalInterval` and
  `Spacer` already resolve to `None` at `GanttCatalogues.cs:348` and are
  unaffected.
- **D7 — Label collisions are never resolved by increasing row height.** A
  collision is resolved by the existing cascade, or reported. This follows
  directly from ADR-0026 and is stated here because ADR-0015 predates it.

## Consequences

- Two label positions disappear from the user-facing set, which is a **schema
  contract change** (ADR-0007 D3 makes the permitted-position set part of the
  workbook schema) and joins ADR-0029's version bump.
- The label planner's cascade is barely touched. R4.7G is predominantly a
  catalogue-and-contract change, not an algorithm change, which is why this ADR
  is narrow where the pre-R4.9 plan originally expected a supersession.
- Any stored `LabelPosition` value of `Above`/`Below` in an existing workbook
  becomes unresolvable. **There are no active users, so no migration is written**;
  validation reports it rather than silently coercing it.
- The `Auto` order's first element was already `Right`, so the owner's stated
  preference was already partly true in the landed code. D6 records the resolved
  intent rather than changing behaviour.

## Alternatives considered

- **Remove the delineator corners too, for a strictly horizontal set.** Rejected:
  the corners are the only positions that can place a label on a full-height
  line, and R3.10's landed `DelineatorBuilder` and its `TopRight → BottomLeft`
  cascade depend on them. Removing them breaks a completed row.
- **Rename `Inside` to `Centre` in the enum** rather than only in the UI.
  Rejected: the enum is the durable, schema-bound contract and `Inside` is its
  existing name; a rename would invalidate stored values for no behavioural
  gain. The UI label is the user's term, so the UI carries it.
- **Reorder `_spanAutoOrder` to put largest-space first.** Rejected: ADR-0015
  D1 records that largest-gap-as-primary was considered and rejected, because it
  would make the owner's stated `Right` preference unobservable. `Right` is tried
  first, and the cascade still measures.
- **Suppress rather than truncate when nothing fits.** Rejected — this was the
  withdrawn reading. The existing truncation is the owner's confirmed behaviour.
