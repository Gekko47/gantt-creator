# Implementation plan — the Type-change transition

> **Status**: Implemented 2026-11-27 (code, tests, docs); the Office host gate is owed.
> Revised after a line-by-line review against the tree, then implemented.
> **Work item**: propose `R5.23-type-change-transition` (next free R5 row); confirm at
> implementation start.
> **Diagnosis**: the session review of 2026-11-27 — see "Root causes" below.

## Outcome

A Type edit produces a row whose type-dependent cells are valid for the NEW type on
every path, including when the render is suspended, and the ribbon, validator and scene
all read the same healed row. The 17×17 ordered matrix is a test, not a hope.

## Root causes this closes

1. `StyleKey` is never re-evaluated on a Type change: `ExcelTypeChangeNormaliser` only
   fills a blank key, `ValidateStyleKey` only checks that the key resolves, and
   `SceneBuilder` then applies a span style's default label position to a delineator.
   Result: `Custom Activity → Delineator` keeps `MyStyle`; the ribbon offers
   `span positions ∩ corners = {Auto, None}` and the chart's description label follows
   the span style's default position.
2. The date-label column carries two vocabularies, but the heal only understands the
   which-date one: a corner left on a span blocks with `UnknownDateLabelPosition`, and
   `Auto`/`None` are members of both enums, so their meaning silently flips.
3. The whole heal runs only inside a render pass, so a Type edit while
   `AutoRefreshEnabled` is off or the chart is unlocked is never healed — and the
   always-on ribbon now shows exactly that unhealed row.

## Rulings (owner-confirmed; record in docs/DECISIONS.md)

- **R1 — Write on the change path.** A Type edit heals its own row immediately, before
  the render decision and irrespective of the switch/lock. It is a repair of the user's
  edit, not a chart build: no scene, no shape, no plot.
- **R2 — One authority.** The transition rules live in Core and consume only
  `EntityTypeCatalog`, `EntityHierarchyCatalog` and the style registry's own
  projections. The Excel adapter writes what Core returns; the validator and the ribbon
  use the same projections.
- **R3 — Vocabulary boundary clears.** When a Type change crosses the which-date ↔
  placement boundary (delineator vs every other dated type), `DateLabelPosition` is
  blanked whatever the token, including `Auto` and `None`. Same-vocabulary transitions
  keep a value only when it parses and is allowed by the new type.
- **R4 — Lane/stack are retained for a Delineator only.** Guide §24 says a delineator
  keeps a populated `Finish`, `LaneId` and `StackIndex` with a non-blocking warning, so
  the normaliser stops clearing lane/stack for that type. `Finish` is already retained
  for every StartOnly type. Splitter/Spacer keep today's clearing: the guide states no
  retention rule for them and the validator already reports the field as irrelevant;
  the pinned Spacer fact stays green.
- **R5 — Style compatibility is capability-based.** A style key is usable by a type
  when it **resolves in the registry** AND its family matches the type's, where the
  family is inferred from the style's non-shared positions: corners mean a delineator
  style, side positions a span/milestone style, plot positions a splitter style. A style
  offering only the shared `Auto`/`None` is neutral and usable anywhere, because `Auto`
  resolves to a valid placement for every builder. The resolution must be explicit:
  `GetAllowedLabelPositions` falls back to the type's default preset for an unknown key,
  so it cannot be the resolution test (it would call every unknown key usable). This
  needs no new schema field and treats a user-authored style with the same rule as a
  built-in.
- **R6 — Unparseable text is not deleted.** A token that parses under neither
  vocabulary is left for the validator; only type-dependent values are cleared.
- **R7 — No automatic re-parenting.** Converting a parent to a type that cannot own
  children leaves the children in place and the validator reports it; a prompt or
  re-parent flow is a separate decision.

## Per-column rules (the matrix source)

| Column | Rule |
| --- | --- |
| `Type` | The edit itself. `EntityTypeCatalog.TryParse` decides the new definition; unknown/blank skips the row. |
| `StyleKey` | If the new type **requires** a key: keep only when `IsStyleKeyUsable` passes (resolves, and `Intersect(newDef.AllowedLabelPositions, style.AllowedLabelPositions)` contains a position other than `Auto`/`None`); otherwise write `newDef.DefaultStyleKey`. If the new type does **not** require one: leave the cell untouched — `StyleKeyWritePlanner` already rewrites it to the type default on every pass, and doing it here would duplicate that rule and break the pinned "all values legal" fact. |
| `LabelPosition` | Clear unless it parses and `EntityTypeCatalog.IsLabelPositionAllowed(newType, position, finalStyleKey, styles)`. Uses the key AFTER the style rule, so the answer matches the validator and the ribbon. |
| `DateLabelPosition` | (a) New type has neither vocabulary (`GetAllowedDateLabelPositions` and `GetAllowedDateLabelPlacements` both empty): clear. (b) Old type known and its vocabulary family differs from the new one: clear (R3). (c) Otherwise clear unless the token parses in the new type's own vocabulary and is allowed. |
| `FillColour` | Clear unless the new type's `ColourCapability` has `Fill`. |
| `StrokeColour` | Clear unless it has `Stroke`. |
| `ParentId` | Keep when `EntityHierarchyCatalog.ParentIdIsRelevant(newType)`; else clear. |
| `SiblingOrder` | Keep when `EntityHierarchyCatalog.MayBeChild(newType)`; else clear. |
| `LaneId`, `StackIndex` | Retain for a Delineator (R4); keep today's clearing for every other non-lane type. The validator reports the non-blocking `NotUsedByType` warning. |
| `Finish` | Retain (R4). The validator reports the StartOnly warning. |
| `Duration` | Engine-owned by the pipeline; untouched here. |
| `Id`, `ActivityId`, `Visible`, `SortOrder` | Untouched. |

## Core design

New `src/GanttCreator.Core/EntityTypeTransition.cs` (pure, no Office):

```csharp
public sealed record TypeTransitionRow(
    string? StyleKey, string? LabelPosition, string? DateLabelPosition,
    string? FillColour, string? StrokeColour, string? ParentId,
    string? LaneId, string? StackIndex, string? SiblingOrder);

public sealed record TypeTransitionChange(string Column, string? NewValue, string Reason);

public static class EntityTypeTransition
{
    /// oldType is null when the caller only knows the new type (the every-pass sweep).
    public static IReadOnlyList<TypeTransitionChange> Apply(
        TypeTransitionRow row,
        GanttEntityType newType,
        GanttEntityType? oldType,
        GanttStyleRegistry? styles);
}
```

- `oldType` is the only information the sheet cannot provide; the change-path caller
  supplies it from the pre-edit cached rows. With `oldType == null` the function is
  idempotent and only applies the same-type rules, so the every-pass sweep remains safe.
- No column name is invented: the change record uses the schema names already in
  `GanttTableSchema`.
- A public helper `EntityTypeCatalog.IsStyleKeyUsable(type, styleKey, styles)` is added
  for R5. It first calls `styles.TryGet(styleKey, out definition)`; only a resolved
  definition is judged, because `GetAllowedLabelPositions` silently falls back to the
  type default for an unknown key. It is shared by the transition, the validator and
  `StyleKeyWritePlanner`, so the predicate has one implementation and three triggers.
- `StyleKeyWritePlanner` is the pass-time owner for the key and already rewrites a
  non-requiring type's cell to the type default (`Plan`, the `!RequiresStyleKey` branch).
  The transition does not touch that case; the two only share the predicate.

## Office and AddIn wiring

1. **Adapter rewrite.** `ExcelTypeChangeNormaliser` keeps its one-read column map and
   protection guard; its per-column bodies are replaced by one `EntityTypeTransition.Apply`
   call and one write per changed cell. `ITypeChangeNormaliser` is in Core and its
   `Normalise()` gains `GanttStyleRegistry? styles`. Callers/implementations to update in
   the same commit: `ExcelTypeChangeNormaliser`, the orchestrator's step-5b call, and
   `FakeTypeChangeNormaliser` in [`RefreshFakes.cs`](../tests/GanttCreator.Office.ContractTests/RefreshFakes.cs:844).
   The orchestrator passes the registry it already read at step 2 (`config.Styles`), so
   no new COM read.
2. **Change-path command.** New `TypeChangeHealCommand.RunForExcel()`:
   - resolve the application; no-op when absent;
   - open the shared protection window (release → write → re-protect);
   - run an `IApplicationStateScope` suppression so the heal's own writes cannot
     re-trigger the scheduler;
   - read the table once, match rows to `RibbonStateService`'s cached pre-edit rows by
     `Id` (never by body index — a sort moves rows) to obtain `oldType`; add one internal
     accessor for that cache. The cache is refreshed only by a pass or a projection, never
     by a selection change, so it is the pre-edit baseline as long as the heal runs before
     the projection in the same macro. A row with no cached match (a new insert, or a row
     an unprojected workbook has never shown) falls back to `oldType = null`;
   - read the style registry, call the Core transition for every row, and apply the
     changes **per row after the whole row is computed**, so a mid-row failure cannot
     leave half a row healed;
   - `RibbonStateService.Instance.RefreshAfterTableChange()` and one technical record
     `gantt.types heal rows=<n> cells=<k>`;
   - never throws; silent apart from the status bar.
3. **Change detection.** `ISheetChangeAnalyzer` gains
   `bool IsTypeColumnChange(string? changedAddress)`, implemented by the same
   address→column mapping as `ShouldAutoRefresh`, restricted to the `Type` column.
4. **Scheduler.** `GanttRefreshScheduler` keeps a `_burstTouchesType` flag, set when an
   accepted change is a Type-column change. One macro per burst, in this order:
   `typeHeal()` when the flag is set → the render when allowed, else the projection.
   The heal runs even when suspended. The injected `typeHeal` is
   `TypeChangeHealCommand.RunForExcel` in production and a counter in tests, so the
   ordering and the "heal while suspended" fact are both assertable.
5. **Validator + planner (R5).** `ValidateStyleKey` adds a Warning
   (`StyleKeyNotAllowedForType`) when the key resolves but `IsStyleKeyUsable` fails, and
   `StyleKeyWritePlanner` heals that case to the type default in its `RequiresStyleKey`
   branch, which today keeps any resolvable key. `GanttEvent` already carries `Type` and
   `StyleKey`, so no Core event change is needed. Positive tests for both the warning and
   the plan. The non-requiring branch is already correct and stays untouched.
6. **Ribbon honesty.** `CaptureDescriptionLabelContext`, `CaptureDateLabelContext` and
   `CaptureDateLabelPlacementContext` return `selected = null` when the stored token is
   not in the offered set (the custom-style capture already does this). A stale or
   hand-typed token then reads as `Default`/`Auto` instead of an out-of-list selection.

## Test plan

1. **Core matrix, `EntityTypeTransitionTests`** — 17×17 ordered pairs. Seed a valid row
   for the old type: `Start` for a dated type, a non-blank description (a blank one is
   only legal for a delineator with `None`), style key only when the type requires one,
   first renderable label position, an old-vocabulary date label, fill/stroke where
   capable, parent when child-capable, lane/stack for a span/milestone, sibling order
   where child-capable. Add two extra seeds for the cross-vocabulary cases (a corner on a
   delineator, a which-date on a span). Use the real registry built from
   `GanttCatalogues.StylePresets`, not an empty one, so the style rule is exercised. For
   each pair assert:
   - the healed row passes `GanttRowValidator` with no `LabelNotAllowedForType`,
     `UnknownLabelPosition`, `DateLabelNotAllowedForType` or `UnknownDateLabelPosition`;
   - `GetAllowedLabelPositions(newType, healed.StyleKey, styles)` contains the healed
     `LabelPosition` (or it is blank);
   - the healed `DateLabelPosition` is blank or legal under the new type's own
     vocabulary, and is blank whenever the vocabulary family changed;
   - `StyleKey` is blank, the type default, or a key `IsStyleKeyUsable` accepts.
2. **Core rule facts** — one per row of the rule table (style reset, name-only style,
   corner→span, which-date→delineator, `None`/`Auto` boundary, lane retention, finish
   retention, unparseable left alone, dateless type clears the date label).
3. **Adapter** — the existing `ExcelTypeChangeNormaliserTests` keep their seams; add the
   Delineator lane-retention fact and the styles-parameter wiring fact.
4. **Scheduler** — a Type burst calls `typeHeal` then the projection when suspended, and
   `typeHeal` then the render when allowed; one heal per burst; a non-Type burst never
   heals.
5. **Ribbon** — the out-of-set token reads as Default/Auto for the three captures.
6. **Contract** — `Custom Activity → Delineator` end to end: healed row offers the four
   corners, and the scene's delineator label position is resolved from the healed row.

## Docs

- `docs/07-GANTT-ENTITY-GUIDE.md` §23/§24: state that `Auto`/`None` belong to both
  vocabularies, and that a Type change blanks the cell across the boundary (owner
  ruling). The sentence at §23 line 743 needs two corrections, not one: a which-date
  value on a delineator does **not** always refuse `DateLabelNotAllowedForType` — a value
  such as `Start` fails the corner parse and reports `UnknownDateLabelPosition`, while
  `Auto` and `None` are accepted as corners.
- `docs/DECISIONS.md`: the R1–R7 rulings in one entry.
- New work item `docs/work-items/R5.23-type-change-transition.md` (ID to confirm) with
  the acceptance criteria and gates.
- `docs/STATUS.md` at implementation.

## Risks and deferred

- **Old type is knowledge, not data.** The change-path heal reads it from
  `RibbonStateService`'s cached pre-edit rows. A row that has never been projected (a
  brand-new insert, or a workbook opened mid-session) has no old type, so only the
  idempotent rules apply. That is safe but not a full boundary reset; the insert path
  already writes a correct style key and label values.
- **A user-authored style has no stored family.** R5 judges it by capability, which is
  the same projection the validator and ribbon use. A style that is a mix of both
  families is accepted for both; that is an authoring question, not a transition one.
- **The non-requiring StyleKey is not the transition's business.** `StyleKeyWritePlanner`
  already rewrites it to the type default on every pass; the change-path heal therefore
  leaves it, and a hand-typed mismatch on an Activity is corrected on the next pass.
  Adding the rule here would be a second authority and would break the existing
  "all legal values are left alone" fact.
- **Hierarchy (R7)** is reported, not repaired. A prompt is a separate work item.
- The heal adds one table read and one config read on a Type burst; the render already
  reads both.
