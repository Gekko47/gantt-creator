# Auto-refresh streamlining — check inventory and further speed-ups

Companion to [auto-refresh-performance-plan.md](auto-refresh-performance-plan.md). That plan closed
P0/P1/P4 and ADR-0054 (O1–O5); this one asks the next three questions: where the remaining time is,
which of the pass's checks are actually needed on an automatic refresh, and whether the chart can
avoid being rebuilt at all when nothing that affects it changed.

Everything below is derived from the code in `src/GanttCreator.Office/GanttRefreshOrchestrator.cs`
and the neighbouring adapters, and from the measured phase timings in the rolling log. Estimates are
labelled as estimates.

## 1. The measured pass, per phase

Automatic passes only, attributed to a build by its timestamp, host-corrected to the pre-P1 session
using the phases no change in this series touched (`read`, `sheet`, `measure` are shown raw so the
correction stays visible).

| Phase | Pre-P1 raw | Fix build raw | Fix, host-corrected (÷1.4–2.1) | What runs in it |
|---|---|---|---|---|
| `read` | 9.5 | 19.9 | **9.5–14.2** | table `Value2`, config read, guard, column presentation, identity repair, type normalise |
| `validate` | 0.2 | 0.1 | 0.1 | `GanttRowValidator.Validate` (pure CPU) |
| `sheet` | 2.9 | 6.2 | **3.0–4.4** | highlight reconcile, outline plan + level reads, duration plan, style-key plan |
| `measure` | 12.1 | 17.0 | **8.1–12.1** | row-height normalisation, panel-grid measurement, scene request factory |
| `build` | 2.6 | 1.2 | 0.6–0.9 | `SceneBuilder` + `SceneShapeRenderer` (pure CPU) |
| `write` | 205.1 | 231.1 | **110.6–165.1** | `ListOwned`, per-shape verification, shape operations, z-order |
| total | 232.3 | 275.5 | **131–197** | |

Two conclusions follow directly:

- **The CPU is free.** `validate` + `build` ≈ 1 ms of a 275 ms pass. Nothing here is worth optimising
  for speed; it is all COM.
- **`write` is ~80% of the pass and is now legitimate work** — 8 shapes per pass really did change.
  The only ways to keep cutting are to write *less of each shape* (proposal B) or to *not write at
  all* on passes that changed nothing that reaches the chart.

The non-write floor is **44.4 ms raw / 21–32 ms corrected**: `read` + `measure` + `sheet`. A pre-P1
zero-write automatic pass measured 42.2 ms average (min 30.7), which is that same floor on a faster
host — consistent, and it says the floor has not moved.

## 2. What the pass actually asks the host

From the adapters, per automatic pass, `N` = body rows (11) and `S` = owned shapes (86–87):

| Where | Host calls | Notes |
|---|---|---|
| `ExcelGanttTableReader.Read` | 1 `DataBodyRange` + 1 `Rows.Count` + 1 `Value2` + 1 table lookup | the whole table in one bulk read |
| `ExcelGanttRowIdentityRepairer.Repair` | **+ 1 full `Value2`** + `FindColumnIndex("Id"/"ParentId")` walks | reads the same body again to check Ids |
| `ExcelTypeChangeNormaliser.Normalise` | **+ 1 full `Value2`** + **11 `FindColumnIndex` walks** | reads the same body again, then resolves 11 columns by name |
| `ExcelConfigCatalogueReader.Read` | several ranges on the VeryHidden sheet | ~unchanged between passes |
| `ExcelColumnPresentationPort.EnsureClassification` | per managed column (`Hidden`, `Locked`) | ~unchanged between passes |
| `ExcelRowWarningHighlighter.ReconcileHighlights` | minimal diff against tracked addresses | 0 writes when the issue set is unchanged |
| `ExcelOutlineGroupWriter.Apply` | **N `OutlineLevel` reads** | writes only rows not already at the target level |
| `ExcelRowHeightNormaliser.Normalise` | `N` heights, or 0 when P0a supplies them | writes only wrong rows |
| `ExcelPanelGridMeasurement.Measure` | per-column `Width` + reserved-row heights + header height | ~unchanged between passes |
| `ShapeReconciler` / `ExcelShapeWriter` | `ListOwned` = **3 reads × S**; verification = **~7 reads × unchanged candidates**; writes for changed shapes; z-order as of O5 | the largest single block |

**The duplicate reads are the clearest waste.** The table read already returns every column as a typed
cell (verified against `GanttRowDto`: Id, LaneId, StackIndex, Type, Description, Start, Finish,
Duration, ParentId, SiblingOrder, StyleKey, LabelPosition, DateLabelPosition, FillColour, StrokeColour,
Visible, SortOrder). The identity repairer and the type normaliser then re-read the same body through
`Value2` and re-resolve their own column indexes. That is three bulk reads of the same range per
refresh.

## 3. Are the checks necessary?

None of them is unnecessary — each prevents a specific defect. The question that matters is
**frequency**: several run on every pass to detect a condition that only a handful of edit kinds can
create.

| Check | Protects against | Can it be conditional? |
|---|---|---|
| Full table read | any row affecting scene, identifiers, hierarchy, layout | **No** — it is the input to everything. Keep. |
| Validation | an invalid row plotting as if valid | **No** — and it costs 0.1 ms. Keep. |
| Protection guard | writing to a protected sheet | **No** — one call, and it is the mutation boundary. Keep. |
| Highlight reconcile | stale error fills | Keep, but it is already differential (0 writes when unchanged). |
| Config read | the config sheet being the single authority for styles/metrics/settings | Yes — **only if** something wrote the config sheet. It is written by the config sheet itself and by the ribbon dialogs, both of which raise a sheet-change event. Proposal C. |
| Column presentation | a user-unhidden engine column staying visible | Yes — only a header/column operation can break it. Proposal C. |
| Outline level reads | a user manually regrouping rows | Yes — only a hierarchy change, or a manual grouping, can stale it. Proposal D (needs a ruling). |
| Row heights + panel measure | a user dragging a row height, or a metrics change | Partly — P0a already removed the duplicate height read; the grid build itself can be reused when its inputs are unchanged. Proposal D (needs a ruling). |
| Identity repair body read | blank/duplicate Ids | **The check is necessary; the second read is not.** Proposal A. |
| Type normalise body read | stale engine-hidden cells after a Type change | Same: necessary check, redundant read. Proposal A. |
| `ListOwned` | a shape the add-in owns but the scene no longer wants (deletions), and a user's shape sharing an id | Yes, with care — the snapshot already carries the last owned set, and the plan already accepts `Owned` as an input. Proposal E (needs a ruling). |
| `GeometryMatches` per unchanged shape | a shape the host moved on its own (row insert/hide/re-height) or a user dragged | Yes — the host-moved cases all change the layout inputs, so a layout hash covers them; a hand-drag does not. Proposal F (needs a ruling). |

## 4. Proposals

Ordered by risk, safest first. Savings are per automatic pass on the measured fixture (86–87 shapes,
11 rows), host-corrected.

### A — One bulk read for the whole preflight. No behaviour change.

The orchestrator already hands the normaliser the `Type` column from the bulk read (P0b). Extend that:
`IGanttRowIdentityRepairer.Repair` and `ITypeChangeNormaliser.Normalise` take `IReadOnlyList<GanttRowDto>`
(and the table reader's column map, or a resolved name→index map made once) instead of re-reading
`DataBodyRange.Value2` and walking `ListColumns` by name. The COM versions stay for the standalone
callers.

- **Saving:** 2 full body `Value2` reads + ~13 `ListColumns` name walks per pass. Estimated **5–10 ms**
  (part of `read`).
- **Fidelity risk:** none. Same cells, same pass, same decisions; the DTOs carry every column as a raw
  typed cell, so nothing is coerced that the adapters previously saw raw.
- **Tests:** an adapter test asserting **zero** host `Value2` reads when rows are supplied, and the same
  repair/normalise outcome as the COM path for the same input; plus the existing adapter suites
  unchanged.
- **Decision:** none. This is a pure deduplication and the natural next commit.

### B — Part-scoped updates (text-only, style-only). Extends ADR-0053's trade.

ADR-0053 splits the payload into *geometry* and *appearance* and moves only the geometry part. It
already computes both hashes; it stops there and re-applies everything when appearance changed. Split
appearance into **text** (description, alignment, font family/size/bold, text colour → `TextFrame2`
only) and **style** (fill, stroke, width, hatch, dash → `Fill`/`Line` only), and plan
`ShapeUpdateScope.Text` or `.Style` when exactly one changed.

- **Saving:** a description or StyleKey edit currently writes 6–10 shapes at 14–21 ms each. Text-only
  is ~10–12 round trips instead of 30–40, style-only similar: **~5–8 ms per changed shape**, i.e.
  **60–100 ms** off such a pass.
- **Fidelity risk:** the same trade the owner accepted on 2026-11-04, extended from the box to any
  unchanged part — a shape hand-recoloured and then re-described keeps the hand colour until an
  explicit Refresh.
- **Tests:** classification tests per part; writer tests asserting a text-only update touches no
  `Fill`; the existing "explicit Refresh re-applies everything" pin unchanged.
- **Decision needed:** confirm the 2026-11-04 ruling covers text/style, not just geometry.

### C — Skip `config read` + `column presentation` when the change did not touch them. P3 revisited.

`ISheetChangeAnalyzer` already classifies the changed address (table body vs header vs outside, and
which columns). Pass that classification into the refresh as a hint: a change wholly inside the body,
in a non-config workbook, cannot have altered the VeryHidden config or the managed columns' hidden
state.

- **Saving:** part of `read`, estimated **3–7 ms**.
- **Fidelity risk:** medium. Whole classes of bug the read currently absorbs would become "detected on
  the next explicit Refresh or the next config-sheet edit". The hint must be conservative: any change
  not positively identified as body-only keeps the full path.
- **Decision needed:** how much self-healing the automatic path should keep.

### D — Reuse the measured panel grid when the layout inputs are unchanged.

Hash the layout inputs (included columns + widths, the reserved-row metrics, body heights, body count,
hidden flags, and the columns that drive lane structure: Type, StyleKey, LabelPosition,
DateLabelPosition, hierarchy). When the hash matches the last successful pass, reuse its `PanelGrid`
and skip the measurement and the outline level reads.

- **Saving:** `measure` + most of `sheet` ≈ **10–16 ms**, plus **~1 ms** of build if the scene comes out
  identical anyway.
- **Fidelity risk:** medium — a missed input means a stale layout until the next structural change. The
  input list is enumerable from `SceneBuildRequest`, so it is auditable, but it is a real invariant to
  maintain. Row-height drags and manual grouping fire no cell event, so they would be healed only by an
  explicit Refresh.
- **Decision needed:** same self-healing question as C.

### E — Pass `Owned` from the snapshot instead of calling `ListOwned`.

`ShapeReconcileRequest.Owned` already exists and is unused by the orchestrator; the reconciler already
keeps `BackToFront` in the snapshot. Supplying it turns `ListOwned` (3 reads × 87 ≈ **8–26 ms**) into a
dictionary lookup, with the full scan kept for the explicit Refresh. To stay correct when a user
deleted an owned shape, a plan that would update a missing shape must fall back to create.

- **Saving:** **8–26 ms** per automatic pass.
- **Fidelity risk:** medium-high. A shape whose ownership tag a user edited (R9.4) or that was deleted
  outside the add-in would be missed until an explicit Refresh, and the create-on-missing fallback is a
  new self-healing behaviour that needs its own tests.
- **Decision needed:** yes.

### F — Layout-gated verification. The largest remaining single item.

ADR-0052 verifies every unchanged candidate because the host moves shapes on row insert, row
hide/show, and row re-height — all of which change the layout inputs (D). If the layout hash matches
the last **successful and fully verified** pass, the only remaining way a shape can have moved is a
user dragging it, which no cell event reports.

- **Saving:** ~7 reads × unchanged shapes ≈ **18–60 ms** per automatic pass — the biggest non-write item
  after `ListOwned` — and it makes a no-op pass ~15–20 ms.
- **Fidelity risk:** high-visible. A hand-dragged bar would stay wrong until an explicit Refresh. Unlike
  a stale fill, this is a *geometry* error, which the owner has so far treated as unacceptable
  ("accuracy of the plot is key"). Recommend keeping verification and taking D/E instead unless the
  owner rules otherwise.
- **Decision needed:** yes.

## 5. The "do not replot" fast path

A pass in which nothing that reaches the scene changed currently costs the full floor: read, measure,
build, then a verify-and-skip loop over 87 shapes (`write` ≈ 100–160 ms corrected, of which almost all
is verification and `ListOwned`). With A + D + E the same pass would be **read + validate + hash ≈
15–25 ms** and issue no host write at all. With A + D alone (keeping verification) it is **~60–110 ms**.

That is the honest ceiling for "not replotting everything": the scene builder and the reconciler are
already delta-driven, so the remaining cost of an unchanged pass is *detection*, not rendering — and
detection is `ListOwned` plus one geometry read per shape.

## 6. Not proposed

- **Skipping validation or the table read** — they are the input to every decision, and together they
  cost ~10 ms.
- **Caching the scene across passes keyed on the payload hash** — the payload hash requires the scene,
  so it cannot gate the scene. Gate the *inputs* instead (D).
- **Any CPU-level optimisation** — `validate` + `build` are ~1 ms.
- **Parallelising COM** — forbidden by the STA rule and useless against a single host.
- **Raising the 50 ms coalescing window** — unchanged from the earlier analysis: the risk is duplicate
  passes, the saving is small.

## 7. Order of work

1. **A** — dedupe the preflight reads. No product decision, no behaviour change, ~5–10 ms. Do it next.
2. **B** — part-scoped updates, once the owner confirms the ADR-0053 trade extends to text/style.
   The largest win on the edit kinds users perform most (description and style-key edits).
3. **D** — layout hash, if the owner accepts that a manual row-height drag or manual grouping is
   repaired by the explicit Refresh rather than the next automatic pass.
4. **E** — `Owned` from the snapshot plus create-on-missing, if the owner accepts the self-healing
   change.
5. **F** — only if plot accuracy on a hand-dragged shape is judged less important than 18–60 ms.

Expected effect on the measured fixture, cumulative: 275 ms → **~265 ms** (A) → **~200 ms** on a
description-edit pass and **~185–205 ms** on a mixed pass (B, which only helps passes whose changes are
text/style-only) → **~190–200 ms** (D) → **~165–190 ms** (E). The no-op pass is where the effect is
dramatic: **275 ms → 15–25 ms** with A + D + E.

## 8. Validation

Contract tests per proposal, as listed. Then one live session, measured with the build-window
comparison already built for the previous entries, holding the host factor to the untouched phases. The
two figures to watch are `write` per changed shape (B) and `measure` (D); `read` (A/C) should fall
without any change in `desired`/`written`.
