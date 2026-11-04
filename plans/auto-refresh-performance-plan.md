# Auto-refresh performance plan — end-to-end review

Status: planning, revised 2026-11-03 after a full end-to-end sweep of the
automatic refresh path. Sources: owner sessions in
`%LOCALAPPDATA%\GanttCreator\logs\gantt-creator-addin.log` (`gantt.refresh
timings` / `gantt.refresh result`), and the code paths in
`GanttCreator.AddIn` / `GanttCreator.Office`.

## Measured profile (warm command pass, 11 rows, 82–87 shapes)

| Phase | Range (ms) | Per-unit / notes |
|---|---|---|
| command window | 50 | was 300; the cell-edit window stays 300 |
| read (table + config) | 7–29 | one bulk `Value2` + config tables |
| validate | 0.1 | |
| sheet (outline, duration, StyleKey, heights) | 2–5 | N COM reads in the outline and height passes |
| measure | 8–27 | 2 COM calls per body row |
| build + translate | 0.5–1.5 | |
| **per-shape writes** | **≈120–240** | **≈16–34 ms each (median ≈21)** |
| **z-order** | **0 or 52–115** | ≈85 when it runs; 0 when skipped |
| total | 88–380 (median ≈236) | was 534–671 before the window change |

Cold first pass after opening: ≈2.1 s (read 113, measure 121, build 183,
write 1571 for 87 updates + 89 z-order) — JIT plus operation count, a one-off.

## Quality-retention principles (the rules every change must obey)

1. **The full read → validate → build pipeline runs on every automatic
   refresh, unchanged.** That is the parity guarantee with the manual Refresh;
   no reduction may come from validating or building less.
2. Reductions may come only from three safe classes:
   - **not re-reading what the same refresh already read** (single-threaded
     STA, events suppressed by the ADR-0020 scope, so nothing can change
     between the two reads);
   - **not re-writing what our own last successful write already left
     correct** (compared against our own record, never a host read-back, so
     R4.7 D5's "null means leave alone" objection does not apply);
   - **not re-issuing host commands whose effect is already in place**
     (verified against live geometry, per ADR-0052).
3. **Host state we did not produce must still be read.** User-dragged row
   heights, manual outline grouping, and ownership-tag edits are exactly the
   states a cache would go blind to; those reads are the quality floor and stay.
4. **The manual Refresh remains the full repair path.** It may get faster, but
   it never trusts a cache beyond verified skip.

## End-to-end catalogue

Legend: cost = measured or COM-call count; reduction = what can be removed;
quality = why the reduction is safe (or what must be preserved).

### A. Identify (event → scheduled)

| Step | Cost | Reduction | Quality argument |
|---|---|---|---|
| `SheetChange` → `ExcelSheetChangeAnalyzer.ShouldAutoRefresh` | **per event**: walks every sheet and `ListObject` to find the table, builds the column map, reads `DataBodyRange` — a full COM scan per changed cell, before the debounce even starts | **P0c**: cache the table/sheet/column-map resolution per workbook; invalidate on workbook switch or table-structure change | The table's location changes only on workbook switch or structural commands, both of which already reset scheduler state |
| debounce | 50 ms command / 300 ms cell | keep; going lower risks duplicate passes for asynchronously delivered events | the generation counter already discards stale timers |
| queue as macro | negligible | — | — |

### B. Read

| Step | Cost | Reduction | Quality argument |
|---|---|---|---|
| `ExcelGanttTableReader.Read()` | one bulk `Value2` (good) **plus** the same table/column-map scan the analyzer does | **P0c** covers it | as above |
| config read (styles, metrics, settings, type catalogue) | inside the 7–29 ms `read` | **P3**: cache, invalidated by every config writer (Theme dialog, Initialise, repair) and by any config-sheet change the analyzer classifies | the config sheet is `xlSheetVeryHidden`, so only the add-in writes it; the writer set is enumerable and each already schedules a refresh |
| identity repair + re-read | conditional re-read already removed | — | — |
| type-change normalise + re-read | conditional re-read already removed | — | — |

### C. Validate and report

| Step | Cost | Reduction | Quality argument |
|---|---|---|---|
| `GanttRowValidator.Validate` | 0.1 ms | none needed | — |
| highlight reconcile | own delta cache, writes only changed fills | none needed | — |

### D. Worksheet writes

| Step | Cost | Reduction | Quality argument |
|---|---|---|---|
| outline rebuild | reads `OutlineLevel` per body row (N COM reads) + writes changed spans | keep the reads: manual grouping via Excel's own UI is host state we do not own, and the reset pass exists to clear it. Flag as a product question only if it ever dominates | principle 3 |
| duration / StyleKey writes | bulk, plan-driven, no-op when correct | none needed | — |
| row-height normalise | per row: row lookup + `RowHeight` + **Type cell re-read** + `Hidden` for planned writes | **P0b**: the reader already read the whole table including `Type` in one bulk call — pass those texts in instead of N per-cell COM reads. Keep the `RowHeight`/`Hidden` reads (principle 3: user drags) | identical data, same refresh |

### E. Measure

| Step | Cost | Reduction | Quality argument |
|---|---|---|---|
| panel measurement | per body row: row lookup + `RowHeight` — **the same heights the normaliser just read** | **P0**: one shared row pass — the normaliser reads each row's height and hidden state and hands them to the measurement; the measurement keeps its column-width, origin, header, padding and anchor reads (ADR-0030/0031/0038 checks) | single-threaded, events suppressed; nothing can change between the two reads |
| scene build + translate | 0.5–1.5 ms | none needed | — |

### F. Render (reconcile)

| Step | Cost | Reduction | Quality argument |
|---|---|---|---|
| `ListOwned` | 2 COM reads per shape (name + tag) | keep — tag edits are the R9.4 scenario a cache would go blind to | principle 3 |
| signature compare | in-memory | — | — |
| verified skip (`GeometryMatches`) | ≈0.3 ms per candidate (indexed lookup + 5 reads) | keep; optionally scope by change set only if it ever matters at 1,000 shapes | this is what makes skip safe |
| per-shape `Update` | ≈16–34 ms: indexed lookup, tag, 4 geometry writes, placement, ≈5 style writes, text walk for text boxes | **P1 write scope**: part hashes (geometry / style / text) against our own last write; a date move becomes geometry-only (≈5 writes instead of ≈12+). Placement stays unconditional (1 call, it is the heal for legacy shapes) | principle 2b |
| creates / deletes | full set / 1 call | none | — |
| z-order | 0 (skipped) or ≈85 ms (one lookup + one command per shape) | **P2 minimal moves**: reorder only the shapes whose index differs from the last applied order; full pass when a create is present | the last applied order is our own record; the resulting order is provable on the fake |

### G. Post-pass (currently unmeasured)

| Step | Cost | Reduction | Quality argument |
|---|---|---|---|
| ribbon invalidation (8 controls) + status bar + state-scope restore | unknown — happens after the `result` line | **P5**: add a `post=` timing to the trace, then scope the invalidation if it is material | selection-dependent controls must re-evaluate after a command moves the selection |
| Excel repaint after `ScreenUpdating` is restored | unknown, grows with shapes touched | falls out of P1 (fewer writes → less repaint) | — |

## Phases

Ordered by measured value; each phase lists its acceptance criteria.

### P0 — Share the row pass (cheap, do with P1)

- **P0a** one shared per-row read pass feeding both the height normaliser and
  the panel measurement (height + hidden state per row, read once).
- **P0b** pass the already-bulk-read `Type` texts into the normaliser.
- **P0c** cache the table/sheet/column-map resolution per workbook for the
  analyzer and the reader.
- *Acceptance*: `sheet` + `measure` together ≤ 15 ms at 11 rows; a 1,000-row
  synthetic workbook measures the row passes in one pass; the analyzer performs
  no per-event table scan (asserted by a contract test with a counting fake).

### P1 — Write scope (largest single win)

> **Shipped 2026-11-04 as P1′ — command-scoped geometry (ADR-0053).** The request
> payload is hashed in two parts, geometry and appearance, per primitive in
> `ShapeDeltaCache`. A shape whose BOX changed and whose appearance did not is
> moved with `IShapeWritePort.UpdateGeometry` instead of being re-applied in full,
> and a box the host already holds (the row-insert case) is not written at all.
> Appearance changes, kind changes and every shape on the explicit Refresh keep
> the full update, which is what makes the manual path the repair pass.
> *Acceptance met*: `A_geometry_only_change_moves_only_that_shape`,
> `A_geometry_only_change_the_host_already_made_is_skipped` (zero writes, two
> verifications), `An_appearance_only_change_re_applies_the_whole_shape`,
> `A_kind_change_re_applies_the_whole_shape`,
> `Without_a_snapshot_a_moved_shape_is_re_applied_in_full`, and
> `UpdateGeometry_moves_the_shape_without_re_applying_text_or_style` (style and
> content recordings both empty, box and placement written). The unshipped
> remainder is appearance-level skipping, rejected below.

### P2 — Minimal z-order moves

- Keep the last applied order; issue send-to-back only for shapes whose index
  differs; fall back to the full pass when the plan contains a create.
- *Acceptance*: a two-row move issues a handful of `ZOrder` calls; the final
  order equals the scene's, proved on the fake's recorded call list plus an
  order simulation; `zorderms` ≤ 15 ms.

### P3 — Configuration cache

- Cache the config read; invalidate on every config writer and on any
  config-sheet change the analyzer classifies.
- *Acceptance*: `read` ≤ 10 ms warm; a Theme-dialog change still takes effect
  on the very next refresh (positive test).

### P4 — Post-pass measurement

- Add `post=` to the timings line (ribbon invalidation, state restore, status
  bar). Scope the invalidation only if the measurement says it is material.
- *Acceptance*: the number exists; any follow-up is decided by it.

### P5 — Cold start

- Re-measure after P1/P2; the 2.1 s first pass is JIT plus 87 operations, and
  write scope cuts the operation count. No JIT tricks.

## Expected cumulative effect (warm command)

Measured baseline (owner's session, 2026-11-04, 46 automatic passes, 86–87
shapes): **191.6 ms average**, of which the reconcile is 169.5 ms and the
pipeline ~30 ms; 8 passes wrote nothing and cost 31–67 ms; the ordering pass ran
on 12 passes at 49–111 ms; a manual Refresh is 859–1 129 ms.

| After | Total (ms) | Composition |
|---|---|---|
| measured baseline | ≈192 | pipeline ≈30 + writes ≈140 + z-order 0/≈60 |
| P0a/P0b/P0c/P4 | shipped | row passes and per-event column reads removed; `post` measured at 0.3–1.4 ms |
| P1′ (geometry scope) | ≈40–100 | a move is ~0.5 ms/primitive, so a relocation pass drops from 257–335 ms to ~40–70 ms and a date edit to the ~40 ms floor |
| P2 | ≈35–70 | z-order is now the largest remaining term (49–111 ms on reordering passes) |
| P3 | ≈30–60 | read ≤10 |
| P5 | measured | first-pass create cost |

Click-to-plot adds the 50 ms command window to each figure.

## Implementation status (2026-11-03)

| Phase | State | Evidence |
|---|---|---|
| P0b — Type texts handed to the normaliser | shipped | `IRowHeightNormalisationPort.Normalise(..., typeTexts)`; depth tests prove no cell read when the list is supplied, and that a short list falls back per row |
| P0a — one row pass shared with the measurement | shipped | `RowHeightNormalisationOutcome.BodyHeightsPt` is returned only when the pass wrote no body row; `IPanelGridMeasurementPort.Measure(columns, bodyRowHeightsPt)` uses it only when its length matches the live body |
| P0c — column map without per-column reads | shipped | `ExcelSheetChangeAnalyzer` reads the header row in ONE bulk `Value2` call, falling back to `ListColumn.Name` for headers that are not non-empty text |
| P4 — post-pass measurement | shipped | `gantt.refresh post ms=...` after the pipeline on both command paths |
| P1′ — command-scoped geometry | shipped | `ShapeDeltaCache` stores kind, geometry and appearance hashes; `ShapeReconciler` plans `ShapeUpdateScope.Geometry` for a box-only change, verifies the live box first, and `ExcelShapeWriter.UpdateGeometry` applies placement and the box only (ADR-0053) |
| P1 (appearance scope) | not implemented | see below |
| P2 — minimal z-order moves | not implemented | see below, and now the largest remaining term |
| P3 — configuration cache | not implemented | see below |
| P5 — cold start | not re-measured | depends on P1′ |

### Why appearance-level skipping (the rest of P1) is not implemented

The GEOMETRY scope ships because geometry can be verified: four absolute numbers
plus `Placement` (ADR-0052), with no `null`-means-leave-alone member. Appearance
cannot. `TextColour`, `FontFamily`, `FillColour` and friends use `null` to mean
"leave the host alone" (R4.7 D5), so after a write the intent is unobservable and
an un-applied part cannot be proven still correct. Reading it back is one read per
property — the same round-trip count as the write it would replace — so a
"skip the style" scope would either cost what it saves or silently accept a stale
style. Writing the style when the style changed is therefore both the cheapest
verified option and the repair.

### Why P2 (minimal z-order moves) is not implemented

The number of moves can only be reduced by knowing each shape's current index,
and `Shape.ZOrderPosition` is one COM read per shape -- the same round-trip
count as the `ZOrder(msoSendToBack)` call it would avoid. The delete-only skip
that already ships (order unchanged, no create) covers the common case; the
remaining full pass happens only when the scene really reordered.

### Why P3 (configuration cache) is not implemented

The configuration sheet is user-editable: it is VeryHidden, not read-only, and
unhiding it is a supported repair path. An invalidation wired only to the
add-in's own config writers would go stale on exactly the edits it cannot see,
and the next automatic refresh would render from a cached style or metric the
user had already changed. The read is one bulk call on a small sheet; the risk
is not worth the milliseconds.

### P5 (cold start)

Depends on P1, so it is not re-measured here. The first pass stays dominated by
create-time COM work.

## Round-trip budget and the post-P1′ optimisation set (2026-11-04)

Where the time actually goes, per COM round trip (RT), derived from the code and the
within-session measurements. Reads on a held RCW measure ≈0.03–0.1 ms; **writes** ≈0.5–1.2 ms
(the host re-layouts); **collection walks** (`Sheets`, `ListObjects`, `Shapes[i]`) ≈0.2–0.5 ms
per step. That asymmetry is the fact the whole plan now rests on.

**Fixed per refresh (every change type)**

| Work | Round trips | Measured |
|---|---|---|
| Application-state scope (5 settings + selection capture/restore) | ~10–12 | inside `post` 0.3–4.5 ms |
| Table read (one bulk `Value2`) + config read | 2 bulk + walks | `read` 4–24 ms |
| Outline / duration / StyleKey / heights / measure | ~15–25 | `sheet` 1.4–7.6 + `measure` 4.5–39 ms |
| `ListOwned`: `Shapes[i]` + `Name` + `AlternativeText` per shape | **3 × N ≈ 258** | ~5–15 ms |
| `GeometryMatches` per unchanged candidate: resolve + lookup + tag + placement + 4 box reads | ~10 × unchanged | reads only, cheap |

**Per changed primitive**

| Path | Round trips | Of which writes | Measured |
|---|---|---|---|
| Move (P1′) | guard 1–2 + resolve ~6 + lookup 1 + tag 1 + box 4 + placement 1 ≈ **14–16** | 5 | 8.0–10.4 ms |
| Full update | + style ~10–15 + text ~8–12 ≈ **30–40** | ~20 | 6.7 ms marginal at scale; 16–20 ms in small batches |
| Create | AddShape + placement + style + text + tag stamp ≈ 25–35 | ~22 | — |
| Delete | resolve + lookup + tag + `Delete` ≈ 10 | 1 | — |

**Z-order pass (when order changed):** resolve + lookup + `ZOrder` per shape ≈ 2–3 RT/shape,
one write each → 0.6–2.7 ms/shape → **50–230 ms per pass**. On a manual refresh it is
~60–215 ms of the ~1–3.6 s.

**Redundant work the budget exposes**

1. `ResolveShapes` re-walks `Sheets` → `ListObjects` → guard → `Shapes` on **every**
   Create/Update/UpdateGeometry/Delete/`GeometryMatches`/`ApplyZOrder` call — ~6 RT,
   two of them collection walks, per primitive, for a target that cannot change mid-refresh
   (`EnableEvents` is off).
2. `ApplyPlacement` writes `Placement = xlMove` unconditionally on every update, including
   the 85 shapes that already report `xlMove`.
3. `ListOwned` reads the ownership tag for **every** shape on the sheet, including user
   shapes whose names cannot parse as ours.
4. `ApplyZOrder` re-resolves and re-looks-up every name the update loop just held.
5. Not material, recorded for completeness: `Enum.IsDefined` and `Kind.ToString()` per
   primitive (µs), 3 × SHA-256 per primitive on commit (µs), band tracing (5 reads on
   band-named shapes only), `SelectBodyCell` (2 RT once per command), JSON boxing none
   (serialisation is hash-input only).

**Status: O1–O5 shipped 2026-11-04 as [ADR-0054](../docs/adr/0054-refresh-round-trip-budget.md).**
Contract-tested without a host (Office 1295, AddIn 656, Architecture 120 green). The live gate
is `Not run`: a host session is needed to confirm the within-session ratios, and no session was
available during the change. The sketches below are kept as the design record.

### O1 — One resolution and one guard per reconcile

- **Sketch:** `IShapeWritePort` gains a batch scope (`using var batch =
  writer.BeginBatch()` in `ShapeReconciler.Reconcile`); the writer resolves `Shapes` and
  queries the guard once and reuses them for every operation in the batch. `FindShapeByName`
  stays per primitive (names change); the batch ends on dispose or on the first refusal.
- **Expected saving:** removes ~6 RT (2 walks) per primitive → **1.5–3 ms per primitive**;
  a 42-primitive relocation pass **60–120 ms**, a 6–12 primitive pass **10–35 ms**, and the
  z-order preflight's resolution **25–60 ms** on reordering passes.
- **Fidelity risk:** low. Same proxies, same checks, once. `Shapes` is a live collection
  proxy, so creates/deletes inside the batch do not stale it; the guard cannot change
  mid-refresh with events off. Keep the per-target check (R4.8 D3) — it just runs once.
- **Validation:** within one session, all-move passes should fall from 8–12 ms/primitive
  toward **4–6 ms**, and `zorderms`/shape from 0.7–2.0 toward **0.4–1.0**.

### O2 — Reuse shape proxies into the z-order pass

- **Sketch:** the batch keeps `Dictionary<string, Excel.Shape>` of proxies it already
  resolved; `ApplyZOrder` uses it, falling back to the indexed lookup for misses; `Delete`
  removes the entry.
- **Expected saving:** one lookup per shape ≈ **20–50 ms** on a reordering pass (overlaps
  O1's preflight saving; count once).
- **Fidelity risk:** low; the dictionary is batch-scoped, not cached across refreshes.

### O3 — Skip the placement write when the shape already reports `xlMove`

- **Sketch:** read `Placement` (one cheap read) before writing; skip when already
  `xlMove`. `GeometryMatches` already reads it, so the reconciler can also pass
  "placement verified" for the skip candidates.
- **Expected saving:** ~0.5–1 ms per primitive on warm passes (every shape is `xlMove`
  after the first heal) → **20–40 ms** on a 42-primitive pass.
- **Fidelity risk:** none — an idempotent write skipped only when the read proves it. The
  first refresh after the rule changed still heals legacy `xlMoveAndSize` shapes.

### O4 — `ListOwned`: read the tag only for names that can be ours

- **Sketch:** check the name's `<rowId>:<role>` shape (a string test, no COM) before
  reading `AlternativeText`; keep the phase-0 inventory logging for chart-shaped names.
- **Expected saving:** ~1 read × N non-ours shapes ≈ **1–3 ms** per refresh. Small; do it
  with O1, not alone.
- **Fidelity risk:** none — a name that cannot parse as ours can never carry our tag, and
  the tag read for candidates is unchanged (R9.4 keeps its evidence).
- **Shipped with the assumption enforced, not assumed.** The scan accepts exactly two forms:
  a role-derived `<ownerId>:<role>` and the chart title, whose identifier is now the single
  constant `ExcelShapeWriter.ChartTitlePrimitiveId`. `Validate` refuses any other identifier
  (`ShapeWriteRefusal.UnsupportedIdentifierForm`), so a shape a refresh could never recognise
  is never created — previously the pre-filter would have made such a shape silently
  un-updatable instead.

### O5 — P2 revisited: minimal z-order moves (the rejection is withdrawn)

- The original P2 rejection assumed reads and writes cost the same. The zero-write passes
  disprove that: 87 shapes verify in ~20 ms, so a read is ~0.03–0.1 ms while a `ZOrder`
  command is a write at ~0.6–2.7 ms. Reading position to avoid writes is now **worth it**.
- **Sketch:** read `ZOrderPosition` for the desired back-to-front list, find the longest
  already-correct **front run**, and issue `ZOrder` only for the shapes behind it; a create
  still forces the full pass (existing rule).
- **Expected saving:** a pass where 2–4 primitives moved: **50–90 ms → 10–15 ms**. A full
  reorder (manual refresh) is unchanged.
- **Fidelity risk:** medium, and it materialised. Two independent facts decide which end may
  be skipped: `ZOrderPosition` 1 is the shape at the **back** (Microsoft's remarks), so a
  correct run has *increasing* positions; and `msoSendToBack` is **absolute**, so untouched
  shapes always end in front of every moved shape and only a front run can be left alone.
  The first implementation inverted the comparison *and* skipped the back end, which
  mis-stacked the chart on both the automatic and the full refresh while every test passed —
  the double gave shapes fixed positions, so a test could only see the call list. Fixed in the
  commit that follows `c55dda0`; the double now models the collection and applies
  `msoSendToBack`, so the tests assert the resulting order.
- **Validation:** `zorderms` on small-reorder passes; the permutation test (all six 3-shape
  starting orders must end at the desired order) is the gate.

### Not proposed (checked and rejected)

- **Appearance-level write scoping** — unchanged rejection (R4.7 D5: `null` means
  "leave the host alone", so an un-applied part cannot be proven).
- **Caching the owned set** across refreshes — R9.4's tag-edit scenario; `ListOwned`'s
  reads stay.
- **Config cache (P3)** — unchanged correctness rejection; the config read is one bulk
  call inside `read`.
- **ScreenUpdating / event reentrancy** — already correct: the ADR-0020 scope disables
  events for the whole refresh, so the pipeline's own writes never re-enter the scheduler,
  and `ScreenUpdating=false` batches the repaint. Nothing to gain.
- **Selection/activation churn** — `SelectBodyCell` is 2 RT once per command; the scope's
  selection restore is 1 `Range.Select()`. Immaterial.

### Expected effect (within-session, quiet host) — the shape O1–O5 were built to produce

| Pass | pre-O | O1/O3/O5 shipped |
|---|---|---|
| date edit (1–3 primitives) | 8–12 ms write | **3–5 ms** |
| relocation (39–42 primitives) | 257–335 ms | **80–140 ms** |
| small reorder (2–4 moved) | 150–250 ms | **60–100 ms** |
| zero-write | 17–34 ms | 17–34 ms |
| manual refresh | 0.9–3.6 s | ~0.7–2.8 s |

Every figure is a within-session comparison; cross-session averages stay unusable because
the host factor (1.5–3.5×) dwarfs the code's effect.

## Explicitly not done, and why

- **Skipping the outline level reads** when the hierarchy looks unchanged:
  manual grouping through Excel's own UI is host state we do not own
  (principle 3). Product question, not a performance decision.
- **Trusting a cached owned-shape set** to skip `ListOwned`'s tag reads: tag
  edits are R9.4's scenario; the reads stay.
- **Skipping row-height reads** from a cache: user drags fire no cell event, so
  a cache would go blind exactly when normalisation matters.
- **Reducing the command window below 50 ms**: risks duplicate passes from
  asynchronously delivered events; the saving is small and the failure is a
  wasted full pass.
- **Incremental scene building**: the build is 0.5–1.5 ms warm; the complexity
  budget is better spent on the write path.
