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

- Part hashes (geometry / style / text) per primitive in `ShapeDeltaCache`,
  compared against our own last successful write; the reconciler passes the
  diff as an update scope; `ExcelShapeWriter.Update` applies only that scope.
- *Acceptance*: a date-only edit writes geometry only (the fake records the
  requested parts); a style-only edit writes no geometry; per-shape time
  ≤ 12 ms; golden scene unchanged; a refused update invalidates the snapshot
  (already the rule).

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

| After | Total (ms) | Composition |
|---|---|---|
| today | ≈236 | pipeline ≈40 + writes ≈120–240 + z-order 0/≈85 |
| P0 | ≈215 | row passes halved |
| P1 | ≈140 | writes ≈60–120 |
| P2 | ≈85 | z-order 0–15 |
| P3 | ≈70 | read ≤10 |
| P4/P5 | measured | — |

Click-to-plot adds the 50 ms command window to each figure.

## Implementation status (2026-11-03)

| Phase | State | Evidence |
|---|---|---|
| P0b — Type texts handed to the normaliser | shipped | `IRowHeightNormalisationPort.Normalise(..., typeTexts)`; depth tests prove no cell read when the list is supplied, and that a short list falls back per row |
| P0a — one row pass shared with the measurement | shipped | `RowHeightNormalisationOutcome.BodyHeightsPt` is returned only when the pass wrote no body row; `IPanelGridMeasurementPort.Measure(columns, bodyRowHeightsPt)` uses it only when its length matches the live body |
| P0c — column map without per-column reads | shipped | `ExcelSheetChangeAnalyzer` reads the header row in ONE bulk `Value2` call, falling back to `ListColumn.Name` for headers that are not non-empty text |
| P4 — post-pass measurement | shipped | `gantt.refresh post ms=...` after the pipeline on both command paths |
| P1 — write scope | not implemented | see below |
| P2 — minimal z-order moves | not implemented | see below |
| P3 — configuration cache | not implemented | see below |
| P5 — cold start | not re-measured | depends on P1 |

### Why P1 (write scope) is not implemented

Part-level skipping cannot be verified at a price lower than writing. A shape
whose style is not re-applied must be known to still carry the style last
written, and the only way to know that is to read the shape back -- fill, line,
font, text, alignment -- which is one read per property, i.e. the same COM
round-trip count as the write it replaces. Without that read, a shape the user
recoloured by hand and then moved (so its geometry changed) would keep the hand
colour until a manual Refresh, because the automatic pass would apply geometry
only.

The whole-shape skip that already ships (ADR-0052) is sound precisely because
geometry is verifiable in a handful of reads (`Left`, `Top`, `Width`, `Height`,
`Placement`), so a skipped shape is proven rather than assumed. Nothing
comparable exists for style, so P1 is deferred until a cheaper verification
exists.

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
