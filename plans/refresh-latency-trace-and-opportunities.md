# Refresh latency — end-to-end trace and reduction opportunities

Status: investigation, 2026-11-19; updated 2026-11-20. The diagnostics-only
instrumentation in §6A, the create-path template reuse in §6C, and the grouped appearance
update in §6B have shipped; the remaining optimisation opportunities are not implemented.

**Shipped (2026-11-19):** the scheduler writes
`gantt.refresh detect source=cell|command events=… detectms=… queuedms=… scheduled="workbook|sheet"`,
and the orchestrator writes `gantt.refresh phases ms table=… config=… guard=… columns=…
identity=… types=… validate=… report=… outline=… duration=… stylekey=… heights=… grid=…
request=… prewrite=… ownedms=… verify=… ops=… hiddenColumns="…" owned=… createdIds=…
deletedIds=…`.
`ShapeReconcileOutcome` now carries `OwnedCount`, `OwnedMilliseconds`, `VerifyMilliseconds`
and `OperationMilliseconds`. Existing log lines are unchanged. The live 20-activity
re-measurement is still owed.

Sources:

- The production rolling log (`%LOCALAPPDATA%\GanttCreator\logs\gantt-creator-addin*.log`,
  1,133 parsable refresh passes, 2026-09-11 → 2026-11-19). Parsed with
  `_artifacts/refresh-log-analysis.ps1` (read-only; not committed).
- The code paths named below.
- The existing performance record: [`auto-refresh-performance-plan.md`](auto-refresh-performance-plan.md),
  [`auto-refresh-streamlining.md`](auto-refresh-streamlining.md), and the ADR-0053/0054 work.
- One measured 20-activity live trace already recorded in [`STATUS.md`](../docs/STATUS.md):
  a warm refresh of **0.19 s**, a cold refresh of **2.5 s**, and 340 staged primitives
  against a 438-shape collection for the editable copy.

Nothing here asserts a host result that was not observed. Projections are labelled.

## 1. The path a user actually waits for

Two entry paths converge on one pipeline:

| # | Step | Host cost | Measured (automatic / manual) | Where |
|---|---|---|---|---|
| 0 | `SheetChange` → address string (`workbook|sheet!A1`) | 2–3 reads per event | **not logged** | [`ExcelApplicationAdapter.OnSheetChange`](../src/GanttCreator.Office/ExcelApplicationAdapter.cs) |
| 1 | `ExcelSheetChangeAnalyzer.ShouldAutoRefresh` | `ActiveWorkbook`, `Sheets` walk, `ListObjects` walk, header-row bulk `Value2`, `DataBodyRange`, `Rows.Count` | **not logged; runs per changed cell before the debounce** | [`ExcelSheetChangeAnalyzer`](../src/GanttCreator.Office/ExcelSheetChangeAnalyzer.cs) |
| 2 | Debounce | none | **300 ms** cell edit / **50 ms** post-command | [`GanttRefreshScheduler`](../src/GanttCreator.AddIn/GanttRefreshScheduler.cs) |
| 3 | `ExcelAsyncUtil.QueueAsMacro` | one message-loop turn | **not logged** | [`GanttRefreshScheduler.OnTimerElapsed`](../src/GanttCreator.AddIn/GanttRefreshScheduler.cs) |
| 4 | Protection window opens (release) | read + unprotect | inside `command − total` | [`WorksheetProtectionWindow`](../src/GanttCreator.Office/WorksheetProtectionWindow.cs) |
| 5 | Orchestrator construction | no COM in the constructors | inside `command − total` | [`RefreshSheetCommand.CreateLiveOrchestrator`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs) |
| 6 | Application-state scope | 5 setting reads + 5 writes + selection capture | inside `total` (scope open) and `command − total − post` (restore) | [`ExcelApplicationStateScope`](../src/GanttCreator.Office/ExcelApplicationStateScope.cs) |
| 7 | `read` | table `Value2` + config tables + guard + columns + identity `Value2` + type `Value2` | auto median **9.5 ms**, manual **14.0 ms** | [`GanttRefreshOrchestrator.RunRefresh`](../src/GanttCreator.Office/GanttRefreshOrchestrator.cs) |
| 8 | `validate` | pure CPU | **0.1 ms** | [`GanttRowValidator`](../src/GanttCreator.Core/GanttRowValidator.cs) |
| 9 | `sheet` | highlight reconcile, outline levels, duration plan, StyleKey plan | auto **2.4 ms**, manual **5.1 ms** | orchestrator steps 6–7b |
| 10 | `measure` | row heights, per-column widths, reserved rows, grid | auto **7.4 ms**, manual **7.9 ms** | [`ExcelPanelGridMeasurement`](../src/GanttCreator.Office/ExcelPanelGridMeasurement.cs) |
| 11 | `build` | `SceneBuilder` + `SceneShapeRenderer` (pure CPU) | **0.6–0.7 ms** | [`SceneBuilder`](../src/GanttCreator.Core/Scene/SceneBuilder.cs) |
| 12 | `write` | `ListOwned`, skip verification, update/create/delete, z-order | auto **50.0 ms**, manual **179.5 ms** (medians) | [`ShapeReconciler`](../src/GanttCreator.Office/ShapeReconciler.cs), [`ExcelShapeWriter`](../src/GanttCreator.Office/ExcelShapeWriter.cs) |
| 13 | Protection window closes (re-protect) | read + protect cycle | inside `command − total` | [`WorksheetProtectionWindow`](../src/GanttCreator.Office/WorksheetProtectionWindow.cs) |
| 14 | Ribbon/status/tab post-pass | invalidate 7 controls, status bar, tab restore | **post 9.3–10.8 ms** | [`RefreshSheetCommand`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs) |

`total` is the pipeline (steps 6–12); `command` is the user-visible whole (steps 4–14);
`post` is steps 14 after the pipeline returns. The gap
`command − total − post` (≈ **25 ms automatic / 53 ms manual**) is the protection window
plus the scope restore/repaint plus orchestrator construction.

## 2. Measured profile

All figures are milliseconds, from the 1,133 parsed passes.

| Phase | Automatic median | Manual median | Share of a full manual rebuild |
|---|---:|---:|---:|
| `read` | 9.5 | 14.0 | 2.2 % |
| `validate` | 0.1 | 0.1 | 0.0 % |
| `sheet` | 2.4 | 5.1 | 0.5 % |
| `measure` | 7.4 | 7.9 | 1.0 % |
| `build` | 0.7 | 0.6 | 0.1 % |
| `write` | 50.0 | 179.5 | **86.0 %** |
| `zorderms` | 0.9 | 2.7 | 0.3 % |
| `post` | 9.3 | 10.8 | 1.2 % |
| **`total`** | **81.8** | **221.2** | 100 % |
| **`command`** | **120.1** | **285.5** | — |

Full manual rebuild subset (`written ≥ 140`, 72 passes, median 145 shapes):
`write` **1 296.4 ms**, `total` ≈ **1 507 ms**, `command` **1 578.7 ms**. The CPU phases
(`validate` + `build`) are **0.1 %**; the whole read/measure/build side is **3.5 %**.

Automatic passes broken down by what they wrote (warm passes only, `total < 1 s`):

| Automatic pass kind | n | `write` median | `total` median | `command` median |
|---|---:|---:|---:|---:|
| zero-write (all verified) | 48 | 10.7 | 27.8 | 53.4 |
| geometry-only updates | 78 | 18.3 | 40.4 | 75.7 |
| full updates (appearance changed) | 221 | 27.6 | 55.8 | 95.7 |
| creates present | 427 | **98.9** | **132.4** | **185.5** |
| the worst create-churn passes | 14–28 shapes / pass | **143–859** | **180–920** | **218–1 060** |

The single worst 20-activity-scale create burst in the retained logs is a full build of
**252 shapes**: `lookup 4.5 + add 456.3 + style 208.5 + ownership 138.5 = 807.8 ms` of
create work alone ([`gantt.shape.create-steps`](../src/GanttCreator.Office/ExcelShapeWriter.cs)).
A later 253-shape burst measured `add 509.0 + style 207.7 + ownership 164.0`.

## 3. Where the per-shape cost comes from

A full `Update` of one text shape is about **40 COM calls**:

| Part | Calls | Code |
|---|---:|---|
| ownership tag read | 1 | `CarriesOwnershipTagFor` → `Shape.AlternativeText` |
| geometry | 4 writes | `shape.Left/Top/Width/Height` |
| placement | 1 read + 0–1 write | `ReadPlacement` / `ApplyPlacement` |
| text frame | ~7 writes + 4 proxy reads | `TextFrame2` word wrap, auto size, anchor, 4 margins, `TextRange` |
| text content | 1 write | `TextRange.Text` |
| paragraph | 1 proxy + 1 write | `ParagraphFormat.Alignment` |
| font | 1 proxy + 3 writes | `Font2.Name/Size/Bold` |
| font fill | 1 proxy + 4 calls | `Font2.Fill.Solid/ForeColor.RGB/Visible` |
| shape fill | ~8 calls | `Fill.BackColor/ForeColor/Transparency/Visible/Solid or Patterned` |
| shape line | ~7 calls | `Line.ForeColor/Weight/DashStyle/Transparency/Visible` |

A bar (non-text) is about 20 calls. Measured `write / written` on 148-shape manual
rebuilds is **9.4–21.7 ms per shape** (median ≈ 10.6), and
`1 296 ms ÷ (148 × ~30 writes) ≈ 0.29 ms per write` — the measured phase decomposes
into the call count times the host's per-write cost almost exactly. That is why the only
lever with real size is **issuing fewer calls**, not doing the CPU work faster.

## 4. The 20-activity sample

The retained logs contain one 20-body-row pass, but it predates the current entity/label
set (82 primitives, `write=32 ms`), so it is not a valid 20-activity baseline. The
20-activity evidence that is current is:

| Observation | Value | Source |
|---|---:|---|
| full build of the 20-activity chart | 252–253 created primitives | log `create-steps n=252/253` |
| create work for 252 shapes | **807.8 ms** (add 456, style 208, ownership 139) | same |
| cold refresh of the 20-activity chart | **2.5 s** | `STATUS.md`, 2026-11-19 |
| warm refresh of the same chart | **0.19 s** | `STATUS.md`, 2026-11-19 |
| a 10-row / 148-shape manual rebuild | `write` 1.3–3.7 s; `command` 1.5–5.0 s | log, 2026-11-19 |
| a 10-row / 148-shape no-op pass | `total` 28 ms, `command` 53 ms | log |

Projection for a 20-activity chart at ~250 primitives, using the measured per-shape costs
(**inference**, not a new host run):

| Action | Estimated `write` | Estimated perceived (with 300 ms debounce) |
|---|---:|---:|
| no-op automatic (all verified) | 15–30 ms | ~350–380 ms |
| one date edit (5–9 primitives move) | 25–120 ms | ~370–470 ms |
| description/label edit (text shapes re-applied) | 60–200 ms | ~410–550 ms |
| full automatic rebuild (all 250 re-applied) | 1.5–5.0 s | 1.8–5.3 s |
| cold first render (250 creates) | 0.8–2.5 s | 2.5 s measured |

The **300 ms cell-edit debounce is the largest single term in every no-op and small-edit
interaction**; the pipeline below it is 28–56 ms. That is a product trade (it coalesces a
typing/paste burst), not a code defect.

## 5. Candidate sources, distilled

Seven candidates were checked against the measured data.

1. **Per-shape COM round-trip volume in `write`** — dominant and proven: 86 % of a full
   rebuild, 50–90 % of a changed automatic pass, and the measured 10.6 ms/shape matches
   ~30 writes × 0.29 ms/write.
2. **Render churn: passes that create/delete 14–28 shapes unnecessarily** — the worst
   automatic passes. In one 2026-11-19 session the desired set alternates between
   `134` and `147` primitives, the splitter band alternates between
   `x=305.75,w=141.526` and `x=5.75,w=441.526` (a 300 pt panel-width difference), and
   `created/deleted` alternate `14/28 ↔ 28/15`. `snapshot` equals the previous `desired`
   each time, so the delta cache is not stale. The two states look like "all user columns
   visible" versus "all user columns hidden" (`ADR-0065`); whether the user is toggling
   the Columns control or something in the measurement is oscillating is **not yet
   known**. It is the first thing to instrument.
3. **Per-event analyzer scan before the debounce** — a `Sheets`/`ListObjects` walk plus a
   bulk header read on **every** changed cell, entirely outside the timing log. Cost per
   event is unmeasured.
4. **Redundant bulk reads inside `read`** — the identity repairer and the type normaliser
   each re-read the whole body through `Value2` and re-resolve columns. Proposal A in
   [`auto-refresh-streamlining.md`](auto-refresh-streamlining.md), still unshipped:
   ~5–10 ms.
5. **Repeated measurement reads** — per-column widths, every body row's height, reserved
   rows. `sheet` + `measure` are 9–13 ms and could share a layout hash / range read:
   ~8–16 ms.
6. **`ListOwned` + skip verification on no-op passes** — 10.7 ms of the 27.8 ms no-op
   pipeline. `ListOwned` is one name read plus one tag read per shape; verification is a
   handful of reads per unchanged shape. Both are R9.4/R4.8 self-healing guarantees;
   removing them is an owner decision, not a free win.
7. **CPU cost / parallelisation** — falsified as a lever at this scale:
   `validate + build` ≈ 0.8 ms of a 1 500 ms manual rebuild.

**Distilled to two sources of the problem:**

- **P1 (certain):** the `write` phase issues ~20–40 sequential, single-property COM calls
   per shape. This is 86 % of the manual path and the largest term on every changed
   automatic pass.
- **P2 (cost confirmed, cause open):** some automatic passes re-create and re-delete
  14–28 shapes because the desired scene changes between passes. These passes cost
  0.14–3.7 s instead of ~30 ms. The cause must be identified before optimising, because
  the fix is different if the user is legitimately toggling a column than if the
  measurement is oscillating.

## 6. Opportunities, ranked by measured effect

### A. Instrument the two blind spots first (no behaviour change) — shipped 2026-11-19, lost, restored 2026-11-20

The log already has `read/validate/sheet/measure/build/write/zorderms/total/post/command`
and per-create steps. The missing evidence is:

1. `detect` — aggregate time and event count in `GanttRefreshScheduler.OnSheetChanged`
   and `ExcelSheetChangeAnalyzer.ShouldAutoRefresh`, per burst.
2. `queue` — time from timer fire to macro entry (the `QueueAsMacro` turn), and the
   workbook/sheet name captured at schedule versus at entry.
3. `read` split: `table`, `config`, `guard`, `columns`, `identity`, `types`.
4. `sheet` split: `outline`, `duration`, `stylekey`; `measure` split: `heights`, `grid`,
   `request`.
5. `write` split: `owned` (`ListOwned`), `verify` (`GeometryMatches`), `ops`, and the
   existing `zorder`/create steps.
6. For P2 only: `owned=N`, the first few `created`/`deleted` ids, the effective
   `HiddenColumns` setting, and the measured column widths, on every pass.

This is diagnostics-only: the trace sink is already an injected seam
([`GanttRefreshOrchestrator`](../src/GanttCreator.Office/GanttRefreshOrchestrator.cs),
[`ExcelShapeWriter`](../src/GanttCreator.Office/ExcelShapeWriter.cs)), and adding
`Stopwatch` boundaries does not change a decision. It is the smallest change that turns
P2 from a hypothesis into a fact.

### B. Batch the update path with `ShapeRange` (P1, largest certain win) — shipped 2026-11-20

`ShapeRange` applies a secondary-format object to every shape in the range:

- Microsoft Learn: setting `FillFormat`/`LineFormat` properties on a `ShapeRange`
  "will set the same properties for all the individual shapes"; `TextFrame2` is a
  `Shape`/`ShapeRange`/`ChartFormat` object.
- Reflection over the shipped interop assembly (`ExcelDna.Interop` 16.0,
  `src/GanttCreator.AddIn/bin/Release/.../Microsoft.Office.Interop.Excel.dll`) confirms
  `ShapeRange.Fill`, `.Line`, `.TextFrame2`, `.Left/Top/Width/Height`, `.ZOrder`.

The requests already carry an appearance signature (`ShapeDeltaCache.AppearanceSignature`,
`ShapeTemplateKey`), so grouping needs no new model:

1. Group text boxes by identical appearance and apply the frame margins, vertical anchor,
   word wrap, auto-size, font family/size/bold, alignment, text colour and fill via one
   `ShapeRange` per group. Write only `TextRange.Text` per shape (the docs are explicit
   that text editing is a single-shape operation).
2. Group bars/lines by identical fill/stroke/dash/width and apply `Fill`/`Line` per group.
3. Group by `(Top, Height)` for geometry: one `ShapeRange.Top/Height` per row group, then
   `Left/Width` per shape.

**Estimated effect (inference from 0.29 ms/write):** a text-shape update falls from ~40
calls to ~8–12, and a 250-shape pass from 1.5–5.0 s to **0.5–1.5 s**. Fidelity is
unchanged: every property is still written, just once per identical group instead of per
shape.

**Live probe (2026-11-19, Excel 16.0.20527.20196 x64):**
[`scripts/probe-shaperange-bulk-style.ps1`](../scripts/probe-shaperange-bulk-style.ps1)
measured 60 rectangles in 3 appearance groups and 30 text boxes in 2 groups.
`Shapes.Range` construction costs **1.1–1.4 ms** per group. Per-shape Fill/Line writes
cost **3.3–3.9 ms** per rectangle and **11.0–14.1 ms** per text box; the same writes
applied to one `ShapeRange` per group cost **8.5–9.7 ms for 20 rectangles** and
**35.9–37.9 ms for 15 text boxes**. Every shape read back the grouped fill, stroke,
weight, dash, transparency, margins, vertical anchor, font size/bold/name, alignment, and
font fill colour with **zero mismatches**. A mixed rectangle+text range accepted both
`Fill` and `TextFrame2` on this host, but the implementation should still group by kind so
a future host cannot refuse it.

Risks to prove with a live probe: a range containing a non-text shape must never receive
`TextFrame2` (group by kind); `Fill` on a line-only range must not be attempted; and the
resulting fill/line values must be read back on a live shape, as R4.6 already does.

**Shipped 2026-11-20.** `IShapeWritePort.ApplyGroupAppearance` groups the update requests
by `ShapeTemplateKey`, proves every name and ownership tag before writing any of a group,
and issues one `ShapeRange` `Fill`/`Line` write per group plus one `TextFrame2` formatting
write for text groups. Text content stays per shape through
`IShapeWritePort.UpdateContent`, and `ShapeReconciler.Reconcile` runs the group pass before
its per-shape loop (`UpdateContent` for non-geometry updates, `UpdateGeometry` for
move-only). Group-by-kind keeps `TextFrame2` off rectangle ranges, so the two probe risks
are closed. Contract tests:
[`ExcelShapeWriterTemplateTests`](../tests/GanttCreator.Office.ContractTests/ExcelShapeWriterTemplateTests.cs)
(one style call per identical group, one string per text shape, a missing shape refuses the
group before any write). Live read-back and measurement:
[`ShapeGroupUpdateIntegrationTests`](../tests/GanttCreator.Office.IntegrationTests/ShapeGroupUpdateIntegrationTests.cs).
Inside one batch, so resolution is hoisted for both sides, the gate measured **20 text
boxes: 454.2 ms shape-by-shape vs 76.7 ms grouped (5.9×)** on Excel 16.0.20527.20196 x64,
with every grouped fill, stroke, weight, and typography read back. Geometry batching (B.3)
remains unshipped.

**Live group cardinality and the style/typography split (2026-11-20).** The new
`gantt.shape group-appearance` trace on the owner's 10-row / 148-primitive workbook read
`shapes=148 groups=22 largest=32 singles=9 ms=1346.5–1531.5`, while the full-refresh `ops`
stayed at 1.45–2.62 s (9.8–17.7 ms/shape). So the combined `ShapeTemplateKey` did coalesce
(22 groups, not 148), but a group write costs ~61–70 ms on average because a text group's
`TextFrame2` block is ~36–40 ms against ~8–10 ms for a style group. `ApplyGroupAppearance`
now buckets the same requests twice, on disjoint member sets: `StyleGroupKey` (kind,
host-owned extent, fill, stroke, line width, hatch, dash, line direction) and
`TypographyGroupKey` (font family/size/weight, alignment, text colour, left inset). A text
box that differs only in fill shares its typography range; one that differs only in text
colour shares its style range. Every shape is resolved and ownership-proven once, before
the first write. Test: `ExcelShapeWriterTemplateTests.A_style_and_typography_group_independently`.
**Read-back verification (shipped 2026-11-20).** The split's own trace then separated the
group pass: `preflightms=9.6`, `stylems=969.2`, `textms=364.3` for 148 shapes. Style cost
6.5 ms/shape and typography 7.3 ms/shape **through the range**, so the host applies the
block per member and no further grouping can help. `ExcelShapeWriter` now compares the live
value first and writes only a mismatch, for geometry (four members, the same tolerance
`GeometryMatches` uses), style (fill visibility/type/pattern/fore/back/transparency and line
visibility/fore/weight/dash/transparency), typography (word wrap, auto-size, anchor, four
margins, alignment, font name/size/bold and font fill) and text content. Only members the
scene actually resolved are compared, so `null` still means "leave the host alone". A
converged refresh therefore pays ~150 ms of reads instead of the ~2.3 s of writes the trace
measured. Tests:
`StyleWriteTests.A_converged_shape_rewrites_neither_style_nor_geometry`,
`A_converged_label_rewrites_neither_typography_nor_text`, plus the existing mismatch facts.

**Live after, same workbook, same host (2026-11-20).** Four explicit refreshes of the
148-primitive chart:

| # | preflight | style | text | group ms | `ops` | `write` | `total` | `command` |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| 1 | 23.0 | 48.3 | 20.1 | 91.5 | 164.3 | 197.9 | 315.8 | 378.9 |
| 2 | 16.2 | 22.7 | 8.8 | 47.7 | 81.6 | 119.5 | 247.6 | **294.3** |
| 3 | 29.0 | 78.7 | 20.2 | 127.9 | 184.1 | 235.5 | 440.0 | 531.3 |
| 4 | 9.7 | 25.1 | 10.6 | 45.5 | 88.9 | 127.4 | 205.8 | **250.0** |

Against the pre-change warm manual passes (`write` 1 473–2 645, `command` 1 688–2 866):
`write` **120–236 ms** (≈7–22× lower), `command` **250–531 ms** (≈5.7–11.4× lower), and the
group pass **1 343–1 532 → 46–128 ms**. The preflight is unchanged (9.6 → 9.7–29.0 ms), which
is the expected signature: the reads stayed, the writes went. Move up/down also improves its
group pass from 41–120 ms to **0.3–2.0 ms**, but its `command` stays 196–470 ms because a
move is geometry writes plus the whole-scene `verify`/`owned`/read audit, not appearance.
The next move-path levers are therefore scoping `verify`/`owned` and the redundant `read`
(opportunities D and F), not more write batching.

### J. Column classification map (shipped 2026-11-20)

`ExcelColumnPresentationRestorer.EnsureClassification` resolved each of 26 schema columns by
scanning all 26 live columns per name: **676 COM `Name` reads per refresh**, the same O(N²)
shape already fixed in `StagingShapeAdapter`. `BuildColumnMap` now builds one ordinal map in
a single walk (26 reads). The first live rerun measured `columns` at **12.8–78.4 ms** across
seven passes; what remains is the per-column *flag* verification
(`Range`/`DataBodyRange`/`EntireColumn`/`Locked`/`Hidden`), not the name scan. A plain
converged 148-shape refresh read `columns=59.8 ms`. Test:
`The_schema_resolves_with_one_header_name_read_per_live_column`.

**Re-ranked profile after C and read-back** (one converged manual refresh, 148 shapes):

| term | ms | what it is |
|---|---:|---|
| `ops` | 138.5 | 148 × ~0.9 ms of read-back (4 geometry reads + tag + text) |
| `columns` | 59.8 | 26 × flag verification + table walk |
| protection open/close | ~87 | release + re-protect + outlining cycle (`command − total − post`) |
| `owned` | 40.0 | `ListOwned` tag read per shape |
| `measure` | 32.5 | row heights + per-column widths + reserved rows |
| `post` | 30.7 | 7 control invalidations + status bar + tab restore |
| `read` remainder | 22.8 | table/config/types/identity |
| `sheet` | 16.0 | report/outline/duration/style-key |
| `zorderms` | 8.2 | already minimal |

Column toggles reproduce the earlier churn as a legitimate cost rather than a defect: each
toggle changes the panel width, so the 10 grid lines are re-created at new x positions,
`desired` moves 134↔148, and one pass costs 439–1 157 ms with `ops` 255–675 and `zorderms`
27–82.

**Next candidates, by measured size and risk** (after L and M, from the post-memo medians in
§M):

1. **Protection window scope** (~40 ms open + ~25 ms close per pass; up to 76–126 ms of open
   time on a shape-churn pass): the window releases and re-protects around the whole
   pipeline, including the read/measure phases that cannot need it. Opening it only when the
   reconcile plan has an operation needs a decision because it changes the ADR-0020/0055
   ordering, and a refused plan must still restore the user's state. This is now the largest
   single item and the only one worth more than 25 ms.
2. **`columns` flag skip** (~38 ms measured): requires a persisted classification version plus
   a `HiddenColumns` comparison; the manual Refresh remains the repair pass, so a stale
   classification is always repairable by the user. A bulk multi-column read was probed and
   disproved (§I), so this is the only route left for the term.
3. **`grid` + `heights`** (~19 ms): panel grid measurement is a self-healing read whose cache
   would need an invalidation rule (row height and column width changes are user actions with
   no event the add-in can trust).
4. **`verify` remainder** (~25 ms): 148 x five geometry reads. Skipping it needs a proof that
   no host-driven move occurred, which is the `measure` row-grid comparison above; not
   separable.

The per-shape read floor shipped as opportunity K below.

### K. Group-pass geometry/content verification (shipped 2026-11-20)

The grouped pass already resolved and ownership-proved every shape; only to hand the whole set
back to `UpdateContent`, which re-resolved each name, re-read its tag, re-read four geometry
members and the string, and then usually wrote nothing. `ApplyGroupAppearance` now performs
those comparisons during its preflight (the same `MatchesGeometry` the verified skip uses,
placement included, so the ADR-0052 host-owned tolerances still apply) and returns the
identifiers that genuinely differ in `ShapeGroupWriteOutcome.Mismatched`. `ShapeReconciler`
dispatches `UpdateContent` for those identifiers only; a `null` list from a port that does not
verify keeps the old visit-every-request behaviour.

What this removes is not a write — the writes were already skipped — but the per-shape
dispatch: name resolution, ownership tag read, request validation and a second set of property
reads for 148 converged shapes. The cost model says that is the ~0.9 ms/shape `ops` figure,
i.e. the 130–140 ms term, minus the reads the group pass now performs anyway (which it must,
to decide).

**Proof.** Contract: `ShapeReconcilerTests.A_fully_verified_group_pass_skips_the_per_shape_content_dispatch`
(empty mismatch → no `Update:` call) and `A_group_pass_that_reports_a_mismatch_still_dispatches_that_shape`
(named id → exactly one). Live: `ShapeGroupUpdateIntegrationTests` asserts an empty list for
unchanged bars, exactly the two changed label ids for a text change, and an empty list again
for the same pass after the content write. Gates: `verify-quick.ps1` PASS 208.5 s (Office
contract 1623), `verify-office.ps1 -DeadlineSeconds 1200` PASS 83/83 in 15 m 27 s.

**Live after (same workbook, 148 primitives, 2026-11-20).** Nine passes:

| pass | desired | upd | geom | cre | del | `ops` | `write` | `total` | `post` | `command` |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| manual (snap=0) | 148 | 148 | 0 | 0 | 0 | 109.9 | 132.8 | 192.3 | 14.3 | **233.8** |
| automatic, 1 shape | 148 | 1 | 0 | 0 | 0 | 1.9 | 65.9 | 145.1 | 17.6 | **195.5** |
| automatic, column toggle | 134↔220 | 102–124 | 99–117 | 14–98 | 14–98 | 188–527 | 282–710 | 352–791 | 18–64 | 521–899 |

The manual pass is the cleanest signal because it has no snapshot, so every shape is a full
update: 148 updates cost `ops` **109.9 ms**, of which the `gantt.shape group-appearance` record
attributes **preflight 62.0** / style 35.1 / text 10.9. Before this change the same pass
measured `ops` 138.5 with preflight 9.6–29.0; the dispatch is gone and the exact reads it used
to repeat have moved into the pass that already held the shapes. **Net −28.6 ms**, not the
−100 ms the term suggested, because the verification reads are the same reads the dispatch
was paying.

The converged automatic pass is now **195.5 ms from click to end**, and its decomposition is
fully attributed:

| term | ms | what it is |
|---|---:|---|
| `verify` | 31.9 | 148 × verified skip (tag, placement, four geometry members) |
| `owned` | 29.6 | `ListOwned` walk: one `Name` + one tag read per shape |
| protection window | ~33 open + ~18 close | release before the pipeline, restore after it |
| `columns` | 27.6 | 26 × column classification flag verification |
| `grid` + `heights` | 23.5 | panel grid measurement |
| `types` | 7.8 | type-change normaliser, which re-reads the body it was already read for |
| `outline` | 5.2 | outline groups |
| `report` | 3.7 | validation report |
| `config`, `duration`, `identity`, `table` | 8.2 | remainder of `read`/`sheet` |
| `ops` | 1.9 | the one changed shape |

Column toggles remain legitimate churn: each changes the panel width, so the 10 grid lines and
22 plot splitters are re-created at new x positions (`create-steps n=98 add=237.7`), and a
show-all/hide-all pair costs 899 + 800 ms. That is 98 creates at ~3.2 ms each, not a defect.

### L. One ownership walk, one tag read per shape (shipped 2026-11-20)

The converged pass read the same shape facts twice in two separate passes: `ListOwned` walked
the collection reading every `Name` and every candidate's `AlternativeText`, then the first
name lookup built the batch's name index by walking every `Name` **again**, after which each
verified shape re-read its own tag. `ListOwned` now feeds both memos while it walks — the batch
name index and a batch-scoped `name → owned` map — and `ApplyOwnershipStamp`/`Delete` maintain
the map at the only two points that can change ownership.

The cache is **empty outside a batch** and cleared at both batch boundaries, because a
remembered "owned" answer that survived a batch would make R4.8 stop protecting a user's shape
the moment they edited its alternative text. `The_tag_memo_does_not_outlive_the_batch` is that
regression test (two batches, a tag edit between them, two live reads);
`The_owned_walk_feeds_the_batch_name_index_and_the_tag_memo` is the saving (two shapes, two
name reads and two tag reads for the whole pass, against 4 and 3 before).

Expected: `verify` 31.9 → ~22 and `owned` unchanged at 29.6, i.e. ~10 ms off every pass, more
on a churn pass where the index walk is over 148–220 shapes.

### C. Reuse create templates on the live path (P1, small change, measured 10×) — shipped 2026-11-19

The writer already implements opt-in template reuse: a shape whose appearance matches an
earlier shape in the same batch is created with `Shape.Duplicate` and only its geometry,
text string and ownership are written. The tests record the measured costs on the
20-activity copy: **2.3 ms for a rectangle, 8.6 ms for a text box, against 24.9 / 34.6 ms
to create them** ([`ExcelShapeWriterTemplateTests`](../tests/GanttCreator.Office.ContractTests/ExcelShapeWriterTemplateTests.cs)).
The live orchestrator constructs the writer **without** `reuseTemplates: true`
([`RefreshSheetCommand.CreateLiveOrchestrator`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs)),
so the live path pays full create cost. Enabling it for live creates is a one-argument
change with an existing test suite; it mainly helps the cold/first render (252 creates)
and every create-heavy automatic pass. It does not help updates.

The live factory now enables the flag
([`RefreshSheetCommand.CreateLiveShapeWriter`](../src/GanttCreator.AddIn/RefreshSheetCommand.cs)),
and `RefreshSheetCommandTests.The_live_shape_writer_reuses_create_templates` pins it.

**Live after (same 106-shape cold pass, Excel 16.0.20527.20196 x64):** `create-steps`
`style` **58.3 → 12.4 ms** (-79 %), `add` 132.7 → 87.7, `ownership` 23.4 → 17.0;
`write` **245.4 → 152.6 ms** (-38 %); `total` 656.1 → 559.9; `command` 958.0 → 819.7.
Host variance is present, but the style block is the measured mechanism: a duplicate
inherits the template's fill, stroke, dash, and typography, so only the first shape of
each appearance is styled.

### D. Deduplicate the preflight reads (P3, free ~5–10 ms)

Proposal A of [`auto-refresh-streamlining.md`](auto-refresh-streamlining.md): hand the
already-read `GanttRowDto` list and the resolved name→index map to
`IGanttRowIdentityRepairer.Repair` and `ITypeChangeNormaliser.Normalise` instead of both
re-reading `DataBodyRange.Value2` and walking `ListColumns` by name. Same data, same pass,
no behaviour change.

### E. Row-height and layout fast paths (P3, ~8–16 ms)

- Read the body range's `RowHeight` **once**. The existing comment is right that a mixed
  body's aggregate is unreliable, but a value equal to the token (and no hidden row) is a
  sound fast path; fall back to per-row reads otherwise.
- Hash the layout inputs (included columns and widths, reserved-row metrics, body heights,
  hidden flags, lane-driving columns) and reuse the last `PanelCellGrid` when the hash
  matches (proposal D). This also removes the outline-level reads on unchanged passes.
- Trade: a manual row-height drag or a manual grouping fires no cell event, so it would be
  healed by the explicit Refresh only. Owner decision.

### F. Analyzer/table-resolution cache — measured 2026-11-20, NOT worth building

The proposal was to cache the table, owning sheet and column map per workbook, because
`ShouldAutoRefresh` walks every sheet and every `ListObject` and re-reads the header row per
changed cell, before the debounce.

**The restored `detect` line measured it, and the answer is no.** Five `source=cell` bursts in the
owner's 18:23 run: `events=8 detectms=18.8`, `events=1 detectms=0.6`, `events=1 detectms=18.9`,
`events=1 detectms=0.3`, `events=4 detectms=18.7`. The steady-state cost is **0.3–0.6 ms per changed
cell**, not the 2.4–4.9 ms assumed here; and the ~18.7 ms figure appears three times with event
counts of 8, 1 and 4, so it is a **recurring one-off that does not scale with the number of
changes** — a per-workbook cache would not remove it, and would remove at most a fraction of a
millisecond per event.

**Ceiling for anyone re-proposing it:** the cache can win at most ~0.6 ms per changed cell, so it is
invisible on a one-cell edit and worth ≤5 ms on a nine-cell paste. The 18.7 ms spike needs its own
explanation first — it is not per-event, so it is either the host evicting the table range between
bursts or a proxy being re-created, and that is a different fix.

### G. Manual refresh: compare delete-and-recreate with update (P1, owner decision)

The measured create-from-template cost (2.3–8.6 ms) is below the measured update cost
(9.4–21.7 ms). A full manual rebuild could delete the owned set and recreate it from
templates instead of updating in place. That is `ShapeReconciler` D2's deferred option and
now has the benchmark evidence D2 asked for. It changes shape identity (selection,
comments, accessibility) and must not be adopted without an owner ruling. The `ShapeRange`
route (B) gets most of the win without changing identity and should be tried first.

### H. What parallelisation can and cannot do

- **COM is not parallelisable.** The add-in runs on Excel's STA; `Task.Run`/`Parallel.For`
  around COM violates the contract and cannot speed up a single host.
- **The CPU is not the problem.** `validate + build` is ~0.8 ms of a 1 500 ms rebuild.
  Parallelising the scene build, validation, translation or hashing would save
  microseconds and add risk.
- **What does work in this codebase is batching at the COM boundary**: `ShapeRange`
  multi-shape formatting (B), `Duplicate` template reuse (C), one bulk `Value2` per range
  (already), one lookup per batch (already, ADR-0054), and one `Shapes.Range(...).ZOrder`
  per ordered set instead of one call per shape (needs a probe; the current O5 pass is
  already only 3–27 ms).
- **Legitimate background work:** pre-JIT the Core scene path at add-in load on a
  thread-pool thread (pure CPU, no COM) to remove the cold-start JIT share; the export PNG
  encode (Raster) is already off the refresh path.

### I. Already rejected, with evidence — do not re-propose

- `Application.Calculation = xlCalculationManual`: the property raises
  `DISP_E_TYPEMISMATCH` on this host (ADR-0020, measured twice); it was removed rather than
  worked around.
- Appearance-level skipping on automatic passes: `null` means "leave the host alone"
  (R4.7 D5), so an un-applied style cannot be proven still correct (ADR-0053/0054).
- Caching the configuration read: the VeryHidden sheet is user-editable and unhiding it is
  a supported repair path (P3 rejection in the performance plan).
- Caching the owned set: a tag edit is R9.4's scenario; the reads stay.
- Lowering the cell debounce below 300 ms: risks duplicate passes for asynchronously
  delivered events.
- **Bulk-reading the column flags from one multi-column range.** Probed live 2026-11-20
  (`_artifacts/probe-column-bulk-flags.ps1`, throwaway 3-column table with column 2 hidden
  and unlocked): `DataBodyRange.Locked` returns `DBNull`, `DataBodyRange.EntireColumn.Hidden`
  returns `False` while column 2 IS hidden, `Range.EntireColumn.Hidden` returns `True` only
  when **all** columns are hidden, and `HeaderRowRange.Locked` returns a scalar. The host
  reports an aggregate, not a per-column array, so the ~130 per-column reads in
  `ExcelColumnPresentationRestorer` cannot be collapsed this way. Reducing them needs the
  persisted classification version (owner decision), not a bulk range.

### M. One column-map pass in the type-change normaliser (shipped 2026-11-20)

`ExcelTypeChangeNormaliser` resolved ten schema column names by scanning the whole live
column collection per name: **10 x 26 = 260 COM `Name` reads per refresh**. The owner's
12:2x trace measured `types` at 7.2–26.7 ms (median 12.5) for ten rows of data — a cost that
does not depend on the data at all. `BuildColumnIndexMap` now builds the zero-based map in
one walk (26 reads) and `IndexOf` answers each lookup from it, preserving both the zero-based
offset (a `Value2` matrix column) and first-match-wins semantics. Test:
`The_column_names_resolve_with_one_read_per_live_column`. Gates: `verify-quick.ps1` PASS
242.5 s (Office contract 1626, Core 2154, AddIn 846, Architecture 123, Raster 41, non-live
Office 29); `verify-office.ps1 -DeadlineSeconds 1200` PASS 83/83 in 14 m 59 s.

This is the third instance of the same defect in this codebase (`StagingShapeAdapter`,
`ExcelColumnPresentationRestorer`, now this), and the same fix each time: one walk, one map.

**Post-memo profile, owner rerun 12:20–12:27** (medians of the fifteen converged automatic
passes, 3–9 shapes edited, 148 primitives):

| term | ms before (10:50) | ms now | note |
|---|---:|---:|---|
| `owned` + `verify` | 61.5 | **~45** | the memo: one name and one tag read per shape |
| `columns` | 27.6 | ~38 | 130 per-column COM calls; bulk read disproved (§I) |
| `types` | 7.8 | ~12.5 → the map fix | 260 name reads before the map |
| `grid` + `heights` | 23.5 | ~19 | panel grid measurement |
| protection window | ~51 | ~40 open + ~25 close | owner decision (ADR-0020/0055) |
| `ops` | 1.9–2.6 | 22–78 | real style writes for the 3–9 changed shapes |

Host variance in this session was large (`read` ranged 13.2–150.9 ms for the same table), so
only the term medians are comparable; the `command` figures are not.

### N. Protection close: cycle only when the outline option was actually lost (shipped 2026-11-20)

`ExcelWorksheetProtectionPort.ProtectWorksheet` released and re-protected the sheet on every
close when the sheet was already protected, because `EnableOutlining` is refused while
protected and is not saved with the workbook. The live trace showed that condition is
**always** true on a warm pass (`protectedBefore=True ... cycled=ok` on every converged
close), so every refresh paid an extra `Unprotect` + `Protect` pair for an option that was
already on.

The cycle now runs only when the host reports the option OFF (or cannot report it), so a
reopened workbook still heals exactly as before and a warm pass skips both calls. The trace
distinguishes the three outcomes: `cycled=ok|refused|skipped`. Tests:
`A_protect_that_finds_the_outline_option_on_does_not_cycle_the_protection` (no `Unprotect`
call at all) and `A_protect_whose_outline_option_is_unreadable_still_cycles` (unknown keeps
the healing path). Gates: `verify-quick.ps1` PASS 243.6 s (Office contract 1628);
`verify-office.ps1` PASS 83/83 in 15 m 14 s.

**The other three candidates were assessed against the same evidence, and two are blocked by
facts rather than by risk:**

- **Skip the window entirely when the pass writes no shapes** — not separable yet. The
  window has to be open before the *sheet* phase, because that phase writes outline levels,
  duration cells, StyleKeys and the validation report, and those run before the first shape
  write. A lazy release therefore needs all four sheet-phase writers to report "nothing to
  do" without a host call first. That is the next change, and it is the only remaining item
  worth ~40 ms.
- **`columns` classification skip** — blocked by the probe in §I: the host reports an
  aggregate for a multi-column `Hidden`/`Locked` read, so there is no cheaper read, and the
  pass deliberately runs only on an unprotected sheet (it refuses a protected one), which
  means the `Locked` value it writes is inert at that moment but still the value the user
  gets when they later protect. Dropping the verification would stop healing a user who
  unhid an engine column.
- **Measure/grid cache** — no sound invalidation signal exists: row heights and column
  widths change through user actions with no event the add-in can trust, and the measurement
  IS the thing that would detect the change.

**Next, in order:** (1) skip the verified geometry check when the layout fingerprint is
unchanged — shipped as opportunity O below; (2) make the four sheet-phase writers report "no
work" from the delta they already have, then open the window lazily from the shape writer's
first mutation — worth ~40 ms open + ~25 ms close per pass. Item 2 is larger than it looks:
the window has to be open before the `columns` classification and the `sheet` phase, and both
refuse a protected target, so a lazy release needs those guards to understand "protected by
us, released by us" rather than "protected".

### O. Verified skip: no geometry verification when the layout is unchanged (shipped 2026-11-20)

The reconciler verified every cached, payload-identical shape against the host — 148 x (tag,
placement, four geometry members) — because the payload proves the scene did not change, not
that Excel left the shape where the scene put it. That cost `verify` 12–25 ms on every
converged automatic pass.

**Owner ruling 2026-11-20: an automatic pass must not repair a user-dragged shape.** The
chart's shapes are locked, so a shape that moved means the user unlocked the chart and edited
it; the explicit Refresh is what repairs that.

That ruling collapses the problem: with the shapes locked, the only way Excel moves one is a
change to the rows or columns it is anchored to — an insert, a delete, a hide, a row height, a
column width, or the table origin. Every one of those is an input to
`ShapeDeltaCache.LayoutFingerprint`, built from the same `PanelCellGrid` the scene was laid out
with, plus the count of heights the pass wrote (a pass that wrote one cannot vouch for its own
measurement). Equal fingerprints therefore mean equal anchors, and the verification is
redundant.

`ShapeReconcileRequest.SkipLiveGeometryCheck` is opt-in and the orchestrator sets it only when
the fingerprint matches the one recorded by the last successful pass. The explicit Refresh runs
without a snapshot, so it never sets it and always verifies — the repair path is intact.

Tests: `A_layout_unchanged_pass_skips_the_live_geometry_check` (zero geometry reads),
`Without_the_layout_proof_the_same_pass_still_verifies_geometry` (the same request without the
flag still reads twice — the skip cannot be reached by accident), and
`A_scheduled_pass_with_an_unchanged_layout_does_not_repair_a_hand_moved_shape` at the
orchestrator level, which proves the fingerprint is recorded and matched across two real
passes. `An_explicit_refresh_rewrites_a_shape_whose_live_geometry_drifted` replaces the old
scheduled-path assertion with the explicit path, which now owns that repair.

### P. Is selection- or action-awareness worth it? (investigated 2026-11-20)

**Measured split of click-to-plot** from the 13:3x/13:4x session (automatic passes, converged):

| segment | ms | source |
|---|---:|---|
| analyzer, per changed cell | 2.4–4.9 | `gantt.refresh detect ... detectms=` |
| debounce | **300** (cell edit) / **50** (after a command) | `GanttRefreshScheduler` |
| macro-queue latency | 10.1–115.3 | `gantt.refresh detect ... queuedms=` |
| pipeline | 163.6–1071.6 | `gantt.refresh timings total=` |
| protection window open + ribbon post | ~45–60 | `command − total − post` |

So for a **cell edit** the perceived latency is **300 + queue + pipeline**, and the 300 ms
window is the single largest fixed term — larger than any remaining pipeline phase. For a
**ribbon command** it is already 50 ms, so nothing is owed there.

`events=9–10` on the cell-edit passes shows those bursts are pastes, where the 300 ms coalescing
is doing real work: without it a nine-cell paste would render up to nine times. A
**single-cell** edit (`events=1`) pays the full window for no coalescing benefit.

**Which items need re-checking is already answered, and it is not the selection.** The delta
cache decides per shape, and the phase trace shows the per-row work is already negligible:
`validate` 0.1–0.6 ms, `types` 0.6–1.3, `identity` 0.1–1.5. The expensive terms are
`columns` (~25–75), `owned` + `verify` (~30–45), `measure` (~20), the protection window
(~45–65) and the real shape writes (`ops`) — none of which depend on which cell is active.
Knowing the active row would let a pass skip validating the other rows, worth well under 1 ms.

**Recommendation.** Selection-awareness: **not worth building.** Action-awareness is worth
exactly one thing — an adaptive cell debounce: keep 300 ms for a burst (`events > 1`, i.e. a
paste or fill), and use a short window when the burst is a single cell.

**Implemented, then withdrawn 2026-11-20 (would have amended ADR-0050 D3).** The change was built
exactly as proposed: a **100 ms** window for the first accepted cell change of a burst, escalating to
300 ms on the second, command unchanged at 50 ms. 100 ms rather than 60 because the escalation only
has to win the gap between a paste's events, which the trace shows is ~0.4 ms, and every extra
millisecond of the short window is cheap insurance for a slow host delivering an event late. Tests:
`The_single_cell_window_is_shorter_than_the_burst_window` (policy pin),
`A_single_cell_change_is_queued_on_the_short_window` (its 250 ms deadline sits inside the burst
window, so a scheduler still using 300 ms for one event fails rather than passing slowly) and
`A_second_change_in_the_same_burst_escalates_to_the_burst_window` (the negative assertion is the
load-bearing half). Static gates were green: `verify-quick.ps1` **PASS 235.5 s** (Office contract
1631, AddIn 849). Expected effect: ~200 ms off the perceived latency of a one-cell edit.

**Why it was withdrawn — the live gate, twice.** `verify-office.ps1` failed two *different*
`AddRowIntegrationTests` in successive runs:

1. `Insert_appends_all_three_command_rows_without_rendering` asserts `Assert.Empty(ListOwned())` after
   three inserts. That assertion was always a race — it depended on the test finishing inside the cell
   window — and a 100 ms window loses it. The rendered set is a complete, correct chart (ChartTitle,
   background, bands), so the product is behaving as designed: the assertion is about a command's
   synchronous effect, and the add-in's scheduler is a second actor in the same host.
2. With that test made deterministic, `An_outside_table_selection_appends_and_still_reserves_the_padding_row`
   failed — a *different* test in the same family. That is the decisive evidence: the whole
   `AddRowIntegrationTests` family asserts row arrangement and shape set immediately after an insert,
   and it only passes because a 300 ms quiet window outlasts the test. The debounce was not exposing
   one stale test; it was exposing an unstated assumption shared by a family.

**Three fixes were attempted, and the third disproved the assumption behind the first two.**

1. Set `AutoRefreshPreferences.Shared.AutoRefreshEnabled = false` before the workbook existed: no
   effect — activating a workbook re-seeds the shared switches from the persisted settings.
2. Write the persisted `AutoRefreshEnabled=FALSE` cell and set the shared switch again after
   `Initialise`, before the inserts: still no effect — the add-in's types load from the packed XLL
   into its own load context, so the test process's `AutoRefreshPreferences.Shared` /
   `ProtectionPreferences.Shared` are **not the instances the add-in uses**. Only workbook-persisted
   values cross that boundary, and they are re-read only on activation, which has already happened.
3. Read `ListOwned()` immediately after the **first** insert, microseconds after the write: still
   found the chart already built. That kills "the inserts caused it" — the render comes from the
   test's **setup** (`Initialise` and the workbook activation's chart-affecting cell changes), which
   the old 300 ms window hid by outlasting the whole test. An insert also legitimately triggers a
   further pass, because the new row IS chart data, so no reading position inside that test can
   assert "empty": the chart is expected to exist and is expected to grow.

**What did work, and what it implies.** Dropping the `RegisterXLL` step for that one test makes it
deterministic (**PASS**, 13 s): nothing in it needs the add-in — `Initialise`,
`GanttRowInserter.Insert`, `ExcelShapeWriter.ListOwned`, the TypeOptions defined name and the
Type-column validation are all plain Office/Excel API. The re-apply is therefore one commit plus one
prerequisite change: make the `AddRowIntegrationTests` family scheduler-independent (drop
`RegisterXLL` wherever the add-in is not needed, then check the other Office families for the same
hidden assumption), then land the adaptive window exactly as specified above.

**Shipped 2026-11-20 (second attempt), after the prerequisite.** The prerequisite probed above is now
done for the whole family, not just one fact: `AddRowIntegrationTests` no longer calls `RegisterXll`
at all — nothing in it needs the add-in — and its class remarks record why, following the pattern
already used by `ColumnVisibilityRoundTripIntegrationTests` and `ProtectionArmingIntegrationTests`.
With the scheduler out of that host the product change landed as specified (100 ms first change,
300 ms escalation, 50 ms command, burst counter reset when a burst's timer fires), and both gates are
green: `verify-quick.ps1` **PASS in 238.5 s** (Core 2154, Office contract 1631, AddIn 847,
Architecture 123, Raster 41, non-live Office 29) and `verify-office.ps1 -DeadlineSeconds 1200`
**PASS 83/83 in 15 m 10 s** on Excel 16.0.20527.20196 x64, including the two facts the first attempt
failed.

**Measured profile for the next pass** (owner's 16:37 session, twelve converged automatic passes on
the 148-primitive workbook): `command` 174–379 ms against `total` 100–294; `read` 36–103 (of which
`columns` 28.6–79.8), `write` 34–213 (`ops` 17–107, `owned` 13–40), `measure` 10–40, and protection
window + macro queue 33–103. The column-toggle passes are genuine churn, confirmed from the same log
by `hiddenColumns=`: `[Type,Activity ID]` → `[]` → `[Type,Activity ID,Description,Start,Finish,Duration]`
→ `[]` → `[Type]`, which re-creates 14–28 primitives at 460–1310 ms each. Selection-awareness remains
**not worth building** for the reasons recorded above.

**The instrument this session lost is restored (§A).** The `gantt.refresh detect` line existed in no
commit — `git log -S "detect source"` matched nothing — so reverting `GanttRefreshScheduler.cs` to
`HEAD` during the withdrawal destroyed the only copy, and the unreachable-blob scan found nothing to
recover. It is re-implemented to the field contract §A already fixed: `source=cell|command`,
`events=` accepted changes in the burst, `detectms=` accumulated `ShouldAutoRefresh` cost,
`queuedms=` timer-fire to macro-entry, `scheduled="book|sheet"` captured from the burst's last
accepted address. It is written inside the macro immediately before the pipeline, and a command pass
reports zero analyzer cost with an empty target. Two facts pin it.

**First reading (owner run 18:23).** Five cell bursts: `events=8 detectms=18.8`, `events=1
detectms=0.6`, `events=1 detectms=18.9`, `events=1 detectms=0.3`, `events=4 detectms=18.7`. That is
0.3–0.6 ms per changed cell at steady state, and a recurring ~18.7 ms one-off that does not scale with
the event count — which **falsifies §F** on its own terms (see there). `queuedms` runs median 20 ms,
max 194.7 ms: the macro-queue turn was previously invisible and is now attributable.

**The two remaining command-side terms got their own traces (2026-11-20).** `post` was one number
(17–50 ms) covering two unrelated jobs, so it is split: `gantt.refresh post-steps ms=state=… tabs=…`.
And the protection window's host calls were only inferred from `command − total − post`, so both
window lines now carry `ms=` — separating "the window" from "the ribbon" for the first time. Neither
line changes behaviour; both exist to size the changes below.

**What the three remaining items need, so the next change can start without re-deriving it.**

- *Lazy protection release* (31–54 ms of every converged pass): the window must be open before the
  `sheet` phase, which writes outline levels, `Duration`, `StyleKey`s and the report **before** the
  first shape write. Saving the open/close pair requires each of those four writers to answer
  "nothing to do" from pure inputs — a plan/apply split across four adapters plus the orchestrator.
  It is one change, not a flag.
- *Column classification on the automatic path* — **shipped 2026-11-20, with no persisted stamp.**
  The owner accepted the trade in the words "the quicker option that still looks largely correct,
  where a full refresh can fix it", so `IColumnPresentationPort.EnsureClassification(bool repair)`
  gained a `repair: false` path that returns **before any host call**, and only the scheduled pass
  uses it. The settings fingerprint this section originally proposed turned out to be unnecessary:
  the initialiser applies the classification, every explicit Refresh re-asserts it, and nothing else
  the add-in does can change it — so there is no state to persist that a flag could get wrong.
  Measured cost removed: `columns=19`, `22.4`, `27.3`, `45.1` ms across four passes in the owner's
  20:1x run. Cost accepted: a column the user unhid stays visible until the next explicit Refresh,
  and the chart still renders from the columns it measures.
- *Grid identifier stability* (column-toggle passes at 460–1325 ms): `chart:grid:{x}` renames every
  interior line when the panel width changes, so a toggle deletes and re-creates 14–28 primitives.
  Keying by band ordinal turns that into a geometry update, but identifiers are pinned by ADR-0013
  and the reference golden, so it needs a contract ruling and a reviewed golden regeneration.

## 7. Acceptance criteria for the next step

1. The instrumentation line exists and names every sub-phase above; a no-op pass, a
   geometry-only pass and a create pass are distinguishable from the log alone.
2. P2 is explained: the log names the effective `HiddenColumns` setting, the measured
   column widths, `owned`, and the created/deleted ids for at least one churn sequence.
3. The 20-activity chart is re-measured after B and C on the same workbook and host build,
   with `create-steps`, `write`, `total` and `command` before/after.
4. Contract tests: a counting fake proves the grouped writer issues **one** style call per
   identical appearance group and **one** `TextRange.Text` per text shape; the existing
   per-shape pins for a changed appearance stay green.
5. Live probe: a `ShapeRange` fill/line/text-frame write is read back and equals the
   per-shape values for a mixed set of bars, text boxes and lines.

## 8. Evidence ledger

| Statement | Status | Evidence |
|---|---|---|
| `write` is 86 % of a full manual rebuild | fact | 72 passes, 145 shapes median, log 2026-11-09 → 2026-11-19 |
| CPU (`validate + build`) is ~0.1 % of a rebuild | fact | same log |
| A full update is ~20–40 COM calls per shape | fact | code path, `ExcelShapeWriter.UpdateCore` → `ApplyText`/`ApplyStyle` |
| Measured update cost is 9.4–21.7 ms/shape at 148 shapes | fact | 2026-11-19 manual passes |
| 252 creates cost 807.8 ms of create work | fact | `gantt.shape.create-steps n=252` |
| Template duplicate costs 2.3 ms (rect) / 8.6 ms (text) vs 24.9 / 34.6 ms create | fact (test-recorded host measurement) | `ExcelShapeWriterTemplateTests` remarks |
| Live writer reuses create templates since 2026-11-19 | fact (shipped) | `RefreshSheetCommand.CreateLiveShapeWriter(reuseTemplates: true)`; `RefreshSheetCommandTests` |
| `ShapeRange` applies Fill/Line/TextFrame2 to every shape in the range | fact | Microsoft Learn, ShapeRange; reflection over shipped interop |
| Grouped update writes appearance once per group and content per shape | fact (shipped) | `IShapeWritePort.ApplyGroupAppearance`; `ShapeReconciler`; `ExcelShapeWriterTemplateTests` |
| Grouped update is 5.9× cheaper than the same writes shape by shape (20 text boxes) | fact (live) | `ShapeGroupUpdateIntegrationTests`, Excel 16.0.20527.20196 x64, 2026-11-20 |
| A 20-activity cold refresh is 2.5 s and a warm refresh 0.19 s | fact (owner-measured) | `docs/STATUS.md`, 2026-11-19 |
| Some automatic passes create/delete 14–28 shapes and cost 0.14–3.7 s | fact | `created=… deleted=…` result lines, 2026-11-19 |
| The alternating 134/147 scene is caused by column visibility or a measurement loop | **unknown** | needs the instrumentation in A |
| Batching would take a 250-shape pass to 0.5–1.5 s | inference | call count × measured 0.29 ms/write |
| Parallelising COM would help | rejected | STA rule; single host |

## 9. Live verification (2026-11-19)

`verify-office.ps1 -DeadlineSeconds 1200` — **PASS, 81/81, 14 m 43 s**, Excel 16.0.20527.20196
x64. The default 780 s deadline stopped the first attempt at 22:16 with the suite still
running; 1200 s completed. The new lines were observed on the live host:

| Pass | `command` | `total` | Split |
|---|---:|---:|---|
| cold, 33 creates | 844.5 ms | 561.6 ms | ops 95.9, heights 81.6, grid 44.6, config 25.9, report 30.4, outline 24.4, identity 12.9, types 11.7; create-steps add 59.2 / style 21.8 / ownership 8.7 |
| cold, 106 creates | 958.0 ms | 656.1 ms | ops 221.8, heights 68.5, grid 42.7, config 27.5, report 51.3, outline 24.9, request 32.0; create-steps add 132.7 / style 58.3 / ownership 23.4 |
| warm, 2 updates | 38.6 ms | 34.4 ms | ownedms 2.2, verify 5.1, ops 9.9 |
| edit burst | — | — | 8 changed cells, detectms 18.7, queuedms 4.3 |
| 106-shape cold create, before C | 958.0 ms | 656.1 ms | style 58.3, add 132.7, ownership 23.4; write 245.4 |
| 106-shape cold create, after C | 819.7 ms | 559.9 ms | style 12.4, add 87.7, ownership 17.0; write 152.6 |

The 8-event burst is the pre-debounce blind spot measured: ~2.3 ms per changed cell plus
4.3 ms of macro-queue latency. The warm pass confirms the floor: 34.4 ms pipeline and
38.6 ms command. The 106-shape cold create is the largest live pass the suite produced;
the retained 252-create pass remains the 20-activity-scale evidence, and opening the
20-activity workbook itself is still owed.

**Grouped update (2026-11-20).** `dotnet test ... --filter ShapeGroupUpdateIntegrationTests`
— **PASS, 2/2**. The first fact creates three rectangles and two text boxes, applies one
grouped style pass and one grouped typography pass, then one `UpdateContent` per label, and
reads every fill, stroke, width, font size, alignment, text colour, string, name and
ownership tag back off the host. The second fact measures, inside a batch, 20 text boxes:
**per-shape appearance 454.2 ms vs grouped 76.7 ms, 5.9×**. Both facts run in
`verify-office.ps1`.
