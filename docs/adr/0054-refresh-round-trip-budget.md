# ADR-0054 — Spend the refresh on writes, not on round trips

## Context

ADR-0052 (verified skip) and ADR-0053 (command-scoped geometry) removed most redundant
*writes*. What remains is dominated by *reads and resolutions that return the same answer for
every shape in a pass*, and by commands issued for shapes that are already correct.

The 2026-11-04 live sessions changed the cost model. Earlier notes treated a read and a write
as comparable; the zero-write passes disprove it. Within a single quiet session, on 86–87
shapes:

| Operation | Observed cost |
|---|---|
| property read on a held RCW (`Left`, `Top`, `Name`, `AlternativeText`, `Placement`) | ~0.03–0.1 ms |
| property write / shape command (the host re-layouts) | ~0.5–1.2 ms |
| collection walk step (`Sheets`, `ListObjects`, indexed `Shapes(i)`) | ~0.2–0.5 ms |
| `ZOrder` command | ~0.6–2.7 ms |

So a pass of 39–42 relocated primitives paid roughly 14–16 round trips each, of which only 5
were writes. Read-to-avoid-a-write is a good trade; read-to-avoid-a-read is not.

Per-shape redundant work at the time of writing:

- `ResolveShapes` walked `Sheets` and `ListObjects` and re-queried the guard **per operation**
  (≈2 walks + 1 query + the leading guard query = ~6 round trips);
- the z-order preflight looked every shape up by name **again**, after the update/geometry pass
  had already resolved it;
- `Update` wrote `Placement = xlMove` on **every** shape on **every** pass, almost always
  writing the value the shape already held;
- `ListOwned` read `AlternativeText` for **every** shape on the sheet, most of which are a
  user's drawings and cannot be ours;
- `ApplyZOrder` sent **every** shape to the back whenever any order changed, ~2–3 round trips
  per shape, when typically 2–4 shapes were out of place.

## Decision

Five changes, all within the existing ports and all covered by contract tests.

1. **One batch per reconcile.** `IShapeWritePort.BeginBatch()` returns an `IDisposable`;
   `ShapeReconciler.Reconcile` opens it once (after the blocking-error check, so a blocked
   refresh still issues no host call) and every pass — verification reads, plan execution,
   z-order — runs inside it. `ExcelShapeWriter` resolves the worksheet/`Shapes` collection once
   and closes it on dispose, on **every** return path. Batches do not nest.

2. **Proxy reuse.** The batch memoises each successfully resolved shape proxy by name;
   `ApplyZOrder`'s preflight reuses it and only falls back to the indexer on a miss. `Create`
   adds the new proxy, `Delete` removes it. The memo is batch-scoped, never cross-refresh.

3. **Conditional placement write.** `Update` reads `Placement` and writes `xlMove` only when
   the host reports something else, so a legacy `xlMoveAndSize` shape still heals on the first
   pass after the rule changed, and no other pass writes at all.

4. **Name-prefiltered ownership scan.** `ListOwned` reads `AlternativeText` only for a name
   that could be one of ours: a role-derived `<ownerId>:<role>` (`ScenePrimitive.CreateId`
   always emits a colon) or the fixed chart-title identifier. Because that is an assumption,
   it is enforced rather than documented: `Validate` refuses an identifier outside those forms
   (`ShapeWriteRefusal.UnsupportedIdentifierForm`), so a shape that a refresh could never
   recognise is never created. The chart title's identifier is now one constant,
   `ExcelShapeWriter.ChartTitlePrimitiveId`, consumed by the orchestrator.

5. **Minimal z-order moves.** `ApplyZOrder` reads `ZOrderPosition` for the desired
   back-to-front list, accepts the longest already-correct **front** run, and sends only the
   shapes behind it to the back, front-to-back, which is the original absolute pass applied to
   a prefix. An already-correct stack issues no command. A host that reports no usable position
   (`<= 0`) falls back to the full pass, so the pass still converges.

   Two facts decide that shape, and the first implementation of this ADR got both wrong:

   - **`Shape.ZOrderPosition` 1 is the shape at the BACK** of the z-order (Microsoft's
     `Shape.ZOrderPosition` remarks: `Shapes(1)` is the back, `Shapes(n)` the front), so a
     correct back-to-front run has strictly *increasing* positions.
   - **`msoSendToBack` is absolute**, so every shape the loop touches finishes behind every
     shape it does not. Only a **front** run can therefore be left alone; skipping a back run
     sends the moved shapes behind it and inverts exactly what was skipped.

   The defect this replaces shipped in `c55dda0` and mis-stacked the chart on both the
   automatic and the full refresh: it skipped a back-most prefix on a decreasing-position test.
   It was invisible to the original tests because the double handed each shape a *fixed*
   position, so the assertions could only check which `ZOrder` calls were made — never the
   order the user sees. The double now models the collection and applies `msoSendToBack` for
   real, and a permutation test asserts the resulting stack for all six 3-shape orders.

The guard consultation stays visible in each entry point (`Create`, `Update`, `Delete`,
`ApplyZOrder`) as ADR-0008 D4's first statement and is short-circuited only **inside** a
batch, because the batch's target resolution already performed the per-target check once and
events are off for the whole pass (ADR-0020). Keeping it inline is deliberate: the
`ProtectionGuardFirstTests` gate inspects the entry method's own tokens, precisely because a
guard moved into a helper once became unreachable.

## What this does not change

- Per-operation validation, request checks and the ownership-tag comparison are untouched: a
  batch removes repeated *answers*, not checks.
- Refusal outcomes are unchanged, and a refusal still stops the pass and disposes the batch.
- The explicit Refresh still plans a full update for every shape (no previous snapshot).
- No appearance-level skipping: `null` means "leave the host alone" (R4.7 D5), so an
  un-applied appearance part cannot be proven (ADR-0053's rejected alternative stands).

## Consequences

- Expected within-session effect (quiet host, 86–87 shapes): a relocation pass from
  ~8–12 ms/primitive toward 4–6 ms; `zorderms`/shape from 0.7–2.0 toward 0.4–1.0; a small
  reorder (2–4 moved shapes) from ~150–250 ms toward 60–100 ms. Cross-session averages remain
  unusable — the host factor measured at 1.5–3.5× dwarfs the code's effect.
- The writer now has per-pass state (`_batchActive`, `_batchShapes`, `_batchProxies`), so it is
  no longer reentrant. That is safe by construction: the refresh runs on the Excel main thread
  with events off, and `BeginBatch` throws on a nested open rather than silently releasing the
  outer pass's resolution.
- A new named identifier outside the two enumerated forms now fails loudly at `Create` instead
  of producing an un-updatable shape. That is the intended failure direction.
- Z-order tests assert the **resulting stack**, not the call list: a call list cannot tell a
  correct pass from an inverted one, which is exactly how the first version of O5 survived its
  own tests. Every 3-shape starting permutation is pinned, plus the reversed and
  unusable-position cases.

## Validation

Contract tests (no host): batch opened once and disposed on success and on refusal, none for a
blocked refresh; proxy reused by the preflight; placement not rewritten when already `xlMove`
and healed when not; the two accepted identifier forms and two rejected ones; the tag read
count for a user shape; z-order for already-correct, partially-correct, front-anchored and
unusable-position stacks.

Live gate: one Excel session with the same action script, measured with
`scripts/_artifacts/analyze-refresh-log.ps1`, comparing **within-session** per-primitive and
z-order ratios only. `Not run` at the time of writing — no host session in this change.
