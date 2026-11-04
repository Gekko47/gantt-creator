# ADR-0053 — Command-scoped shape updates: a shape that only moved is moved

## Context

A refresh reconciles the scene against the owned shapes and, after ADR-0052, skips a shape
whose payload is unchanged **and** whose live geometry still matches. That leaves one large
class of write: a shape whose payload changed only because its **box** changed.

The live profile (2026-11-03 and 2026-11-04 sessions, 86–87 shapes, 11 rows) measured it:

| Pass | Writes | Cost |
|---|---|---|
| automatic, 6–12 changed shapes | 6–12 | 150–250 ms |
| automatic, zero changed shapes | 0 | 31–67 ms |
| promote/demote (relocation) | 39–42 | 257–335 ms |
| manual Refresh (full repair) | 86–87 | 859–1 129 ms |

Marginal costs: **~15–18 ms per changed shape** for a full update, **~0.7 ms per shape** for
the ordering pass, and a **~20 ms** fixed cost inside the reconcile. A relocation command
writes 39–42 primitives because the row content rotates: every primitive in the moved window
changes its `Top` (and nothing else), and each is currently re-styled and re-lettered to get
there.

## Decision

Split the payload into two hashed parts and drive the write from the diff:

- **geometry** — `Kind` and `Geometry` (the box);
- **appearance** — fill, stroke, line width, hatch, dash, font family/size/bold, text,
  alignment, text colour.

`ShapeDeltaCache` stores both hashes per primitive alongside the full-payload hash. When a
refresh finds a shape whose geometry hash changed and whose appearance hash did **not**, the
reconciler plans a `ShapeUpdateScope.Geometry` operation, and `IShapeWritePort.UpdateGeometry`
applies placement plus the box only. Everything else keeps the full update:

- appearance changed (style key, type change, label change) → full update;
- both parts changed → full update;
- no previous snapshot (**the explicit Refresh**) → full update, which is what makes the
  manual path the repair pass;
- kind changed → the scene's kind decides `AddShape`, so no geometry-only path is offered.

Before a geometry-only write is planned, the reconciler asks
[`IShapeWritePort.GeometryMatches`](../../src/GanttCreator.Office/IShapeWritePort.cs) whether the
host already holds that box. An inserted or deleted worksheet row moves every `xlMove` shape
with its cells, so the answer is often **yes** — and then no write is issued at all. That is
why adding a row costs the existing shapes nothing.

## What this trades

A shape the user recoloured by hand and then moved keeps the hand colour until an explicit
Refresh. This is accepted, explicitly, by the product owner (2026-11-04): the automatic path is
there to capture what changed, the manual Refresh is the full repair, and the automatic path
should not pay 15–18 ms per primitive to re-assert formatting that did not change.

The alternative — verifying appearance before skipping it — was rejected in the performance
plan: `TextColour`, `FontFamily` and friends use `null` to mean "leave the host alone" (R4.7
D5), so intent is unobservable after a write and a read-back cannot prove the shape still
matches. Geometry has no such member, which is exactly why ADR-0052 could verify it and why
this ADR can safely move it.

## Consequences

- A relocation, a row-height change, or a date edit inside the existing plot scale costs a
  handful of property writes per primitive instead of a full re-apply.
- The explicit Refresh remains the only path that re-asserts formatting, and it already did
  (it carries no previous snapshot).
- The classification is pure, in-memory, and testable without Excel: the plan carries the
  scope, so a contract test can assert that a move plans `Geometry` and a restyle plans `All`.
- A future change that adds a host-visible member to `OfficeShapeRequest` must decide which
  part it belongs to; adding it to neither (like `ZLayer`, which orders the plan) means a
  change to it alone still produces a full update, which is safe.
