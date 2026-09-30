# Pre-R4.9 decision register

> **Authority.** This register records the product owner's resolutions of the
> proposals in [`GanttCreator_R3_R4_Architecture_Roadmap_Review_REV5.md`](GanttCreator_R3_R4_Architecture_Roadmap_Review_REV5.md)
> and [`GanttCreator_Revised_Implementation_Plan_REV6.md`](GanttCreator_Revised_Implementation_Plan_REV6.md),
> decided 2026-09-29. Where either document proposes something this register
> rejects, **the register wins**, and an agent must not implement the rejected
> proposal. This exists because REV5/REV6 remain readable and several of their
> proposals were changed; without this file, following them literally is the
> most likely way to reintroduce the drift it was written to remove.
>
> **Not a substitute for.** The binding contracts are the entity guide, the
> ADRs, and the work items. This register is the bridge between the two input
> documents and those artefacts.

## 1. How REV5/REV6 enter the process

The product owner directed on 2026-09-29: *"REV5 & 6 are new additions to the
process to enhance and refine the functionality of the code."* They are
therefore **tracked inputs**, added to `AGENTS.md`'s source-of-truth order
below the entity guide and above the roadmap. They do not outrank the entity
guide — which is why decisions 1, 2 and 4 below require guide revisions rather
than guide overrides.

## 2. Resolved decisions

| # | Decision | Status of the REV5/REV6 proposal |
| --- | --- | --- |
| D1 | **Fixed row heights.** `GanttRowHeightPt` is the sole authority for both Excel row height and Core lane height. Live lane auto-growth is removed. | **Accepted**, and it contradicts landed code |
| D2 | **Critical Interval is a real child row** with its own `Start`/`Finish`, its own styling, and no date inheritance. | Accepted |
| D3 | **Keep the existing unique shape key.** `PrimitiveId` remains the `Shape.Name` and the reconciliation key. | **Rejected**: the `HostShapeKey` proposal |
| D4 | **`Above`/`Below` label positions are removed.** | Accepted, with one retention |
| D5 | **The critical overlay is a filled rectangle**, not a line. | **Rejected**: the `SceneLine` proposal |
| D6 | **`Right` is the preferred label position**, then largest-space-that-fits, then the existing truncation, then `None`. | Partially accepted; see §3.2 |

### 2.1 D1 — fixed row heights

REV5 §9 and REV6 §10 require auto-growth removal. This **reverses landed
behaviour and a checklist invariant**: `LaneLayoutBuilder.cs:170` computes
`Math.Max(metrics.LaneHeightPt, contentHeight)`, and
`08-TEST-CHECKLIST.md` §B states *"Lane height grows to fit content
(`max(LaneHeightPt, contentHeight)`); events are never compressed."* The guide
also states the content-height rule at §411. The owner's decision is
authoritative; ADR-0026 carries it and the checklist item is replaced by exact
worksheet-row equality.

### 2.2 D2 — Critical Interval is a real child row

Uncontested, and already partly true: `GanttEntityType.CriticalInterval = 5`
exists, `EntityTypeCatalog` defines it, and `GanttRowValidator` already requires
`ParentId` for it. What is new is the authoring surface — a child row a user
creates through Add Child — and independent dates.

### 2.3 D3 — the shape key stays as it is

REV5 §16 and REV6 §16 propose separating the logical `PrimitiveId` from a
bounded `HostShapeKey`. **Rejected.** R4.7's ledger already records this as
rejected on evidence, closed as L18, with the product owner approving on
2026-09-28: a bounded `Name` would stop `Name` being the reconciliation key,
and `ExcelShapeWriter.Create` already refuses an identifier past
`MaxShapeNameLength = 255` as a typed `ShapeWriteRefusal`, so the failure mode
is "this entity does not render and says why". There are no active users, so
no migration is required or written.

### 2.4 D4 — label positions

`Above` and `Below` are removed from `GanttLabelPosition` (currently
`Above = 5`, `Below = 6`), from the `EntityTypeCatalog` capability sets, from
`LabelPlanner`'s milestone order, from entity guide §616-617, and from
`08-TEST-CHECKLIST.md` §C — in one commit. The horizontal set retained is
`None`, `Auto`, `Left`, `Right`, and `Inside` (surfaced in the UI as
**Centre**). The four delineator corner positions (`TopLeft`, `TopRight`,
`BottomLeft`, `BottomRight`) and the splitter positions (`DataPanelLeft`,
`PlotCentre`, `Both`) are **retained**, because the landed R3.10
`DelineatorBuilder` and the splitter band depend on them.

### 2.5 D5 — the critical overlay is a filled rectangle

REV5 §7 and REV6 §9 require *"a truthful line-like scene primitive. Do not emit
a thin rectangle that each renderer reinterprets."* **Rejected**, on the
owner's stated grounds that a rectangle is easier to size.

The resolved representation (**as amended by the owner ruling of 2026-09-30**):

- the scene primitive **stays `SceneRect`**, as `CriticalOverlayBuilder`
  already emits;
- the rect is **filled**, using a new **`CriticalFill`** colour token carrying
  **`#FF0000`** — the same red as the existing `CriticalStroke` token at
  `GanttCatalogues.cs:164`;
- the rect's **height is half the predetermined `ActivityHeightPt`**, and it is
  **centred on its own visual slot**, exactly as an ordinary span bar is. It is
  **not** half the parent bar's resolved height and **not** top-aligned to the
  parent;
- the rect's **horizontal extent comes from its own Start and Finish**, clipped
  only to the **plot**. It may be shorter than, equal to, or longer than its
  parent, and an interval lying wholly outside its parent is still drawn — the
  parent is never read for geometry at all, and governs lane membership and
  nothing else. The link survives only through R4.7B's lane projection;
- the **`CriticalLinePt` metric is retired**. The `CriticalInterval` preset
  keeps its `ActivityHeightPt` reference, which supplies the half-height
  denominator;
- `EntityTypeCatalog.cs:132` currently grants `CriticalInterval`
  `EntityColourCapability.Stroke` only, and `GanttRowValidator` plus checklist
  §C both treat a fill override as a blocking error. **All three sites — the
  catalogue, the validator, and the checklist — change in the same commit**, or
  the entity renders filled while the validator rejects the user's own
  override.

**Superseded on 2026-09-30, and the reason is recorded here because two documents
still carried the old rule.** This section previously read *"the rect's height is
half the parent activity bar's resolved height, and it is top-aligned on the
parent's top edge … the builder already receives `ParentVisibleBounds`"*. That was
the geometry in force when this register was written, and the owner later replaced
it. `CriticalOverlayRequest` no longer carries `ParentVisibleBounds` at all, and
`TryBuild` no longer takes a parent-bounds dictionary. A parent-dependent overlay
made the entity's geometry depend on a row it did not own, which is what made
`CriticalIntervalOutsideParent` — and the silent deletion of an interval falling
outside its parent — possible in the first place.

A note on the drift this resolves: the guide previously said all three
renderers draw a *line* from this rect, but **no renderer implemented that** —
`SceneShapeRenderer` has no critical branch, and the PowerPoint (Phase 7) and
raster (Phase 8) renderers do not exist. Zero of three matched the guide. The
filled rect removes the reinterpretation entirely, so Excel, PowerPoint and
raster can each draw the primitive they are given.

### 2.6 D6 — label cascade

The owner's ruling: *"Preference order is Right, then auto based on largest
space available that fits the text string. and then none for those that have
none."* See §3.2 — the existing truncation is retained, not replaced.

## 3. Proposals deliberately not implemented

An agent reading REV5/REV6 will encounter these. They are recorded here so they
are not implemented by mistake.

### 3.1 A new `HostShapeKey`

Rejected — see §2.3.

### 3.2 Replacing the truncation fallback

An earlier draft of this plan read the owner's cascade as dropping ADR-0015's
D4/D5 widest-gap truncation. **That reading was wrong and is withdrawn.** The
owner confirmed the existing function is reused: *"[t]here is a function already
for this, it will truncate if there is no space, nothing new needs to be
built."*

`LabelPlanner` already implements the whole cascade:
`TruncatedToFitCode = "LabelTruncatedToFit"` (line 130), `WasTruncated` on the
result (65, 71), `_spanAutoOrder` already beginning at `Right` (139-141), and
the `FreeRight` / `FreeLeft` occupancy measurement (658, 685). **ADR-0015 is
therefore not superseded.** ADR-0028 amends it on label *positions* only and
leaves D4, D5 and D6 intact.

### 3.3 `CriticalOverlayHeightRatio` as a configurable ratio token

Not introduced as a user-facing setting. The half-height is derived from the
predetermined `ActivityHeightPt` of the `CriticalInterval` preset itself, which
needs no new token; see §2.5. The parent is never read for geometry at all — it
governs lane membership and nothing else.

## 4. No migration work

The product owner directed: *"No need to migrate anything as there are no
active users, this is still in development."* Consequences:

- `GanttTableSchema.Default` is edited **in place** to add the `Duration` and
  `SiblingOrder` columns;
- no legacy-workbook upgrade path is written or tested;
- `GanttSchemaVersion.CurrentSchemaVersion` still advances **3 → 4**, because
  the column schema, the metric-token set, the style catalogue and the
  type-catalogue capability metadata are all schema contracts under ADR-0007,
  and an unchanged version would make the change invisible to config-integrity
  checks.

## 5. Fixture and golden consequences

**Corrected 2026-09-30: the premise of this section was false and is withdrawn.**
It previously read that `tests/golden/scene/reference-scene.json` "contains **no
critical overlay primitive**" and that its only `critical` occurrence was the
warning `CriticalIntervalOutsideParent`. That was wrong on both counts. The
snapshot **did** carry a critical primitive, and the second interval had been
*silently deleted* — not absent, deleted — for falling outside its parent span.
The golden was the evidence that the drift existed, not evidence that it was
harmless, and reading it as "critical-free" is how the divergence survived.

The committed golden now carries **two** critical primitives
(`…000c1:critical` and `…000c2:critical`), regenerated in `806a632` after the
parent-independent geometry landed in `017199f`. The second one came back because
removing the parent clip stopped deleting it. Both are **filled rectangles at
half the predetermined `ActivityHeightPt`, centred on their own visual slot, and
derived from their own dates** — the representation the register's §2.5 now
records.

R4.7E's **colour-capability** change (the later commit) did not regenerate the
golden: it adds a critical child to the **vertical integration fixture** rather
than the canonical fixture, which pins the fill contract without touching the
snapshot. The distinction matters because two different claims were conflated —
"the geometry change did not move the golden" was never true, and "the
capability flip did not move it" is.

Expected golden regenerations across the whole pre-R4.9 sequence: **three** — R4.7B,
R4.7D, and R4.7E's geometry change. Each requires human review and a stated
reason, and each ships in the same commit as its code change, per R3.17's recorded
D3 precedent.

## 6. Roadmap rows this register drives

Defined in [`03-ROADMAP.md`](03-ROADMAP.md) revision 10 and implemented from
[`work-items/`](work-items/): R4.7A identity/hierarchy · R4.7B projection ·
R4.7C workbook presentation · R4.7D row geometry · R4.7E critical interval ·
R4.7F duration and date-only semantics · R4.7G label positions · R4.7H size
presets · R4.8A orchestration. All must complete before **R4.9**, which becomes
integration only and must invent no architecture.

Decisions are carried by ADRs **0025** governance, **0026** fixed row geometry,
**0027** critical interval representation, **0028** label positions, **0029**
schema version 4.
