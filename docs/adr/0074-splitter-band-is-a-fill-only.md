# ADR-0074 — The §10 splitter band is a fill only

- **Status:** Accepted (owner ruling 2026-11-27)
- **Supersedes:** ADR-0022 D3 and D4
- **Related:** entity guide §10, ADR-0027 (`SceneLine` users), ADR-0052 D11 (host-owned rules), R11.1

## Context

ADR-0022 gave a `Splitter` a filled band plus two `MajorBoundaryPt` lines at `ZLayer.Frame` along
its top and bottom edges, so the section header read as a framed stripe. The owner's 2026-11-27
ruling is that the fill is the entity: the two rules and the framing they create are not wanted.

## Decision

- **D1 — The entity emits one rectangle and no lines.** `SplitterBuilder` emits a `SceneRect` at
  `ZLayer.Section` (30) plus zero, one or two labels. `SceneLine` is no longer part of §10.
- **D2 — The border inputs are removed, not ignored.** `SplitterRequest` loses `BorderStyle` and
  `BorderWidthPt`; their validators and their tests go with them. A retained-but-unused input
  would make its guard unreachable, which `AGENTS.md` treats as a defect.
- **D3 — The band still spans the plot only.** ADR-0068 D1 is unchanged; `PanelLeftPt` remains the
  `DataPanelLeft` label field.
- **D4 — The rule retires from the host-owned list.** ADR-0052 D11 no longer names the splitter's
  rules among the host-owned members. The predicate is geometry-based and needs no code change.

## Consequences

- A section header is a `#FFE699` band. On the live sheet its separation is the fill alone; in an
  editable copy the table's own row borders run beside it.
- `MajorBoundaryPt` remains consumed by the frame, the plot's four sides, the header cell edges and
  the closing line, so no metric becomes inert.
- The scene golden loses two primitives (recorded in the R11.1 golden commit).
