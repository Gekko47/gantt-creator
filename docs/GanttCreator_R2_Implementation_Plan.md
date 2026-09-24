# R2 findings implementation plan

> Approved by the product owner on 2026-09-23 from review of
> [`GanttCreator_R2_Implementation_Findings.md`](GanttCreator_R2_Implementation_Findings.md).
> This plan reconciles that R2.6-era review with the current `stage-inspect`
> source and is subordinate to `AGENTS.md`, the roadmap, architecture, and the
> entity guide.

## Outcome

Harden the landed R2 reader/validator/reporting boundary, complete the
configuration and authoring foundation, then use validated `GanttEvent`
objects as the only schedule input to the deterministic scene and first live
Excel renderer. Do not rewrite R2 or pull exports, recovery UI, localisation,
or advanced controls forward.

## Binding decisions

1. `GanttEntityType` is the durable machine identity. The 16 exact display
   names remain persisted workbook schema values. A display-name migration does
   not renumber or rename the enum identity; no localisation is implemented now.
2. `Custom Activity` requires a known named style. The named style explicitly
   carries `DefaultLabelPosition`, `AllowedLabelPositions`, and
   `ColourCapability`; no renderer invents a fallback.
3. 1904 workbooks remain rejected by ADR-0006. Serial conversion and the
   refusal move behind a dedicated date-system abstraction.
4. Excel error and unsupported values remain distinguishable from blank values
   in a neutral Core cell-state model; COM error enums never enter Core.
5. Relevant invalid values are blocking errors. Populated irrelevant values
   produce one warning and are normalised out of `GanttEvent`.
6. A Critical Interval must reference a valid, mapped span event. R3 owns
   geometric intersection/clipping; R2 owns reference validity only.
7. Initialise adopts only a pristine worksheet. A cell-empty worksheet with
   shapes, comments, names, tables, pivots, queries, hyperlinks, or formatting
   beyond `A1` is preserved and Initialise creates a new Gantt sheet.
8. Initialise never repairs. Read/Refresh paths locate the table by identity,
   not sheet name. R2.10 is an explicit repair command.
9. R2.8 generates IDs in the normal authoring flow. IDs remain visible core
   data; missing/duplicate repair remains R2.10.
10. `GanttValidationIssue[]` is authoritative. Classic Notes are one reporter,
    and the reporter consults `IWorksheetProtectionGuard` before mutation.
11. R2.8 canonical types are `As-Planned Activity`, `As-Planned Milestone`, and
    `Delineator`.
12. R2.10 and R3.13 are mandatory but execute after R4.9 so the first usable
    live Gantt is not blocked by recovery UI or mutation-test infrastructure.

## Implementation order

| Order | Work item | Outcome |
| ---: | --- | --- |
| 1 | R2.1a | Pin stable Type identity and clarify 16 selectable Types versus 26 visual sections |
| 2 | R2.2a | Restrict worksheet adoption to pristine sheets |
| 3 | R2.4a | Isolate date-system support/refusal behind an abstraction |
| 4 | R2.4b | Preserve neutral Excel error/unsupported cell states through the DTO |
| 5 | R2.5a | Make validation/normalisation and parent-reference rules explicit |
| 6 | R2.6a | Make Notes a guarded, replaceable reporter |
| 7 | R2.7b | Materialise style label/colour capabilities under ADR-0009 |
| 8 | R2.8 | Add generated-ID activity, milestone, and delineator row commands |
| 9 | R2.9 | Materialise Type options and registry-aware validation |
| 10 | R3.1-R3.12, R3.14 | Build and prove the deterministic scene without Office |
| 11 | R4.1-R4.9 | Render and explicitly Refresh the first live Excel Gantt |
| 12 | R2.10 | Add explicit safe configuration/identity repair |
| 13 | R3.13 | Add the approved mutation-test harness |
| 14 | R4.10 | Profile the 1,000-event scene/refresh path |

R3.13 and R2.10 retain their stable IDs; ADR-0010 records the sequencing
exception to the roadmap's normal table order.

## Work-item details

### R2.1a — Type identity contract

- Keep the enum and display names as separate surfaces.
- Add a numeric/name pin for all 16 enum members.
- Update XML docs and terminology to say “16 selectable Types” and “visual/entity
  guide sections”.
- No Type additions, localisation, parser changes, or schema-version bump.

### R2.2a — Safe worksheet adoption

- Replace cell-content-only blankness with a pristine-sheet decision.
- Check `A1`-only used range plus shapes, legacy/threaded comments, worksheet
  and workbook names, tables, pivot tables, query tables, and hyperlinks.
- Non-pristine but cell-empty sheets take the create path unchanged.
- Existing managed conflicts remain typed zero-mutation refusals.
- Contract and live tests cover each artefact category.

### R2.4a — Date-system boundary

- Introduce an Office-free date-system kind/converter contract.
- Keep Windows 1900 conversion and Mac 1904 refusal under ADR-0006.
- The reader checks support before reading the table body.
- Preserve serial 60, culture invariance, and existing live refusal evidence.

### R2.4b — Neutral cell state

- Introduce `GanttCellState` (`Empty`, `Value`, `ExcelError`, `Unsupported`),
  typed `GanttCell<T>`, and a neutral Excel error-code enum.
- Convert every `GanttRowDto` field to a typed cell.
- Preserve known and future Excel errors instead of flattening them to null.
- R2.5a consumes the states; no geometry or rendering is added.

### R2.5a — Validation normalisation

- Derive field applicability from the one Type catalogue and date mode.
- Add blocking codes for relevant Excel errors, unsupported values, and invalid
  parent rows; every validator has a positive test.
- Irrelevant populated values warn once and normalise to null.
- A Critical Interval parent must be a valid mapped span event.
- Keep all-errors and `Severity -> Row -> Field -> Code` ordering.

### R2.6a — Reporter boundary

- Inject `IWorksheetProtectionGuard` into the Notes reporter.
- Add a typed protected-target refusal before any note mutation.
- Extend architecture mutation discovery to note writes and register the
  reporter as a mutating adapter.
- Preserve user notes and next-run repair of partial owned-note writes.

### R2.7b — Style capability schema

- ADR-0009 appends `DefaultLabelPosition`, `AllowedLabelPositions`, and
  `ColourCapability` to `tblGanttStyles`.
- Built-in values project from their owning Type contract; user rows must carry
  and validate the same fields.
- Schema version remains 1 because configuration v1 was not shipped before
  R2.7 materialised it; old development workbooks are pre-release drift.
- Update catalogue hash, writer, reader, preservation, and save/reopen tests.

### R2.8 — Add-row commands

- Use the approved canonical types, generated IDs, header-name mapping, and
  body-end append.
- Mutate only the visible table, never render, and consult the shared guard.
- Re-probe appended-row Type validation in R2.9.

### R2.9 — Type options and style registry

- Create `GanttCreator.TypeOptions` only from materialised catalogue cells.
- Build a pure style registry from validated configuration.
- Reject unknown style keys. Custom Activity uses style-defined label/colour
  capabilities; all other Types use `EntityTypeCatalog` capabilities.
- The dropdown remains a convenience; pasted unknown Types remain blocking.

### R3 — Scene engine

- Every builder consumes validated `GanttEvent` and resolved configuration.
- R3.8 alone owns critical parent-span intersection/clipping warnings.
- R3.12 owns one deterministic reference fixture, scene validator, snapshot,
  and benchmark. The fixture covers planned/actual/baseline overlap, critical
  children, delay, procurement, milestones, stacked events, delineators, row
  overrides, warnings, and parent/child identity.

### R4.1-R4.9 — First live Excel Gantt

- Renderers translate scene primitives only; they never reinterpret rows.
- Refresh is explicit. Cell, Type, selection, colour, and label edits do not
  invoke rendering.
- Blocking Refresh preserves the last valid owned scene.
- R4.9 is the first live product demonstration and the prerequisite for the
  deferred R2.10/R3.13 work.

## Canonical fixture decision

R3.12 owns one deterministic reference fixture covering the entity-guide
minimum plus the findings brief's construction-delay cases. Focused tests remain
separate. The Office integration suite may materialise the same logical fixture
in a workbook, but no renderer or Office integration invents alternate schedule
semantics.

## Verification

Each work item runs its targeted tests, Release build with warnings as errors,
`verify-quick.ps1`, and `git diff --check`. Office-host gates are Required for
R2.2a, R2.4a, R2.4b, R2.6a, R2.7b, R2.8, R2.9, R2.10, and applicable R4
rows. `verify.ps1` runs after the final commit of the implementation branch.

## Exclusions

No Office/COM types in Core; no second helper worksheet; no schedule data in
`_GanttCreatorConfig`; no shape source of truth; no live redraw; no row-number
identity; no parallel Type catalogue; no 1904 support; no localisation; no
PowerPoint/PNG work; no new dependency; no migration or repair hidden inside
Initialise.

## Definition of complete

The plan is complete when every row above has a landed work-item guide, code,
same-change tests, current `docs/STATUS.md`/`docs/REPO-MAP.md`, applicable Office
evidence, and green branch-final verification. The findings brief must then match
the landed disposition with no stale “current implementation” claims.
