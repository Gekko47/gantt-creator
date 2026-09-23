# ADR-0007 — Configuration-sheet catalogues: tables, settings home, and catalogue hash

- **Status**: Accepted
- **Date**: 2026-09-21
- **Context**: Work item R2.7 must materialise the complete first-release
  catalogue on `_GanttCreatorConfig` before Phase 3 reads it
  (roadmap Phase 2 note). Three open questions required human approval
  (manifest decision D-G1): the storage format, the exact table set, and
  the home for the settings catalogue (the architecture contract listed
  `tblGanttTypes`, `tblGanttStyles`, `tblGanttMetrics`, `tblGanttConfig`,
  and the `GanttCreator.TypeOptions` name, but no settings table). The
  architecture allows "single JSON or table block"; the roadmap requires
  settling table names before the first write.
- **Decision** (human-approved 2026-09-21):
  - **D1 — Excel Tables, not JSON.** Each catalogue is a `ListObject` on
    `_GanttCreatorConfig`. Rationale: R2.9's in-cell `Type` dropdown needs
    `GanttCreator.TypeOptions` to refer to a materialised cell range (a
    defined name cannot refer to a JSON string), the architecture contract
    already names the tables, and tables stay inspectable and repairable
    without the add-in (R2.10). The recorded trade-off: tables cost a
    larger COM surface and a wider corruption-detection surface than one
    JSON cell, paid for with the catalogue hash, typed reader refusals,
    and the D4 regeneration policy; JSON additionally caps at 32,767
    characters per cell.
  - **D2 — Table set (5 tables + the later defined name).** `tblGanttTypes`
    (16 rows projected from `EntityTypeCatalog.Entries`),
    `tblGanttStyles` (built-in presets, one per non-empty
    `DefaultStyleKey` — 15 of 16; user customs appended only by R5.10),
    `tblGanttMetrics` (23 rows), `tblGanttSettings` (11 rows), and
    `tblGanttConfig` (schema version, catalogue hash, workbook ID,
    add-in version last used). `GanttCreator.TypeOptions` is defined at
    R2.9 against `tblGanttTypes`. The architecture contract's separate
    `tblGanttLabelPositions` table is deferred: v1 carries the valid label
    choices per entity in the `tblGanttTypes.AllowedLabelPositions`
    column, which is the same grouping the entity guide requires.
  - **D3 — Settings storage: `tblGanttSettings`.** A key/value table
    (`SettingKey`, `SettingValue`; text values), keeping `tblGanttConfig`
    purely add-in metadata. The architecture "VeryHidden configuration
    worksheet contract" table is amended in the same change. Approved
    first-release keys: `ChartTitle` (blank allowed), `ShowTitle` (`TRUE`),
    `TimeScale` (`Month`), `LegendPosition` (`Right`), `ExportIncludeDataPanel`
    (`TRUE`), `ExportIncludeLegend` (`TRUE`), `PlotStartMode` (`DataRange`),
    `PlotFinishMode` (`DataRange`), `AlternateBanding` (`TRUE`),
    `ShowMinorGrid` (`TRUE`), `ShowMajorGrid` (`TRUE`).
    `PlotStartMode`/`PlotFinishMode`/`TimeScale` value enums are fixed by
    R3.3/R5.1; until then the stored text is the default only.
  - **D4 — Regeneration policy.** The writer never destroys user content:
    `tblGanttTypes` and `tblGanttMetrics` regenerate only from the
    code-owned catalogues; `tblGanttStyles` regenerates built-in keys and
    preserves non-built-in rows; `tblGanttSettings` preserves present
    values and adds missing keys. Settings are written by approved dialogs
    only; R2.7 writes defaults when a key is absent and never overwrites a
    present value. The settings table always regenerates to exactly the
    approved key set: an unapproved key is not carried forward, because
    the reader validates the exact key set (D3) and preserving one would
    leave the workbook unreadable after every regeneration. Dialog-authored
    keys therefore require a schema-version and catalogue-hash decision
    (D5/D6) owned by the settings-dialog work item, not a silent row.
  - **D5 — Schema version stays 1.** v1 never shipped config tables;
    materialising them completes v1's definition. Future catalogue changes
    bump `GanttSchemaVersion.CurrentSchemaVersion` per its remarks.
  - **D6 — Catalogue hash.** `tblGanttConfig` stores a SHA-256 (lowercase
    hex, 64 chars) over a canonical, culture-invariant serialisation of
    the code-owned catalogue definitions (types, styles, metrics, colours,
    typography, settings, schema version), computed by Core
    (`GanttCatalogues.ComputeCatalogueHash()`); never over raw worksheet
    text. Doubles serialise with the round-trip `"R"` format in invariant
    culture; rows join with `|`; tables emit as
    `table|col,col|value|value` lines in contract order.
  - **D7 — Table geometry and `tblGanttStyles` column layout.** Tables are
    anchored at fixed cells on `_GanttCreatorConfig` (deterministic,
    gap-separated, documented in `GanttCatalogues`):
    `tblGanttTypes` at `A1`, `tblGanttStyles` at `A20`, `tblGanttMetrics`
    at `A40`, `tblGanttSettings` at `A70`, `tblGanttConfig` at `A90`.
    `tblGanttStyles` columns: `StyleKey`, `DisplayName`, `FillColour`,
    `StrokeColour`, `HatchPattern` (`None`, `ForwardDiagonal`,
    `BackwardDiagonal`, `Cross`), `HatchPitchPt`, `HatchLinePt`,
    `TextColour`, `StandardOutlinePt`, `ActivityHeightPt`,
    `MilestoneSizePt`. **ADR-0009 (2026-09-23) appends `DefaultLabelPosition`, `AllowedLabelPositions`, and `ColourCapability` so Custom Activity and R3 style resolution have explicit named-style capabilities.** `tblGanttTypes` columns: `TypeName`, `DisplayName`,
    `Kind`, `DateMode`, `DefaultStyleKey`, `ColourCapability`,
    `RequiresStyleKey`, `AllowedLabelPositions` (space-joined).
    `tblGanttMetrics` columns: `MetricName`, `DefaultValue`, `Minimum`,
    `Maximum`. `tblGanttConfig` columns: `ConfigKey`, `ConfigValue`.
  - **D8 — Visible-table appearance (human-approved plan amendment,
    2026-09-21).** Initialise also sets the visible `tblGanttData`
    appearance: `ListObject.ShowAutoFilter = false` and
    `ListObject.ShowTableStyleRowStripes`/`ShowTableStyleColumnStripes`
    `= false`. These are direct boolean properties of `ListObject` in the
    installed PIA (verified by reflection against
    `ExcelDna.Interop` 16.0.0 — no separate `TableStyleInfo` member
    exists there). Banding is neutralised by flipping flags, never by
    assigning a built-in style by name (built-in style names are
    localised; an English name can fail on non-English installs). The
    Table Design contextual ribbon tab itself is Excel shell behaviour
    and is out of scope.
- **Consequences**: Phase 3 (R3.3/R3.9) reads stable table names.
  R2.10's repair contract regenerates through the same writer, so the D4
  guarantees are proven once. The wider per-cell corruption surface is
  detected by typed reader refusals and the stored hash. The catalogue
  enumeration remains owned by the entity guide; any token change is a
  guide change plus a schema-version decision.
- **Alternatives considered**: one JSON document in a single cell
  (rejected — no cell range for the R2.9 defined name, 32,767-character
  cap, not inspectable without the add-in); settings rows inside
  `tblGanttConfig` (rejected — mixes user-mutable values with add-in
  metadata and complicates the R2.10 repair contract).
