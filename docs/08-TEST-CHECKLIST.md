# Pre-commit test checklist

> Always-on gate. For every commit that touches production code, verify the
> applicable items below. Items are grouped by the bug class they prevent.
> Domain-specific items reference `docs/02-ARCHITECTURE.md` and
> `docs/07-GANTT-ENTITY-GUIDE.md`.

---

## A. Numeric code — culture, non-finite, overflow

Applies to: any public method that parses, formats, or converts numeric values.

- [ ] Culture-roundtrip test: code sets `CurrentCulture` to a comma-decimal
      culture (e.g. `de-DE`) and asserts the exact output / parsed value.
- [ ] Non-finite input tests: `NaN`, `+Infinity`, `-Infinity`, overflowed
      exponents (`"1e999"`). Each consumer's documented rejection is asserted per its contract: `ExportSize.ParseValue` raises `FormatException`; `ExportSize.ToPixels` raises `ArgumentOutOfRangeException`; numeric casts guarded by `IsFinite` raise the documented guard exception. No silent
      propagation of non-finite or out-of-range values.
- [ ] Every `double`/`float`→`int` cast is guarded by `IsFinite` + range
      check; tests cover each violation.
- [ ] Rounding happens at exactly one documented boundary; no repeated
      rounding through the pipeline (see `docs/02-ARCHITECTURE.md`
      "Coordinate and rounding policy").

---

## B. Scene and geometry invariants

Applies to: `GanttCreator.Core` scene construction, layout, clipping.

- [ ] No scene primitive has `NaN`, `Infinity`, negative width, or negative
      height. Assert by walking the built scene.
- [ ] `TimeScale.DateToX` is monotonically increasing across the plot range;
      test with shuffled, reversed, and duplicate dates.
- [ ] Inclusive duration policy: `DurationDays = Finish.DayNumber - Start.DayNumber + 1`; activity width is `DurationDays * DayWidth`, with the left edge fixed at `DateToX(Start)`. `Finish` is never incremented. Milestones and delineators use only their point date X.
- [ ] Effective stack indices are derived in Core from deterministic row/parent-child position; the visible `StackIndex` cell is neither trusted nor required for layout. Generated indices are consecutive; sparse compatibility values preserve order without empty height, and duplicate effective values share one centre.
- [ ] **Lane height equals the measured worksheet row height (revision 8, ADR-0026).**
      `LaneHeightPt` is the lane height, **not a minimum the layout may exceed**:
      lane auto-growth is removed, not capped. A lane that grows disagrees with
      the Excel row it sits in, and a projected child must not be able to grow the
      lane its parent owns. When lane content does not fit the fixed height the
      scene emits `LaneContentExceedsRowHeight` and **neither compresses the
      content nor grows the lane**. Test with stacked events exceeding the height,
      with a projected child, and across expand/collapse.
- [ ] Clipping never expands bounds; clipped events produce the documented
      warning, not a silent no-op (unless the entity contract says otherwise).
- [ ] Deterministic ordering: shuffled input produces the same scene
      serialization as ordered input. Test with the same fixture shuffled >=3
      ways.

---

## C. Entity contract compliance

Applies to: any code that creates, validates, or renders entity instances.

- [ ] `Type` values come only from the central `EntityTypeCatalog`; unknown
      pasted values are blocking errors. No renderer maintains its own type
      list.
- [ ] Per-row overrides (`FillColour`, `StrokeColour`, `LabelPosition`) are
      validated against the selected Type's capabilities. **A `Critical Interval`
      with a `FillColour` override is now VALID (revision 8, ADR-0027)** — the
      critical overlay is a filled rectangle — and the validator, the
      `EntityTypeCatalog` capability and this item must all agree. Test both
      directions: a permitted fill is accepted and reaches the host, and a
      capability the type does not hold is still a blocking error.
- [ ] Milestones and delineators read `Start` only; `Finish`, `LaneId`,
      `StackIndex` are not read for geometry. Populate them and verify a
      non-blocking "not used" warning, not silent data loss.
- [ ] `Custom Activity` requires a valid `StyleKey`; unknown key is a
      validation error. The renderer must not fall back silently.
- [ ] Critical intervals clip to both the plot **and** the visible parent
      span. Out-of-parent portions warn and clip per the approved policy.

---

## D. VeryHidden configuration worksheet

Applies to: any code that reads or writes `_GanttCreatorConfig`.

- [ ] Exactly one visible worksheet plus one `xlSheetVeryHidden` config sheet.
      No other sheets exist.
- [ ] Config sheet contains **no** activity/event rows, descriptions, schedule
      dates, scene primitives, rendered shapes, formulas that determine chart
      geometry, logs, or export staging.
- [ ] Schema version, catalogue hash, table headers, and defined names are
      validated on initialise/refresh; corrupt or missing config is detected.
- [ ] User style presets are preserved during migration; the repair path
      never silently drops custom configuration.

---

## E. Renderers consume, never recalculate

Applies to: `GanttCreator.Office`, `GanttCreator.Raster`, and any future
renderer.

- [ ] Renderers consume resolved scene primitives and do not independently
      move labels, recalculate dates, substitute colours, change line widths,
      or reorder entities.
- [ ] Z-order matches the entity guide''s layer table exactly; test that
      entities appear in the documented back-to-front order.
- [ ] Live renderer owns only shapes with the `GanttCreator` tag/prefix;
      unowned shapes/cells are never deleted or reformatted.
- [ ] Refresh is idempotent: two refreshes produce identical owned IDs,
      counts, and bounds.
- [ ] Blocking errors leave the last valid chart unchanged; partial temporary
      export shapes are removed.
- [ ] **No renderer reinterprets a scene primitive (revision 8, ADR-0027).** The
      critical interval is a filled rectangle in the scene and must be drawn as a
      filled rectangle by every renderer, with no critical-specific branch. This
      is the direct replacement for revision 6's "the scene rect is not the host
      object" rule, which no renderer implemented. Assert per renderer, and assert
      the rectangle is **filled** — an outline-only or line-only result is a
      defect, not a variant.

---

## F. Date handling

Applies to: any code that reads or writes Excel dates.

- [ ] Date parsing uses the Excel cell value and workbook date system, not
      locale-dependent display text.
- [ ] Workbook 1900 date system is required; 1904 is either supported and
      tested or rejected explicitly (document the choice in an ADR).
- [ ] Leap day, year boundary, and one-day activities are tested.
- [ ] Date values around Excel''s invalid/edge serial values are tested.

---

## G. Build pipeline and artifacts

Applies to: `scripts/verify-quick.ps1`, `scripts/verify.ps1`, CI config.

- [ ] Every test that reads from `bin/` or `publish/` is traceable to a
      verify-script step that produces that artifact. Document the guarantee
      in the work item.
- [ ] `.vscode/tasks.json` (or any JSON config) is validated by an
      architecture test: correct schema, `dependsOn` labels resolve, dependency
      ordering is correct.

---

## H. Suppression scope

Applies to: `Directory.Build.props`, `tests/Directory.Build.props`, `NoWarn`.

- [ ] Test-only suppressions (CA1707, IDE0011) live in
      `tests/Directory.Build.props`, never in the root. An architecture test
      asserts this.
- [ ] Every `NoWarn` addition has a written reason in a comment and an ADR if
      it materially affects production code.

---

## I. Work-item traceability

Applies to: every `docs/work-items/R*.md` file.

- [ ] Acceptance criteria that name a test count, a command, or an artifact
      link to a specific test or script step. Drift between the doc and the
      code is a defect.
- [ ] The stated test count matches the actual `[Fact]`/`[Theory]` count in
      the committed code.

---

## J. COM ownership (when COM interop is touched)

Applies to: any code path that calls Excel or PowerPoint via COM.

- [ ] Proxies are captured to local variables, released in reverse order.
- [ ] No chained COM expressions (`app.ActiveWorkbook.Worksheets[1]...`).
- [ ] Application state is saved and restored even after failure
      (`try/finally`).

---

## K. Failure injection

Applies to: any new public method with external dependencies.

- [ ] Controllable failure points around: worksheet read, each
      application-state change, shape create/style/group/copy/delete,
      clipboard acquisition, file dialog, temporary file, encode, metadata.
- [ ] For every injected failure assert: cleanup, preserved user content,
      restored application state, one user-facing error, one technical record.

---

## L. Identity and hierarchy

Applies to: R4.7A, and every read or mutation that touches `Id`/`ParentId`.

- [ ] `ParentId` is authoritative and `SiblingOrder` is engine-maintained. Row
      number is never persistent identity; moving a row must not rewrite an ID.
- [ ] Identity is assigned when a blank entry row first becomes a meaningful
      entity, and a completely blank row stays a non-entity with no ID. One
      `IsSemanticallyEmptyRow()` rule is reused, never re-implemented per call site.
- [ ] Exactly one blank entry row remains after every mutation.
- [ ] The seven-child maximum is enforced: an eighth child is refused with a
      typed error, and the refusal has a positive test.
- [ ] Every broken, ambiguous, cyclic, or self-referential `ParentId` is
      reported as a blocking validation issue before any render. Multi-level
      nesting is either explicitly supported or explicitly rejected — never
      silently accepted.
- [ ] A pasted row receives a **new** ID; a cloned parent/child pair retains its
      relationship under **new** IDs; a duplicated managed structure is reseeded
      **atomically**, never one duplicated row at a time.
- [ ] Deleting a parent **promotes** rather than deletes its children: they retain
      their IDs and schedule data, and `ParentId`, `SiblingOrder`, lane ownership
      and outline grouping are recomputed.
- [ ] Engine columns are hidden and locked, and that state is repaired if a user
      unhides them.

---

## M. Projection and source visibility

Applies to: R4.7B, and any code deciding which lane an entity draws on.

- [ ] Projection is resolved in **Core** before scene construction. Excel must not
      decide which lane an entity belongs to.
- [ ] `SourceRowVisible` and `RenderVisible` are separate concepts. A collapsed
      child is `SourceRowVisible = false` **and** `RenderVisible = true`.
- [ ] An Excel row hidden by hierarchy collapse is **never** mapped to
      `GanttEvent.Visible = false`.
- [ ] Expand/collapse may change source-row visibility and later lane Y
      positions, but must not change whether a child renders, its dates, style,
      identity, or parent relationship.
- [ ] A projected child creates no independent lane and no lane growth; multiple
      children may overlap one parent lane, distinguished by date geometry, Type,
      style, z-order and label.
- [ ] Only Types permitted by the Core capability matrix may be a child, and the
      matrix has a test for every member.

---

## N. Layout and size presets

Applies to: R4.7D, R4.7H, and any code deriving plot geometry.

- [ ] Exactly one authority derives `PlotBounds` from the size preset, the
      measured text-panel width, the date range and the scale. No renderer or
      command derives it independently.
- [ ] `PlotWidth = PresetWidth − TextPanelWidth − Margins/Borders`, and the title
      and time-header bands use the same resulting composition geometry.
- [ ] Text and data columns **retain their measured widths** under every preset.
      A preset never resizes user text columns to make itself fit.
- [ ] Insufficient remaining width for a valid plot **refuses** Refresh with an
      actionable error. It never silently shrinks the text panel or the plot.
- [ ] A4, Presentation 16:9 and Presentation 4:3 each have a geometry snapshot,
      and the aspect ratio is exact rather than approximated.
