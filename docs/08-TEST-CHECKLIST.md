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
      exponents (`"1e999"`) throw the documented exception, not silent
      propagation.
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
- [ ] Inclusive-finish policy: a one-day span has exactly one day of width;
      right edge uses `DateToX(Finish + 1 day)`. Milestones and delineators do
      **not** add the extra day.
- [ ] `StackIndex` equal values share the same vertical centre; sparse values
      preserve order but do not create empty height.
- [ ] Lane height grows to fit content (`max(LaneHeightPt, contentHeight)`);
      events are never compressed. Test with stacked events exceeding minimum.
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
      validated against the selected Type''s capabilities. A `Critical Interval`
      with a `FillColour` override is a blocking error.
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
