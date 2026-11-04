using System.Diagnostics;
using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// The pipeline-shaped <see cref="IGanttRefreshOrchestrator"/>: it composes the
/// already-landed R4.7 and R4.8 contracts in the fixed order R4.8A D2 requires.
/// </summary>
/// <remarks>
/// <para>
/// <b>The order is the product, not an implementation detail.</b> Validation must
/// precede mutation, the plot must be resolved before the scene is built, and the
/// scene must be complete before a single shape moves. Each of those is a rule a user
/// would notice breaking — a half-rendered chart, a chart sized against stale rows —
/// so the sequence lives in one named type rather than in a Ribbon callback whose
/// ordering is a property of how it was written.
/// </para>
/// <para>
/// <b>What "before any shape mutation" means concretely.</b> Steps 1-13 below issue
/// no shape write at all: they read, validate, plan, and build. The
/// <see cref="IShapeWritePort"/> is first touched by
/// <see cref="ShapeReconciler.Reconcile"/> in the final step, and its
/// <c>ListOwned</c> read is itself skipped when the refresh is blocked. A test can
/// therefore inject a failure at any step and assert the shape port recorded zero
/// calls — which is D3, and it is only assertable because every collaborator is a
/// port.
/// </para>
/// <para>
/// <b>The three writes that are not shape mutations.</b> <c>Duration</c>, the <c>StyleKey</c> column, and the outline
/// are worksheet writes that happen before the reconciliation, because all three are
/// derived from validated rows and all three must be current before the geometry is
/// measured. They are reported separately from the shape count so a caller can tell
/// "the chart was not touched" from "the sheet was not touched at all".
/// </para>
/// <para>
/// <b>Partial refresh is reported, not rolled back.</b> REV5 §16 defers transactional
/// Refresh to R4.10, so a mid-reconciliation refusal is surfaced as
/// <see cref="GanttRefreshRefusal.PartialReconciliation"/> with the count of
/// operations that did land. Pretending a partial write did not happen would be worse
/// than the tear itself.
/// </para>
/// </remarks>
public sealed class GanttRefreshOrchestrator(
    IGanttTableReader tableReader,
    IConfigCatalogueReader configReader,
    IWorksheetProtectionGuard protectionGuard,
    IPanelGridMeasurementPort panelMeasurement,
    IDurationWritePort durationWriter,
    IRowHeightNormalisationPort rowHeightNormaliser,
    IOutlineGroupPort outlineWriter,
    ISceneBuildRequestFactory requestFactory,
    IShapeWritePort shapeWriter,
    IColumnPresentationPort? columnPresentation = null,
    IGanttRowIdentityRepairer? identityRepairer = null,
    IApplicationStateScope? stateScope = null,
    IStyleKeyWritePort? styleKeyWriter = null,
    IRowWarningHighlighter? rowWarningHighlighter = null,
    ITypeChangeNormaliser? typeChangeNormaliser = null,
    ChangeTracker? changeTracker = null,
    ShapeDeltaCache? shapeDeltaCache = null,
    bool useDeltaCache = true,
    Action<string>? trace = null)
    : IGanttRefreshOrchestrator
{
    private readonly IGanttTableReader _tableReader = tableReader ?? throw new ArgumentNullException(nameof(tableReader));
    private readonly IConfigCatalogueReader _configReader = configReader ?? throw new ArgumentNullException(nameof(configReader));
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? throw new ArgumentNullException(nameof(protectionGuard));
    private readonly IPanelGridMeasurementPort _panelMeasurement = panelMeasurement ?? throw new ArgumentNullException(nameof(panelMeasurement));
    private readonly IDurationWritePort _durationWriter = durationWriter ?? throw new ArgumentNullException(nameof(durationWriter));
    private readonly IRowHeightNormalisationPort _rowHeightNormaliser = rowHeightNormaliser ?? throw new ArgumentNullException(nameof(rowHeightNormaliser));
    private readonly IOutlineGroupPort _outlineWriter = outlineWriter ?? throw new ArgumentNullException(nameof(outlineWriter));
    private readonly ISceneBuildRequestFactory _requestFactory = requestFactory ?? throw new ArgumentNullException(nameof(requestFactory));
    private readonly IShapeWritePort _shapeWriter = shapeWriter ?? throw new ArgumentNullException(nameof(shapeWriter));

    /// <summary>
    /// Writes the <c>StyleKey</c> column so column K matches each row's current
    /// <c>Type</c>, or <see langword="null"/> to skip the step.
    /// </summary>
    /// <remarks>
    /// Optional rather than required so a caller with no Excel application, and a
    /// test that is not about style-key state, need not construct one. When it is
    /// absent the pipeline behaves exactly as it did before this row.
    /// </remarks>
    private readonly IStyleKeyWritePort? _styleKeyWriter = styleKeyWriter;

    /// <summary>
    /// The row/cell fill highlighter, or <see langword="null"/> to skip row
    /// highlighting. Replaces <see cref="IGanttValidationReporter"/>: fills are
    /// cosmetic cell fills rather than cell notes, and no <c>SpecialCells</c>
    /// scan is needed to clear stale reports.
    /// </summary>
    private readonly IRowWarningHighlighter? _rowWarningHighlighter = rowWarningHighlighter;

    /// <summary>Tracks pending automatic changes until a complete render succeeds.</summary>
    private readonly ChangeTracker? _changeTracker = changeTracker;
    private readonly ShapeDeltaCache? _shapeDeltaCache = shapeDeltaCache;
    private readonly bool _useDeltaCache = useDeltaCache;

    /// <summary>
    /// Optional diagnostic sink for one refresh's delta and geometry decisions, or
    /// <see langword="null"/> for no trace.
    /// </summary>
    /// <remarks>
    /// <b>Why a trace and not a behaviour change.</b> The scheduled refresh skips any
    /// shape whose payload signature is unchanged, and the manual Refresh bypasses the
    /// snapshot. A stale shape therefore has two possible causes — a scene that really
    /// changed, or a live shape Excel moved or resized on its own — and no code path
    /// currently says which. This sink names the decision per splitter band and the
    /// structural writes in the same run, so the live host can be diagnosed instead of
    /// guessed at. The production sink is the rolling log; it degrades internally and
    /// must not throw, exactly like the application-state scope's sink.
    /// </remarks>
    private readonly Action<string>? _trace = trace;

    /// <summary>The scene role suffix that identifies a splitter band primitive.</summary>
    private const string _splitterBandRoleSuffix = ":splitter-band";

    /// <remarks>
    /// <para>
    /// Optional rather than required so a caller with no Excel application, and a test
    /// that is not about column state, need not construct one. When supplied the step
    /// runs; when it is absent the pipeline behaves exactly as it did before this row,
    /// which is what keeps this change additive rather than a behaviour change for
    /// existing callers.
    /// </para>
    /// <para>
    /// <b>Why the step exists at all.</b> R4.7C applies the classification on
    /// Initialise only, and its own code comment records that "un-hide one and refresh
    /// restores it" is a repair belonging to R4.8A. Without this step a user who
    /// unhid an engine column kept it hidden in name only: the column stayed visible
    /// and the authoring surface no longer matched the schema the validator enforces.
    /// </para>
    /// </remarks>
    private readonly IColumnPresentationPort? _columnPresentation = columnPresentation;

    /// <summary>
    /// Seeds or repairs row identifiers, or <see langword="null"/> to skip the step.
    /// </summary>
    /// <remarks>
    /// R4.8A D2 names "process blank and new entities and IDs" as a preflight step.
    /// Identity repair repairs only the <c>Id</c> column — malformed, missing, or
    /// later duplicate — and so is the narrow operation that belongs here. The wider
    /// promotion-on-delete and clone-reseed behaviour belongs to the hierarchy row,
    /// not to a Refresh.
    /// </remarks>
    private readonly IGanttRowIdentityRepairer? _identityRepairer = identityRepairer;

    /// <summary>
    /// Normalises type-dependent engine-hidden columns on Type change, or
    /// <see langword="null"/> to skip the step.
    /// </summary>
    private readonly ITypeChangeNormaliser? _typeChangeNormaliser = typeChangeNormaliser;

    /// <summary>
    /// The application-state scope, or <see langword="null"/> for the no-op scope.
    /// </summary>
    /// <remarks>
    /// Optional rather than required, so a caller with no Excel application — and a
    /// test that is not about restoration — does not have to construct one.
    /// </remarks>
    private readonly IApplicationStateScope? _stateScope = stateScope;

    /// <summary>The scene-to-shape renderer, held as a field so a test can observe
    /// that the orchestrator owns the translation step rather than a caller.</summary>
    private readonly SceneShapeRenderer _renderer = new(ChartOriginDelta.Identity);

    /// <inheritdoc />
    public GanttRefreshOutcome Refresh()
    {
        // D8: the application settings and the selection are captured and
        // restored on EVERY path, success or failure, by the ADR-0020 scope. A
        // refresh turns off screen updating and events for speed and then adds
        // hundreds of shapes; if it refused half way, the user would be left with
        // events disabled and a status bar still carrying our text, with no way to
        // tell that anything is wrong. The `using` is what makes that guarantee
        // total: every return below runs Dispose, so there is no path that can
        // skip the restore.
        //
        // `DisplayStatusBar` is deliberately NOT suppressed (owner decision,
        // 2026-10-01). It used to be hidden here and the progress text written
        // immediately afterwards, which meant the message was set and restored
        // correctly but never seen -- the two are different Excel properties, and
        // hiding the bar hides the text inside it. Leaving visibility alone also
        // means the user's own setting decides: a user who keeps the status bar
        // hidden is not shown one for the duration of a refresh, which is their
        // preference and not ours to override.
        using IApplicationStateScope scope = _stateScope ?? NullApplicationStateScope.Instance;

        scope.SuppressScreenUpdating();
        scope.SuppressEvents();
        scope.SuppressAlerts();
        scope.SetStatusBarText("Rendering the Gantt chart…");
        scope.CaptureSelection();

        return RunRefresh();
    }

    /// <summary>
    /// The pipeline itself, with the application state already suppressed.
    /// </summary>
    /// <returns>The typed outcome.</returns>
    private GanttRefreshOutcome RunRefresh()
    {
        long refreshStarted = Stopwatch.GetTimestamp();

        // Step 1-2. Read the complete user's table and configuration. Automatic
        // refresh may use its payload-signature snapshot to skip unchanged host
        // writes, but it never validates only changed rows: identifiers,
        // hierarchy, and layout constraints can depend on any row. Both reads are
        // read-only, so a refusal here has mutated nothing at all.
        GanttTableReadOutcome rows = _tableReader.Read();
        if (!rows.Succeeded)
        {
            // No active workbook is its own refusal, not a table problem: telling a
            // user with no open workbook that "the Gantt table was not found on the
            // active sheet" describes a workbook they do not have. Every other read
            // refusal does mean the table could not be read, and keeps that mapping.
            GanttTableReadRefusalReason readRefusal =
                rows.Refusal ?? GanttTableReadRefusalReason.TableMissing;
            return readRefusal == GanttTableReadRefusalReason.NoActiveWorkbook
                ? Refuse(GanttRefreshRefusal.NoActiveWorkbook, Describe(readRefusal))
                : Refuse(GanttRefreshRefusal.TableMissing, Describe(readRefusal));
        }

        ConfigReadOutcome config = _configReader.Read();
        if (!config.Succeeded)
        {
            return Refuse(
                GanttRefreshRefusal.ConfigurationUnreadable,
                "The Gantt Creator configuration could not be read from the hidden configuration sheet.");
        }

        // Step 3. The mutation boundary, consulted before any write. A protected
        // worksheet would otherwise be discovered halfway through, leaving a
        // half-written sheet behind.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return Refuse(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttRefreshRefusal.NoActiveWorkbook
                    : GanttRefreshRefusal.TargetProtected,
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? "There is no active workbook to refresh."
                    : "The Gantt worksheet is protected, so the chart cannot be updated. Unprotect it and try again.");
        }

        // Step 4. The managed columns' hidden/locked classification, restored before
        // anything is validated or measured. A user who unhid an engine column would
        // otherwise keep seeing it, and the panel measurement that follows would
        // measure a column the schema says is not there. R4.7C applies this on
        // Initialise and records that this repair belongs to R4.8A.
        if (_columnPresentation is { } columns)
        {
            ColumnPresentationOutcome presentation = columns.EnsureClassification();
            if (!presentation.Succeeded)
            {
                return Refuse(
                    GanttRefreshRefusal.ColumnPresentationRefused,
                    "The managed columns could not be restored to their schema state.");
            }
        }

        // Step 5. Row identity: a row with no usable Id, or a duplicate one, cannot be
        // validated or projected, so it is seeded here rather than surfacing as a
        // blocking error the user has no way to clear.
        if (_identityRepairer is { } identity)
        {
            GanttRowIdentityRepairOutcome repaired = identity.Repair();
            if (repaired.Refusal is not null)
            {
                return Refuse(
                    GanttRefreshRefusal.IdentityRefused,
                    "The row identifiers could not be repaired.");
            }

            // The step-1 snapshot is STALE only when repair actually wrote. Repairing
            // nothing is the normal case, and re-reading the whole table for it is a
            // second full COM read on every Refresh for no new information — the id
            // repairer reports its own write count precisely so this can be skipped.
            if (repaired.RepairedCount > 0)
            {
                // Repair writes new Id values into the worksheet, and those writes are
                // the whole point of the step. Validating the pre-repair rows would
                // report the blank or malformed Ids that repair just fixed — so a
                // workbook needing only an ID refresh refused with
                // BlockingValidationErrors, which is both wrong and unfixable by the
                // user. Re-read through the same port, and take the same refusal path a
                // failed first read takes, because from here on it is indistinguishable
                // from one.
                rows = _tableReader.Read();
                if (!rows.Succeeded)
                {
                    GanttTableReadRefusalReason rereadRefusal =
                        rows.Refusal ?? GanttTableReadRefusalReason.TableMissing;
                    return rereadRefusal == GanttTableReadRefusalReason.NoActiveWorkbook
                        ? Refuse(GanttRefreshRefusal.NoActiveWorkbook, Describe(rereadRefusal))
                        : Refuse(GanttRefreshRefusal.TableMissing, Describe(rereadRefusal));
                }
            }
        }

        // Step 5b. Normalise type-dependent engine-hidden columns so a Type
        // change does not leave stale values the validator would reject. Writes
        // land on the sheet, so rows are re-read afterwards — exactly as identity
        // repair does — so validation sees the corrected values.
        if (_typeChangeNormaliser is { } normaliser)
        {
            GanttTypeChangeNormaliseOutcome normalised = normaliser.Normalise();

            // Same rule as identity repair: re-read only when cells were written. The
            // normaliser walks every row to decide, so a no-op is the common case and
            // its count is the evidence that the snapshot is still current.
            if (normalised.CellsWritten > 0)
            {
                // Re-read so validation sees the corrected engine-hidden cells.
                rows = _tableReader.Read();
                if (!rows.Succeeded)
                {
                    GanttTableReadRefusalReason rereadRefusal =
                        rows.Refusal ?? GanttTableReadRefusalReason.TableMissing;
                    return rereadRefusal == GanttTableReadRefusalReason.NoActiveWorkbook
                        ? Refuse(GanttRefreshRefusal.NoActiveWorkbook, Describe(rereadRefusal))
                        : Refuse(GanttRefreshRefusal.TableMissing, Describe(rereadRefusal));
                }
            }
        }

        long readCompleted = Stopwatch.GetTimestamp();

        // Step 6. Validate every row before derived worksheet writes or shape
        // mutation. This is required on both explicit and automatic refreshes.
        //
        // The style registry is passed so a Custom Activity / Custom Milestone
        // row whose user-created style has been deleted is reported by name
        // (a warning) rather than passing through unexamined. The registry
        // comes from the same read at step 2, so no second sheet read is
        // introduced.
        GanttValidationOutcome validation = GanttRowValidator.Validate(rows.Rows, config.Styles);
        long validationCompleted = Stopwatch.GetTimestamp();

        // Step 6b. Bring the cell notes into line with what was just found, BEFORE
        // deciding whether to render.
        //
        // This runs on BOTH paths, which is the whole point. A refresh that now finds
        // no errors must REMOVE the notes a previous Validate left behind, and a
        // refresh that finds errors must write the current ones even though it will
        // not render. Calling it only on the success path would leave the stale-note
        // defect in place for the user who had just fixed their data.
        //
        // A refusal here is surfaced rather than swallowed: the notes are the visible
        // evidence of why the refresh did not complete, and a silent failure would
        // leave the sheet showing the previous report while the chart stayed put.
        // Nothing has been mutated on the sheet at this point, so this refusal reports
        // zero worksheet writes.
        GanttValidationReportOutcome reported = ReportValidation(validation.Issues);
        if (!reported.Succeeded)
        {
            return Refuse(
                GanttRefreshRefusal.ValidationNotesRefused,
                "The validation notes on the table could not be updated, so the refresh stopped before the chart was touched.");
        }

        // Partial render: rows with blocking errors are excluded from the scene
        // (they render blank) but every other row plots normally. Fills are still
        // painted for the invalid rows so the user sees the current error set.
        // validation.Events already contains only rows without blocking errors.
        //
        // When every row is blocked there is nothing to render: refuse before any
        // downstream worksheet write so the existing chart and sheet are untouched.
        if (validation.Events.Count == 0 && validation.Issues.Any(i => i.Severity == GanttValidationSeverity.Error))
        {
            return GanttRefreshOutcome.Refused(
                GanttRefreshRefusal.BlockingValidationErrors,
                "The table has errors, so the chart was not rebuilt. The existing chart is unchanged.",
                validation.Issues);
        }

        // The events carry each row's StyleKey straight from column K as read at
        // step 1. If the user changed a row's Type since the last Refresh, that cell
        // still holds the previous row's style key until we rewrite it here. The
        // write must land before the scene is built, and the events the scene is
        // built from must reflect what we wrote — otherwise a type change takes two
        // Refreshes to show the right style.
        IReadOnlyList<GanttEvent> events = validation.Events;

        // Step 6. Worksheet presentation. The outline is rebuilt from the validated
        // hierarchy, never read back, so collapsing a group stays presentation.
        OutlineGroupOutcome outline = _outlineWriter.Apply(events);
        if (!outline.Succeeded)
        {
            return Refuse(GanttRefreshRefusal.OutlineRefused, "The row outline groups could not be updated.");
        }

        // Step 7. Duration. Calculated in Core and written in one bulk pass; an
        // already-correct column produces no write at all.
        DurationWritePlan plan = DurationCalculator.Plan(
            events,
            ReadCurrentDurations(rows.Rows));
        DurationWriteOutcome duration = _durationWriter.Write(plan);
        if (!duration.Succeeded)
        {
            return Refuse(GanttRefreshRefusal.DurationWriteRefused, "The Duration column could not be written.");
        }

        // Step 7b. StyleKey: column K matches each row's current Type. A row whose
        // cell already carries its type's default is absent from the plan, so a
        // correctly written sheet is not touched; a stale cell (for example K4
        // stamped As-Planned while the row is now As-Built) is overwritten.
        //
        // The style registry is passed so a style-key row whose user-created
        // style was deleted is HEALED: its cell is rewritten to the type's
        // code-owned default key, which always resolves. Without the registry
        // the planner cannot tell a deleted style from a deliberate choice, so
        // the dangling reference would survive every Refresh and the row
        // would keep rendering from a style that no longer exists.
        StyleKeyWritePlan stylePlan = StyleKeyWritePlanner.Plan(events, config.Styles);
        int styleCellsWritten = 0;
        if (_styleKeyWriter is { } styleWriter)
        {
            StyleKeyWriteOutcome style = styleWriter.Write(stylePlan);
            if (!style.Succeeded)
            {
                return RefuseAfterWrites(
                    GanttRefreshRefusal.StyleKeyWriteRefused,
                    "The StyleKey column could not be updated.",
                    validation.Issues,
                    duration.CellsWritten,
                    styleKeyCellsWritten: style.CellsWritten);
            }

            // The sheet now holds the new keys. Mirror them into the events so the
            // scene renders the corrected style instead of the stale key read at
            // step 1. Without this a type change needs a second Refresh.
            events = ApplyStyleKeyToEvents(events, stylePlan.Writes);
            styleCellsWritten = style.CellsWritten;
        }

        long worksheetWritesCompleted = Stopwatch.GetTimestamp();

        // Step 8. Normalize row heights before measuring and building the scene.
        int rowHeightsWritten;

        // The heights the normaliser measured, when it measured without writing any of
        // them. Step 9 needs exactly these figures, so a pass that changed none hands
        // them over rather than making the measurement read every body row again.
        IReadOnlyList<double> measuredBodyHeights;
        {
            RowHeightNormalisationOutcome heights = _rowHeightNormaliser.Normalise(
                ThemedMetric(config.Metrics, "GanttRowHeightPt"),
                ThemedMetric(config.Metrics, "SplitterHeightPt"),
                ThemedMetric(config.Metrics, "SpacerHeightPt"),

                // ADR-0030 D4/D5: the header row IS the period band's row, so its
                // height is what makes the period band's bottom coincide with the first
                // body row's top; the reserved row above it carries the table title and
                // the year band. Both come from the workbook's metrics for the same
                // reason the three above do -- a literal here is a second authority
                // that can drift from the value the user retunes.
                ThemedMetric(config.Metrics, "PeriodBandHeightPt"),
                ThemedMetric(config.Metrics, "YearBandHeightPt"),

                // ADR-0031 D2: the chart's top and bottom margins are worksheet rows, so
                // they are normalised to their own token rather than borrowing a band's
                // height. It is the LAST argument because it is the newest row, and the
                // workbook's metric -- not a literal -- is the authority for how tall
                // it is.
                ThemedMetric(config.Metrics, "ChartPaddingRowHeightPt"),

                // ADR-0038 D1: the reserved anchor row below the body, which the
                // plot-spanning shapes paint through so Excel resolves their bottom cell
                // anchor below the insert point. A THIRD reserved-row token, not the
                // padding height: this row is the sub-row anchor, and a margin-height row
                // here would be a visible strip of sheet rather than an anchor.
                ThemedMetric(config.Metrics, "ChartAnchorRowHeightPt"),

                // The Type column was read in step 1's single bulk Value2 call, so the
                // normaliser is handed that column rather than reading one cell per body
                // row through COM to decide the same heights. Same data, same refresh,
                // same decision.
                ProjectTypeTexts(rows.Rows));
            if (!heights.Succeeded)
            {
                // The Duration column has already been written at this point, so the
                // outcome reports it rather than claiming nothing was touched.
                return RefuseAfterWrites(
                    GanttRefreshRefusal.RowHeightRefused,
                    "The row heights could not be normalised.",
                    validation.Issues,
                    duration.CellsWritten);
            }

            rowHeightsWritten = heights.RowsWritten;
            measuredBodyHeights = heights.BodyHeightsPt;
        }

        // Step 9-10. Measure the live panel, then resolve the scene
        // inputs. The measurement is a read; the factory performs no host call
        // of its own.
        //
        // The body heights are passed through only when the normalisation above wrote
        // none of them: then they are still the sheet's state, and re-reading them is one
        // COM call per row for a value this refresh already holds. A pass that wrote a
        // height returns an empty list, so the measurement reads the host -- the sheet has
        // changed, and this is not a cache that may outlive the change.
        PanelGridOutcome grid = _panelMeasurement.Measure(
            ExcelSceneBuildRequestFactory.MeasuredColumns,
            measuredBodyHeights);
        if (!grid.Succeeded)
        {
            return RefuseAfterWrites(
                GanttRefreshRefusal.MeasurementRefused,
                "The worksheet columns could not be measured.",
                validation.Issues,
                duration.CellsWritten);
        }

        long measurementCompleted = Stopwatch.GetTimestamp();

        // The style registry comes from the configuration read, not from
        // GanttCatalogues. The config sheet is the single source of the effective
        // named styles — including any user-authored ones — and rebuilding a
        // registry from the built-in catalogue here would silently drop them, so a
        // custom style would validate and then fail to resolve at render time.
        SceneBuildRequestOutcome request = _requestFactory.Create(
            events,
            config.Settings,
            config.Styles,
            grid.Grid!,
            config.Metrics);
        if (!request.Succeeded)
        {
            return RefuseAfterWrites(
                GanttRefreshRefusal.SceneRequestRefused,
                request.Message!,
                validation.Issues,
                duration.CellsWritten);
        }

        // Step 11. Build the scene. A refusal here is still before any shape write.
        SceneBuildOutcome scene = SceneBuilder.TryBuild(request.Request);
        if (!scene.Succeeded)
        {
            return RefuseAfterWrites(
                GanttRefreshRefusal.SceneBuildRefused,
                "The chart could not be composed: " + scene.Refusal + ".",
                validation.Issues,
                duration.CellsWritten);
        }

        // Step 12. Translate the scene into host shape requests.
        SceneTranslationOutcome translated = _renderer.Translate(scene.Result!.Scene);
        if (translated.Refusals.Count > 0)
        {
            return RefuseAfterWrites(
                GanttRefreshRefusal.TranslationRefused,
                translated.Refusals.Count + " chart element(s) could not be drawn, so the chart was not rebuilt.",
                validation.Issues,
                duration.CellsWritten);
        }

        long buildCompleted = Stopwatch.GetTimestamp();

        // ADR-0043 D1: the live title is an owned textbox in the reserved row.
        // Read the title settings, compute the reserved-row bounds, and add the
        // title request to the desired set so the reconciler creates, updates,
        // or deletes it alongside the scene primitives. ShowTitle=false or a
        // blank title omits the request, so an owned ChartTitle shape is deleted
        // by the reconciler's Owned-minus-Desired difference.
        OfficeShapeRequest? titleRequest = null;
        string? chartTitle = config.Settings.TryGetValue("ChartTitle", out string? title) ? title : null;
        bool showTitle = !config.Settings.TryGetValue("ShowTitle", out string? showText)
            || !string.Equals(showText, "FALSE", StringComparison.Ordinal);
        if (showTitle && !string.IsNullOrWhiteSpace(chartTitle))
        {
            PanelCellGrid panelGrid = grid.Grid!;

            // The title occupies the RESERVED ROW exactly: the heading space
            // ADR-0043 gave it, and the same row that carries the year band over
            // the plot columns. The title and the year band therefore share one
            // height and one top edge.
            //
            // The height is the reserved row's MEASURED height, derived from the
            // grid, and NOT the TitleBandHeightPt token. The two disagreed - the
            // token is 24pt and the reserved row is normalised to
            // YearBandHeightPt (18pt) - so the textbox began 6pt ABOVE the
            // reserved row, in the top padding row, and the chart frame's top
            // edge (the reserved row's top) cut straight through the title
            // instead of enclosing it. One row has one authority: the row.
            double reservedRowHeightPt =
                panelGrid.OriginTopPt - panelGrid.HeaderHeightPt - panelGrid.TopPaddingHeightPt;
            RectD titleBounds = new(
                panelGrid.OriginLeftPt,
                panelGrid.TopPaddingHeightPt,
                panelGrid.Columns.Sum(column => column.WidthPt),
                reservedRowHeightPt);
            titleRequest = new OfficeShapeRequest(
                PrimitiveId: "ChartTitle",
                Kind: OfficeShapeKind.TextBox,
                Geometry: new OfficeShapeGeometry(Bounds: titleBounds),
                ZLayer: ZLayer.Frame,
                FontFamily: GanttCatalogues.LabelFontFamily,
                FontSizePt: 12.0,
                Bold: true,
                Text: chartTitle.Trim(),
                Alignment: GanttTextAlignment.Centre,
                TextColour: ColourHex.Parse("#404040"));
        }

        // Build the complete desired list: scene primitives plus the optional
        // title shape. The reconciler compares Desired against Owned, so an
        // omitted title request deletes an existing owned ChartTitle shape.
        List<OfficeShapeRequest> desired = [.. translated.Requests];
        if (titleRequest is { } req)
        {
            desired.Add(req);
        }

        // Step 13. The first and only shape mutation of the whole pipeline. Every
        // validation, measurement, and scene construction above has completed; from
        // here the reconciliation stops at the first refusal (D3).
        ShapeDeltaSnapshot? previousSnapshot = _useDeltaCache ? _shapeDeltaCache?.Snapshot : null;

        // The snapshot is a cheap first filter, not a verdict: the reconciler verifies
        // every candidate skip against the live shape's geometry and placement before
        // honouring it (see ShapeReconciler). That verification is what makes a
        // scheduled refresh converge to the manual-refresh state even after the host
        // moved or resized a shape, without forcing a full rewrite on every structural
        // change.
        TraceRefreshSummary(
            desired,
            previousSnapshot,
            outline,
            rowHeightsWritten,
            grid.Grid!.RowHeightsPt.Count);

        // Host writes are not transactional. Invalidate before entering the
        // mutation boundary so an unexpected exception cannot leave a cache
        // describing a partially changed shape collection.
        _shapeDeltaCache?.Invalidate();
        ShapeReconcileOutcome reconcile = ShapeReconciler.Reconcile(
            _shapeWriter,
            new ShapeReconcileRequest(desired, HasBlockingErrors: false)
            {
                PreviousSnapshot = previousSnapshot,
            });

        TraceReconcileResult(
            reconcile,
            desired.Count,
            new RefreshTimings(
                refreshStarted,
                readCompleted,
                validationCompleted,
                worksheetWritesCompleted,
                measurementCompleted,
                buildCompleted,
                Stopwatch.GetTimestamp()));

        if (reconcile.Succeeded)
        {
            _changeTracker?.Reset();
            _shapeDeltaCache?.Commit(desired);
            return GanttRefreshOutcome.Ok(validation.Issues, reconcile.CompletedCount, duration.CellsWritten, styleCellsWritten, scene.Result!.Scene.Warnings);
        }


        // A reconciliation that stopped after at least one write leaves the sheet in
        // a state the user can see: part old scene, part new. That is reported
        // distinctly rather than as a plain refusal, because "the existing chart is
        // unchanged" would be false. REV5 defers transactional Refresh to R4.10, so
        // the tear is surfaced, not undone.
        //
        // The failing primitive and the host's own reason are named, because
        // "Excel refused a change" on its own tells the user nothing they can
        // act on: a ZOrderRefused means every shape was written and only the
        // stacking is wrong, while an OperationRefused names the one shape the
        // host rejected.
        string detail = reconcile.Failure is { } failure
            ? $" The refused element was '{failure.PrimitiveId}' ({failure.Refusal})."
            : string.Empty;

        return reconcile.CompletedCount > 0
            ? GanttRefreshOutcome.Refused(
                GanttRefreshRefusal.PartialReconciliation,
                "The chart was partly updated before Excel refused a change ("
                    + reconcile.CompletedCount
                    + " element(s) written), so it may show a mixture of the old and new layout. Refresh again to finish."
                    + detail,
                validation.Issues,
                duration.CellsWritten,
                reconcile.CompletedCount,
                styleCellsWritten)
            : GanttRefreshOutcome.Refused(
                GanttRefreshRefusal.ReconciliationRefused,
                "Excel refused the first chart change, so the existing chart is unchanged."
                    + detail,
                validation.Issues,
                duration.CellsWritten);
    }

    /// <summary>
    /// Reads one metric from the workbook's effective metrics, falling back to the
    /// token catalogue's default when the workbook does not carry it.
    /// </summary>
    /// <param name="metrics">The workbook's effective metric values.</param>
    /// <param name="tokenName">The metric token name, for example <c>GanttRowHeightPt</c>.</param>
    /// <returns>The workbook's value, or the catalogue default.</returns>
    /// <remarks>
    /// The workbook is the authority for a metric the user has set; the catalogue is
    /// the authority for one it has not. A token added to the catalogue after a
    /// workbook was written therefore still normalises to a defined height rather
    /// than to nothing.
    /// </remarks>
    private static double ThemedMetric(IReadOnlyDictionary<string, double> metrics, string tokenName) =>
        metrics.TryGetValue(tokenName, out double value) ? value : GanttCatalogues.MetricDefault(tokenName);

    /// <summary>
    /// Writes one diagnostic line summarising the delta decision this run will take,
    /// plus one line per splitter band with its desired rectangle and whether its
    /// payload matches the cached signature.
    /// </summary>
    /// <param name="desired">The complete desired shape set, translated from the scene.</param>
    /// <param name="previousSnapshot">The committed signature snapshot, or null when bypassed.</param>
    /// <param name="outline">The outline result of this run.</param>
    /// <param name="rowHeightsWritten">How many row heights this run wrote.</param>
    /// <param name="gridRowCount">How many body rows the measurement returned.</param>
    /// <remarks>
    /// <para>
    /// <b>What this distinguishes.</b> A splitter band whose scene rectangle is
    /// unchanged after a move or collapse produces a byte-identical payload, so the
    /// delta cache skips its update. That is correct only while the live shape still
    /// matches. Excel shapes are created with <c>Placement = xlMoveAndSize</c>, so the
    /// host itself moves and resizes them when rows are inserted, hidden, un-hidden,
    /// or re-heighted — and nothing invalidates the snapshot for that. The band line
    /// separates "the scene moved the band" (<c>sig=rewrite</c>) from "the scene left
    /// it exactly where it was" (<c>sig=match</c>), and the structural counters in the
    /// summary line say whether this run wrote row structure at all.
    /// </para>
    /// <para>
    /// Invariant culture and fixed order, so two runs are comparable in a log.
    /// </para>
    /// </remarks>
    private void TraceRefreshSummary(
        List<OfficeShapeRequest> desired,
        ShapeDeltaSnapshot? previousSnapshot,
        OutlineGroupOutcome outline,
        int rowHeightsWritten,
        int gridRowCount)
    {
        if (_trace is null)
        {
            return;
        }

        int bands = 0;
        int skippedBands = 0;
        foreach (OfficeShapeRequest request in desired)
        {
            if (!request.PrimitiveId.EndsWith(_splitterBandRoleSuffix, StringComparison.Ordinal))
            {
                continue;
            }

            bands++;
            bool skipped = previousSnapshot is not null
                && previousSnapshot.Signatures.TryGetValue(request.PrimitiveId, out string? signature)
                && string.Equals(signature, ShapeDeltaCache.Signature(request), StringComparison.Ordinal);
            if (skipped)
            {
                skippedBands++;
            }

            RectD? bounds = request.Geometry.Bounds;
            string rect = bounds is { } value
                ? string.Create(
                    CultureInfo.InvariantCulture,
                    $"x={value.X:0.###} y={value.Y:0.###} w={value.Width:0.###} h={value.Height:0.###}")
                : "bounds=none";
            TraceRefresh(string.Create(
                CultureInfo.InvariantCulture,
                $"gantt.refresh band delta={_useDeltaCache} id={request.PrimitiveId} {rect} sig={(skipped ? "match" : "rewrite")}"));
        }

        TraceRefresh(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.refresh delta={_useDeltaCache} snapshot={previousSnapshot?.Signatures.Count ?? 0} desired={desired.Count} bands={bands} bands-skipped={skippedBands} outline-groups={outline.GroupsApplied} outline-ungrouped={outline.RowsUngrouped} heights-written={rowHeightsWritten} grid-rows={gridRowCount}"));
    }

    /// <summary>Timestamp boundaries of one refresh, for the diagnostic timing line.</summary>
    /// <param name="Started">When the pipeline started.</param>
    /// <param name="Read">After the table/configuration reads.</param>
    /// <param name="Validated">After row validation.</param>
    /// <param name="WorksheetWrites">After the outline, duration, StyleKey and row-height writes.</param>
    /// <param name="Measured">After the panel measurement.</param>
    /// <param name="Built">After the scene was built and translated.</param>
    /// <param name="Written">After the reconciliation completed.</param>
    private readonly record struct RefreshTimings(
        long Started,
        long Read,
        long Validated,
        long WorksheetWrites,
        long Measured,
        long Built,
        long Written);

    /// <summary>
    /// Writes one diagnostic line naming what the reconciliation actually planned and
    /// completed, plus the phase timings of the pass.
    /// </summary>
    /// <param name="reconcile">The reconciliation result.</param>
    /// <param name="desiredCount">How many shapes the scene declared.</param>
    /// <param name="timings">The phase boundaries of this pass.</param>
    /// <remarks>
    /// The timings are what separate "the refresh is slow because the pipeline is
    /// expensive" from "the refresh is slow because the host writes are expensive" — the
    /// distinction the last optimisation decision turned on. Each figure is the elapsed
    /// milliseconds for one phase, in pipeline order.
    /// </remarks>
    private void TraceReconcileResult(
        ShapeReconcileOutcome reconcile,
        int desiredCount,
        RefreshTimings timings)
    {
        if (_trace is null)
        {
            return;
        }

        ShapeReconcilePlan plan = reconcile.Plan;

        // geom= is the count of move-only updates (ADR-0053). It is the figure that says how
        // much of a pass was a relocation rather than a re-apply, and it is the only way to see
        // from the log that a promote moved 40 primitives instead of restyling them.
        int geometryOnly = plan.Operations.Count(operation =>
            operation.Action == ShapeReconcileAction.Update && operation.Scope == ShapeUpdateScope.Geometry);
        TraceRefresh(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.refresh result ok={reconcile.Succeeded} desired={desiredCount} updated={plan.Of(ShapeReconcileAction.Update).Count} geom={geometryOnly} created={plan.Of(ShapeReconcileAction.Create).Count} deleted={plan.Of(ShapeReconcileAction.Delete).Count} written={reconcile.CompletedCount} zorder={reconcile.ZOrderApplied} zorderms={reconcile.ZOrderMilliseconds:0.#} refusal={reconcile.Refusal?.ToString() ?? "none"}"));
        TraceRefresh(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.refresh timings ms read={ElapsedMs(timings.Started, timings.Read):0.#} validate={ElapsedMs(timings.Read, timings.Validated):0.#} sheet={ElapsedMs(timings.Validated, timings.WorksheetWrites):0.#} measure={ElapsedMs(timings.WorksheetWrites, timings.Measured):0.#} build={ElapsedMs(timings.Measured, timings.Built):0.#} write={ElapsedMs(timings.Built, timings.Written):0.#} total={ElapsedMs(timings.Started, timings.Written):0.#}"));
    }

    /// <summary>Elapsed milliseconds between two <see cref="Stopwatch"/> timestamps.</summary>
    private static double ElapsedMs(long from, long to) =>
        Stopwatch.GetElapsedTime(from, to).TotalMilliseconds;

    /// <summary>Writes one trace line, or does nothing when no sink is wired.</summary>
    /// <param name="message">The line to write.</param>
    private void TraceRefresh(string message) => _trace?.Invoke(message);

    /// <summary>
    /// Mirrors the written <c>StyleKey</c> cells into the validated events so the
    /// outline, duration, and scene steps render from the corrected keys rather than
    /// the stale keys read from the worksheet at step 1. Without this a type change
    /// takes two Refreshes to take effect.
    /// </summary>
    /// <param name="events">The validated events, each carrying the StyleKey read
    /// from column K at step 1.</param>
    /// <param name="writes">The StyleKey writes just applied to the worksheet.</param>
    /// <returns>
    /// The events with each written row's StyleKey replaced by the value now on the
    /// sheet, or <paramref name="events"/> unchanged when nothing was written.
    /// </returns>
    private static IReadOnlyList<GanttEvent> ApplyStyleKeyToEvents(
        IReadOnlyList<GanttEvent> events,
        IReadOnlyList<StyleKeyWrite> writes)
    {
        if (writes.Count == 0)
        {
            return events;
        }

        Dictionary<int, string?> byRow = new Dictionary<int, string?>(writes.Count);
        foreach (StyleKeyWrite write in writes)
        {
            byRow[write.RowNumber] = write.NewStyleKey;
        }

        List<GanttEvent> result = new List<GanttEvent>(events.Count);
        foreach (GanttEvent @event in events)
        {
            if (byRow.TryGetValue(@event.RowNumber, out string? key))
            {
                result.Add(@event with { StyleKey = key });
            }
            else
            {
                result.Add(@event);
            }
        }

        return result;
    }

    /// <summary>
    /// Reads the <c>Duration</c> text each row currently shows, so the planner can
    /// exclude unchanged rows instead of rewriting the whole column.
    /// </summary>
    /// <summary>
    /// Projects the already-read <c>Type</c> texts into a body-index-ordered list so the
    /// row-height normaliser does not re-read each body row's Type cell through COM.
    /// </summary>
    /// <param name="rows">The rows just read, in body order.</param>
    /// <returns>
    /// One entry per body index, <see langword="null"/> where a row is missing or its
    /// Type cell is blank or unreadable; the normaliser reads those from the column.
    /// </returns>
    /// <remarks>
    /// The table reader has already read the whole body -- <c>Type</c> included -- in a
    /// single bulk <c>Value2</c> call. Passing that column to the normaliser keeps the
    /// height decision identical while removing one parameterised COM read per body row
    /// from every refresh. A blank entry becomes <see langword="null"/>, which both this
    /// projection and the normaliser's own read treat as "managed row": the two paths
    /// reach the same decision from the same data.
    /// </remarks>
    private static string?[] ProjectTypeTexts(IReadOnlyList<GanttRowDto> rows)
    {
        if (rows.Count == 0)
        {
            return [];
        }

        // Body numbers are one-based and normally contiguous, but the array is sized
        // from the highest number actually read: a gap then stays null rather than
        // shifting every later row's type onto the wrong index.
        int highestRowNumber = 0;
        foreach (GanttRowDto row in rows)
        {
            if (row.RowNumber > highestRowNumber)
            {
                highestRowNumber = row.RowNumber;
            }
        }

        string?[] texts = new string?[highestRowNumber];
        foreach (GanttRowDto row in rows)
        {
            int index = row.RowNumber - 1;
            if (index >= 0 && index < texts.Length)
            {
                texts[index] = row.TypeCell.HasValue ? row.TypeCell.Value : null;
            }
        }

        return texts;
    }

    private static Dictionary<int, string?> ReadCurrentDurations(IReadOnlyList<GanttRowDto> rows)
    {
        Dictionary<int, string?> current = [];
        foreach (GanttRowDto row in rows)
        {
            current[row.RowNumber] = row.DurationCell.HasValue ? row.DurationCell.Value : null;
        }

        return current;
    }

    /// <summary>
    /// Validation issue codes that are genuinely row-level: they affect the
    /// entire row's identity or classification, so the whole row is filled.
    /// Everything else is a cell-level data issue and highlights only the
    /// offending cell.
    /// </summary>
    private static readonly HashSet<string> _rowLevelIssueCodes = new(StringComparer.Ordinal)
    {
        GanttValidationCodes.IdMissingOrMalformed,
        GanttValidationCodes.DuplicateId,
        GanttValidationCodes.UnknownType,
    };

    /// <summary>
    /// Brings the row/cell highlight fills into line with the current validation findings.
    /// Uses <see cref="IRowWarningHighlighter.ReconcileHighlights"/> so only the
    /// minimal set of COM writes is performed: clear addresses that no longer
    /// need highlighting, fill addresses that newly need it, and skip
    /// addresses unchanged from the previous refresh.
    /// </summary>
    /// <param name="issues">The validation issues just found, in deterministic order.</param>
    /// <returns>The highlighter's outcome; a success with zero when no highlighter is wired.</returns>
    /// <remarks>
    /// An EMPTY issue list is still passed through rather than skipped, because
    /// <see cref="IRowWarningHighlighter.ReconcileHighlights"/> with an empty list
    /// clears all previous highlights: that is the call that makes a corrected row
    /// lose its fill. Skipping it when the list is empty would leave stale fills
    /// behind.
    /// <para>
    /// Unlike the previous <see cref="IGanttValidationReporter"/>, no
    /// <c>SpecialCells</c> scan is performed — fills are cosmetic cell-interior writes
    /// and L21 is closed.
    /// </para>
    /// </remarks>
    private GanttValidationReportOutcome ReportValidation(IReadOnlyList<GanttValidationIssue> issues)
    {
        if (_rowWarningHighlighter is null)
        {
            return GanttValidationReportOutcome.Ok(0);
        }

        // Build targets from every current issue and reconcile the entire
        // previous highlight set. A changed-row scope could leave stale
        // highlights on rows whose state was affected by another row's edit.
        IReadOnlyList<int> rowsToConsider = [.. issues
            .Select(i => i.RowNumber)
            .Distinct()
            .Order()];

        List<HighlightTarget> targets = [];
        foreach (int row in rowsToConsider)
        {
            IReadOnlyList<GanttValidationIssue> rowIssues = [.. issues.Where(i => i.RowNumber == row)];

            // A row-level blocking error fills the entire row pastel red.
            bool hasRowLevelError = rowIssues.Any(
                i => i.Severity == GanttValidationSeverity.Error &&
                     _rowLevelIssueCodes.Contains(i.Code));
            if (hasRowLevelError)
            {
                targets.Add(new HighlightTarget(row, null, GanttValidationSeverity.Error));
                continue;
            }

            // Cell-level issues highlight only the offending cell.
            foreach (GanttValidationIssue issue in rowIssues)
            {
                GanttValidationSeverity severity = issue.Severity;
                int columnIndex = ResolveColumnIndex(issue.Field);
                if (columnIndex < 1)
                {
                    // Field not in schema — anchor on the Id column.
                    columnIndex = 1;
                }
                targets.Add(new HighlightTarget(row, columnIndex, severity));
            }
        }

        // Single differential reconciliation: clear stale fills, write new fills,
        // skip unchanged addresses. No full-table clear-then-rewrite.
        return _rowWarningHighlighter.ReconcileHighlights(targets);
    }

    /// <summary>
    /// Resolves the 1-based body-column index for a validation field name,
    /// or -1 when the field is not in the table schema.
    /// </summary>
    private static int ResolveColumnIndex(string fieldName)
    {
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        for (int i = 0; i < columns.Count; i++)
        {
            if (string.Equals(columns[i].Name, fieldName, StringComparison.Ordinal))
            {
                return i + 1;
            }
        }
        return -1;
    }

    private static GanttRefreshOutcome Refuse(GanttRefreshRefusal refusal, string message) =>
        GanttRefreshOutcome.Refused(refusal, message);

    /// <summary>
    /// Refuses with the worksheet writes that already landed before this point.
    /// </summary>
    /// <param name="refusal">Why the refresh did not complete.</param>
    /// <param name="message">The explanation shown to the user.</param>
    /// <param name="issues">The validation issues found.</param>
    /// <param name="durationCellsWritten">
    /// How many <c>Duration</c> cells were written before the refusal.
    /// </param>
    /// <param name="shapesWritten">How many shape operations ran before the refusal.</param>
    /// <param name="styleKeyCellsWritten">How many <c>StyleKey</c> cells were written before the refusal.</param>
    /// <returns>The refusal outcome, carrying the real counts.</returns>
    /// <remarks>
    /// <b>Why the counts cannot stay zero.</b> The class remark promises the caller can
    /// tell "the chart was not touched" from "the sheet was not touched at all". A
    /// plain <see cref="GanttRefreshOutcome.Refused"/> hardcodes both counts to zero,
    /// so a refusal that happened <em>after</em> the Duration write reported that
    /// nothing had been written when the Duration column had in fact been changed —
    /// untrue, and the kind of untruth a user hits as a workbook that is dirty for no
    /// stated reason. Every refusal reached after a mutation therefore reports what it
    /// actually did.
    /// </remarks>
    private static GanttRefreshOutcome RefuseAfterWrites(
        GanttRefreshRefusal refusal,
        string message,
        IReadOnlyList<GanttValidationIssue> issues,
        int durationCellsWritten,
        int shapesWritten = 0,
        int styleKeyCellsWritten = 0) =>
        GanttRefreshOutcome.Refused(refusal, message, issues, durationCellsWritten, shapesWritten, styleKeyCellsWritten);

    private static string Describe(GanttTableReadRefusalReason refusal) =>
        refusal switch
        {
            GanttTableReadRefusalReason.NoActiveWorkbook =>
                "There is no active workbook to refresh.",
            GanttTableReadRefusalReason.TableMissing =>
                "The Gantt table was not found on the active sheet.",
            GanttTableReadRefusalReason.DateSystemUnsupported =>
                "This workbook uses the 1904 date system, which Gantt Creator does not support.",
            _ => "The Gantt table could not be read.",
        };
}
