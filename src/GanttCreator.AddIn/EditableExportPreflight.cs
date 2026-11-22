using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Builds the export scene request from the workbook's measured inputs (the R6.3
/// <c>CreateExport</c> seam), as a delegate so the command test can pin the export variant
/// without a host.
/// </summary>
/// <param name="events">The validated events to compose.</param>
/// <param name="settings">The workbook's effective settings map.</param>
/// <param name="registry">The workbook's style registry.</param>
/// <param name="grid">The measured panel grid.</param>
/// <param name="themedMetrics">The workbook's themed metric values.</param>
/// <returns>The export request, or a typed refusal.</returns>
internal delegate SceneBuildRequestOutcome ExportRequestBuild(
    IReadOnlyList<GanttEvent> events,
    IReadOnlyDictionary<string, string> settings,
    GanttStyleRegistry registry,
    PanelCellGrid grid,
    IReadOnlyDictionary<string, double> themedMetrics);

/// <summary>Resolves the export bounds and scene (the R6.1 <c>TryResolve</c> seam).</summary>
/// <param name="request">The export request the scene is built from.</param>
/// <param name="panelTheme">The resolved panel styles, or null for chart-only.</param>
/// <param name="selection">The export selection the user's setting chose.</param>
/// <returns>The bounds resolution carrying the scene, or a typed refusal.</returns>
internal delegate ExportBoundsOutcome ExportBoundsResolve(
    SceneBuildRequest? request,
    PanelTheme? panelTheme,
    ExportSelection selection);

/// <summary>Why the shared editable-export pre-flight produced no preparation (R7.6 D4).</summary>
/// <remarks>
/// <para>
/// Every member is a typed, expected outcome rather than an exception. The construction
/// members exist so a dependency the pre-flight cannot run without is reported, never
/// dereferenced, and the stage members name the first refusal in the pipeline order.
/// </para>
/// <para>
/// <b>The pre-flight owns the read/validate/measure/resolve chain, never the message text.</b>
/// Each consuming command translates the typed refusal into its own user-facing line, so a
/// Copy Editable message never names the wrong command (the R7.6 review's F1 finding) and the
/// established Copy Editable strings stay byte-identical.
/// </para>
/// </remarks>
internal enum EditableExportPreflightRefusal
{
    /// <summary>No pre-flight request was supplied.</summary>
    NullRequest = 0,

    /// <summary>No table reader was supplied.</summary>
    MissingTableReader = 1,

    /// <summary>No configuration reader was supplied.</summary>
    MissingConfigReader = 2,

    /// <summary>No panel-grid measurement port was supplied.</summary>
    MissingGridMeasurement = 3,

    /// <summary>No export-request factory was supplied.</summary>
    MissingExportRequestFactory = 4,

    /// <summary>No bounds resolver was supplied.</summary>
    MissingBoundsResolver = 5,

    /// <summary>The table could not be read.</summary>
    TableReadRefused = 6,

    /// <summary>The workbook configuration could not be read.</summary>
    ConfigReadRefused = 7,

    /// <summary>The live panel grid could not be measured.</summary>
    MeasurementRefused = 8,

    /// <summary>The export scene request was refused.</summary>
    ExportRequestRefused = 9,

    /// <summary>The export bounds or scene were refused, including the post-resolution guard.</summary>
    BoundsRefused = 10,
}

/// <summary>
/// The resolved export inputs both export commands consume (R7.6 D4/D5): the validated
/// events, the measured grid, the resolved bounds with their scene, and the placement inset.
/// </summary>
/// <param name="Events">The validated events the export scene is composed from.</param>
/// <param name="Grid">The measured live panel grid the export panel reproduces.</param>
/// <param name="Resolution">
/// The R6.1 bounds resolution carrying the scene and the zero-origin
/// <see cref="ExportBounds.Bounds"/>; already guarded to a finite, strictly positive extent.
/// </param>
/// <param name="InsetPt">
/// The slide-placement inset: the themed <c>ChartOuterPaddingPt</c> token expressed through
/// <see cref="ChartPaddingPt.Uniform(double)"/> (D-R7.4). Scene-space, so the fit scales it
/// with the group.
/// </param>
internal sealed record EditableExportPreparation(
    IReadOnlyList<GanttEvent> Events,
    PanelCellGrid Grid,
    ExportBoundsOutcome Resolution,
    double InsetPt);

/// <summary>The typed result of one shared editable-export pre-flight run (R7.6 D4).</summary>
/// <param name="Preparation">The resolved preparation, or <see langword="null"/> on a refusal.</param>
/// <param name="Refusal">The typed pre-flight refusal, or <see langword="null"/> on success.</param>
/// <param name="TableReadRefusal">The reader's own refusal for a <see cref="EditableExportPreflightRefusal.TableReadRefused"/>.</param>
/// <param name="ConfigRefusal">The reader's own refusal for a <see cref="EditableExportPreflightRefusal.ConfigReadRefused"/>.</param>
/// <param name="MeasurementRefusal">The port's own refusal for a <see cref="EditableExportPreflightRefusal.MeasurementRefused"/>.</param>
/// <param name="ExportRequestMessage">
/// The export factory's actionable message for an
/// <see cref="EditableExportPreflightRefusal.ExportRequestRefused"/>, or <see langword="null"/>.
/// </param>
/// <param name="BoundsOutcome">
/// The bounds outcome carrying the refusal for a <see cref="EditableExportPreflightRefusal.BoundsRefused"/>,
/// including the post-resolution guard's synthetic <see cref="ExportBoundsRefusal.InvalidBounds"/>
/// outcome, so each command reuses its existing bounds message mapping.
/// </param>
/// <param name="DroppedRowCount">
/// How many body rows carried a blocking validation error and were left out of the partial
/// export, or zero when nothing was dropped. Carried on refusals too, because a measurement or
/// bounds refusal happens after validation has already counted them.
/// </param>
internal sealed record EditableExportPreflightOutcome(
    EditableExportPreparation? Preparation,
    EditableExportPreflightRefusal? Refusal,
    GanttTableReadRefusalReason? TableReadRefusal = null,
    ConfigReadRefusalReason? ConfigRefusal = null,
    PanelGridRefusalReason? MeasurementRefusal = null,
    string? ExportRequestMessage = null,
    ExportBoundsOutcome? BoundsOutcome = null,
    int DroppedRowCount = 0)
{
    /// <summary>Gets whether the pre-flight produced a usable preparation.</summary>
    internal bool Succeeded => Preparation is not null;

    /// <summary>Creates the success outcome.</summary>
    /// <param name="preparation">The resolved preparation.</param>
    /// <param name="droppedRowCount">The blocking-error row count.</param>
    /// <returns>The success outcome.</returns>
    internal static EditableExportPreflightOutcome Ok(
        EditableExportPreparation preparation,
        int droppedRowCount) =>
        new(preparation, null, DroppedRowCount: droppedRowCount);

    /// <summary>Creates the refusal outcome.</summary>
    /// <param name="refusal">The typed refusal.</param>
    /// <param name="tableReadRefusal">The reader's refusal, when one exists.</param>
    /// <param name="configRefusal">The configuration reader's refusal, when one exists.</param>
    /// <param name="measurementRefusal">The measurement port's refusal, when one exists.</param>
    /// <param name="exportRequestMessage">The export factory's message, when one exists.</param>
    /// <param name="boundsOutcome">The bounds outcome, when one exists.</param>
    /// <param name="droppedRowCount">The blocking-error row count, when it is known.</param>
    /// <returns>The refusal outcome.</returns>
    internal static EditableExportPreflightOutcome Refused(
        EditableExportPreflightRefusal refusal,
        GanttTableReadRefusalReason? tableReadRefusal = null,
        ConfigReadRefusalReason? configRefusal = null,
        PanelGridRefusalReason? measurementRefusal = null,
        string? exportRequestMessage = null,
        ExportBoundsOutcome? boundsOutcome = null,
        int droppedRowCount = 0) =>
        new(
            null,
            refusal,
            tableReadRefusal,
            configRefusal,
            measurementRefusal,
            exportRequestMessage,
            boundsOutcome,
            droppedRowCount);
}

/// <summary>The inputs one shared editable-export pre-flight consumes.</summary>
/// <param name="TableReader">Reads every body row into neutral DTOs. Required.</param>
/// <param name="ConfigReader">Reads the workbook's effective settings, styles and metrics. Required.</param>
/// <param name="GridMeasurement">Measures the live data panel. Required.</param>
/// <param name="ExportRequestFactory">Builds the export scene request. Required.</param>
/// <param name="BoundsResolver">Resolves the export bounds and scene. Required.</param>
/// <param name="TechnicalRecord">
/// Receives the dropped-row record, or <see langword="null"/> to discard it.
/// </param>
internal sealed record EditableExportPreflightRequest(
    IGanttTableReader? TableReader,
    IConfigCatalogueReader? ConfigReader,
    IPanelGridMeasurementPort? GridMeasurement,
    ExportRequestBuild? ExportRequestFactory,
    ExportBoundsResolve? BoundsResolver,
    Action<string>? TechnicalRecord = null);

/// <summary>
/// R7.6 D4: the one authority that reads, validates, measures, resolves and guards the
/// editable export, shared by Copy Editable and Send to PowerPoint.
/// </summary>
/// <remarks>
/// <para>
/// <b>The order is the pipeline order and it is the defect this row closes.</b> Read →
/// configuration → validate → measure → export request → selection → bounds. The
/// configuration precedes validation because the validator needs the workbook's style
/// registry to resolve a style-key row's capabilities; validating first silently dropped a
/// Custom Activity / Custom Milestone / Delineator row's own <c>LabelPosition</c> (and its
/// colour overrides), so the copy rendered the style's default position while Refresh
/// rendered the row's (the reported 2026-11-19 label defect).
/// </para>
/// <para>
/// <b>The post-resolution guard is this row's Step 0 obligation.</b> The resolver already
/// refuses a non-positive <c>request.PlotBounds</c>, but the resolved
/// <see cref="ExportBounds.ChartBounds"/> was not checked. Before any staging shape exists,
/// a resolved bounds whose zero-origin width or height is not finite and strictly positive is
/// refused with the existing <see cref="ExportBoundsRefusal.InvalidBounds"/>, so the R7.4 fit
/// never reaches its degenerate arm. <see cref="RectD"/> already refuses non-finite or
/// negative extents at construction, so the reachable arm is a zero extent.
/// </para>
/// <para>
/// <b>Nothing here creates a shape.</b> Every refusal leaves the workbook untouched, so no
/// cleanup is owed; the lifecycle's <c>finally</c> remains the only staging cleanup authority.
/// </para>
/// </remarks>
internal static class EditableExportPreflight
{
    /// <summary>
    /// The settings key that selects the export composition (catalogue default <c>TRUE</c>).
    /// The key and its default are owned by <see cref="GanttCatalogues"/>; this constant names
    /// the one read, and an explicit case-insensitive <c>FALSE</c> is the only chart-only
    /// request.
    /// </summary>
    internal const string ExportIncludeDataPanelSettingKey = "ExportIncludeDataPanel";

    /// <summary>The themed metric token the placement inset reads (D-R7.4).</summary>
    internal const string OuterPaddingToken = "ChartOuterPaddingPt";

    /// <summary>Runs the shared pre-flight over the supplied seams.</summary>
    /// <param name="request">The run's inputs; see <see cref="EditableExportPreflightRequest"/>.</param>
    /// <returns>The preparation or the typed refusal; every expected state is a result.</returns>
    internal static EditableExportPreflightOutcome Run(EditableExportPreflightRequest? request)
    {
        if (request is null)
        {
            return EditableExportPreflightOutcome.Refused(EditableExportPreflightRefusal.NullRequest);
        }

        // The dependency checks run in pipeline order, so the first missing member names the
        // stage that would have failed. Every check is a positive-test target.
        if (request.TableReader is not { } tableReader)
        {
            return EditableExportPreflightOutcome.Refused(EditableExportPreflightRefusal.MissingTableReader);
        }

        if (request.ConfigReader is not { } configReader)
        {
            return EditableExportPreflightOutcome.Refused(EditableExportPreflightRefusal.MissingConfigReader);
        }

        if (request.GridMeasurement is not { } gridMeasurement)
        {
            return EditableExportPreflightOutcome.Refused(EditableExportPreflightRefusal.MissingGridMeasurement);
        }

        if (request.ExportRequestFactory is not { } exportRequestFactory)
        {
            return EditableExportPreflightOutcome.Refused(EditableExportPreflightRefusal.MissingExportRequestFactory);
        }

        if (request.BoundsResolver is not { } boundsResolver)
        {
            return EditableExportPreflightOutcome.Refused(EditableExportPreflightRefusal.MissingBoundsResolver);
        }

        // Step 1 — read the full table. Read-only; nothing was mutated on a refusal.
        GanttTableReadOutcome read = tableReader.Read();
        if (!read.Succeeded)
        {
            return EditableExportPreflightOutcome.Refused(
                EditableExportPreflightRefusal.TableReadRefused,
                tableReadRefusal: read.Refusal);
        }

        // Step 2 — read the workbook configuration (settings, styles, metrics). It precedes
        // validation because the validator must see the workbook's style registry; without it a
        // style-key type cannot resolve its capabilities and the row's own LabelPosition and
        // colour overrides are silently dropped (2026-11-19 label-parity fix).
        ConfigReadOutcome config = configReader.Read();
        if (!config.Succeeded)
        {
            // No validation has run yet, so no row has been dropped: zero, never inferred.
            return EditableExportPreflightOutcome.Refused(
                EditableExportPreflightRefusal.ConfigReadRefused,
                configRefusal: config.Refusal);
        }

        // Step 3 — validate with the workbook's registry. Rows with blocking errors are
        // excluded from the export's event set by the validator's own contract; the count is
        // surfaced, never silently dropped (owner ruling 2026-11-19).
        GanttValidationOutcome validation = GanttRowValidator.Validate(read.Rows, config.Styles);
        int droppedRows = validation.Issues
            .Where(issue => issue.Severity == GanttValidationSeverity.Error)
            .Select(issue => issue.RowNumber)
            .Distinct()
            .Count();
        if (droppedRows > 0)
        {
            // The tag is preserved from the R6.8 log contract; the count is schedule-free.
            request.TechnicalRecord?.Invoke(string.Create(
                CultureInfo.InvariantCulture,
                $"gantt.copy rows-dropped={droppedRows}"));
        }

        // Step 4 — measure the live panel. The single-argument overload reads the host's
        // current row heights, so the export path performs no normalisation writes.
        PanelGridOutcome grid = gridMeasurement.Measure(
            EditableExportRequest.MeasuredColumns(config.Settings));
        if (!grid.Succeeded || grid.Grid is not { } measuredGrid)
        {
            return EditableExportPreflightOutcome.Refused(
                EditableExportPreflightRefusal.MeasurementRefused,
                measurementRefusal: grid.Refusal,
                droppedRowCount: droppedRows);
        }

        // Step 5 — build the EXPORT scene request (EditableExport profile, no live fallback),
        // never the live request.
        SceneBuildRequestOutcome export = exportRequestFactory(
            validation.Events,
            config.Settings,
            config.Styles,
            measuredGrid,
            config.Metrics);
        if (!export.Succeeded || export.Request is not { } exportRequest)
        {
            return EditableExportPreflightOutcome.Refused(
                EditableExportPreflightRefusal.ExportRequestRefused,
                exportRequestMessage: export.Message,
                droppedRowCount: droppedRows);
        }

        // Step 6 — the export selection rides the persisted inclusion setting; the panel
        // theme is supplied only for table-plus-chart.
        ExportSelection selection = SelectExportSelection(config.Settings);
        ExportBoundsOutcome resolution = boundsResolver(
            exportRequest,
            selection == ExportSelection.TablePlusChart
                ? ExportPanelTheme.Resolve(config.Metrics)
                : null,
            selection);
        if (!resolution.Succeeded)
        {
            return EditableExportPreflightOutcome.Refused(
                EditableExportPreflightRefusal.BoundsRefused,
                boundsOutcome: resolution,
                droppedRowCount: droppedRows);
        }

        // Step 6b — the post-resolution guard (R7.6 Step 0). Any staging shape created from a
        // degenerate bounds would be uncreatable or unfittable; refuse before the lifecycle.
        return GuardResolvedBounds(resolution, droppedRows)
            ?? EditableExportPreflightOutcome.Ok(
                new EditableExportPreparation(
                    validation.Events,
                    measuredGrid,
                    resolution,
                    ResolveInsetPt(config.Metrics)),
                droppedRows);
    }

    /// <summary>
    /// Refuses a resolved bounds whose zero-origin extent is not finite and strictly positive.
    /// </summary>
    /// <param name="resolution">The successful resolution to guard.</param>
    /// <param name="droppedRows">The blocking-error row count to carry on a refusal.</param>
    /// <returns>The refusal outcome, or <see langword="null"/> when the bounds is usable.</returns>
    /// <remarks>
    /// The refusal reuses the existing <see cref="ExportBoundsRefusal.InvalidBounds"/> and its
    /// established message, so each command's bounds translation stays the single authority.
    /// </remarks>
    private static EditableExportPreflightOutcome? GuardResolvedBounds(
        ExportBoundsOutcome resolution,
        int droppedRows) =>
        IsUsableBounds(resolution.Bounds!.Bounds)
            ? null
            : EditableExportPreflightOutcome.Refused(
                EditableExportPreflightRefusal.BoundsRefused,
                boundsOutcome: new ExportBoundsOutcome(null, ExportBoundsRefusal.InvalidBounds, null),
                droppedRowCount: droppedRows);

    /// <summary>Maps the stored inclusion setting to the export selection.</summary>
    /// <param name="settings">The workbook's effective settings map.</param>
    /// <returns>Chart-only on an explicit case-insensitive <c>FALSE</c>; otherwise table-plus-chart.</returns>
    /// <remarks>
    /// Fail-open toward the catalogue default (<c>TRUE</c>), matching the automatic-refresh
    /// parse precedent: a typo cannot silently change what an export covers, and the catalogue
    /// remains the key's authority.
    /// </remarks>
    internal static ExportSelection SelectExportSelection(IReadOnlyDictionary<string, string> settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return settings.TryGetValue(ExportIncludeDataPanelSettingKey, out string? stored)
            && string.Equals(stored?.Trim(), "FALSE", StringComparison.OrdinalIgnoreCase)
                ? ExportSelection.ChartOnly
                : ExportSelection.TablePlusChart;
    }

    /// <summary>Resolves the scene-space placement inset from the themed metrics (D-R7.4).</summary>
    /// <param name="themedMetrics">The workbook's themed metrics, possibly empty.</param>
    /// <returns>The uniform inset in points.</returns>
    /// <remarks>
    /// The value is the <c>ChartOuterPaddingPt</c> token expressed as
    /// <see cref="ChartPaddingPt.Uniform(double)"/>. A missing or unusable value falls back to
    /// the catalogue default rather than refusing: the inset is presentation, and the fit's own
    /// frame guard still refuses a degenerate result.
    /// </remarks>
    internal static double ResolveInsetPt(IReadOnlyDictionary<string, double> themedMetrics)
    {
        ArgumentNullException.ThrowIfNull(themedMetrics);
        double padding = themedMetrics.TryGetValue(OuterPaddingToken, out double stored)
            && double.IsFinite(stored)
            && stored >= 0
                ? stored
                : GanttCatalogues.MetricDefault(OuterPaddingToken);
        return ChartPaddingPt.Uniform(padding).MaxPt;
    }

    /// <summary>Whether a resolved bounds is usable for a composition.</summary>
    /// <param name="bounds">The zero-origin export bounds.</param>
    /// <returns><see langword="true"/> when both extents are finite and strictly positive.</returns>
    /// <remarks>
    /// The finite checks are belt-and-braces: <see cref="RectD"/> refuses a non-finite or
    /// negative extent at construction, so only a zero extent can reach here from the resolver.
    /// </remarks>
    internal static bool IsUsableBounds(RectD bounds) =>
        double.IsFinite(bounds.Width)
        && bounds.Width > 0
        && double.IsFinite(bounds.Height)
        && bounds.Height > 0;
}
