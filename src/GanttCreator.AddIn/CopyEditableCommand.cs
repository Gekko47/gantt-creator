using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>Runs the staging lifecycle (the R6.5 <c>Run</c> seam).</summary>
/// <param name="request">The assembled lifecycle request.</param>
/// <returns>The lifecycle outcome, with residue named whenever cleanup was incomplete.</returns>
internal delegate StagingLifecycleOutcome StagingLifecycleRun(StagingLifecycleRequest? request);

/// <summary>Why the Copy Editable command produced no verified copy (R6.8).</summary>
/// <remarks>
/// Every member is a typed, expected outcome rather than an exception
/// (<c>docs/02-ARCHITECTURE.md</c> "Error handling"). The construction members exist so a
/// dependency the command cannot run without is reported, never dereferenced.
/// </remarks>
internal enum CopyEditableRefusal
{
    /// <summary>No command request was supplied.</summary>
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

    /// <summary>No staging lifecycle was supplied.</summary>
    MissingLifecycle = 6,

    /// <summary>No application-state scope was supplied.</summary>
    MissingStateScope = 7,

    /// <summary>No staging shape write port was supplied.</summary>
    MissingStagingWriter = 8,

    /// <summary>No grouping port was supplied.</summary>
    MissingGroupingPort = 9,

    /// <summary>No shape copy port was supplied.</summary>
    MissingCopyPort = 10,

    /// <summary>No clipboard format reader was supplied.</summary>
    MissingClipboardReader = 11,

    /// <summary>The table could not be read, so no copy was attempted.</summary>
    TableReadRefused = 12,

    /// <summary>The workbook configuration could not be read, so no copy was attempted.</summary>
    ConfigReadRefused = 13,

    /// <summary>The live panel grid could not be measured, so no copy was attempted.</summary>
    MeasurementRefused = 14,

    /// <summary>The export scene request was refused, so no copy was attempted.</summary>
    ExportRequestRefused = 15,

    /// <summary>The export bounds or scene were refused, so nothing was staged.</summary>
    BoundsRefused = 16,

    /// <summary>The staging composition refused; its shapes were cleaned up.</summary>
    CompositionRefused = 17,

    /// <summary>The staged set could not be grouped; the children were restored.</summary>
    GroupingRefused = 18,

    /// <summary>The verified copy refused (see the lifecycle outcome's copy result).</summary>
    CopyRefused = 19,

    /// <summary>Cleanup reported residue, which is a failure and never a success.</summary>
    CleanupResidue = 20,

    /// <summary>The lifecycle ran cleanly but the copy seam never ran (empty composition).</summary>
    NothingCopied = 21,

    /// <summary>The lifecycle refused for a construction reason after the command's checks.</summary>
    LifecycleRefused = 22,
}

/// <summary>The typed result of one Copy Editable command run (R6.8).</summary>
/// <param name="Succeeded">Whether the verified copy completed and every temporary shape was removed.</param>
/// <param name="StagedShapeCount">
/// The staged child count on success (the shapes a paste ungroups into), or zero on a refusal.
/// Count-only: never schedule content.
/// </param>
/// <param name="DroppedRowCount">
/// How many body rows carried a blocking validation error and were left out of the partial
/// export (owner ruling 2026-11-19), or zero when nothing was dropped.
/// </param>
/// <param name="Refusal">The typed refusal, or <see langword="null"/> on success.</param>
/// <param name="Message">
/// The one-line user feedback: the D1 success sentence with the count (and the dropped-row
/// clause when rows were omitted), or one actionable refusal line.
/// </param>
internal sealed record CopyEditableOutcome(
    bool Succeeded,
    int StagedShapeCount,
    int DroppedRowCount,
    CopyEditableRefusal? Refusal,
    string Message);

/// <summary>The inputs one Copy Editable run consumes.</summary>
/// <remarks>
/// <b>Why the seams are delegates.</b> The read-path stages are interfaces, but the export
/// request build, the bounds resolution and the lifecycle invocation are static methods a
/// host-free test cannot fake. They ride delegate seams exactly as R6.5's copy seam does, so
/// the command test pins the pipeline order and inspects the assembled
/// <see cref="StagingLifecycleRequest"/> without a host.
/// </remarks>
/// <param name="TableReader">Reads every body row into neutral DTOs. Required.</param>
/// <param name="ConfigReader">Reads the workbook's effective settings, styles and metrics. Required.</param>
/// <param name="GridMeasurement">Measures the live data panel. Required.</param>
/// <param name="ExportRequestFactory">Builds the export scene request. Required.</param>
/// <param name="BoundsResolver">Resolves the export bounds and scene. Required.</param>
/// <param name="Lifecycle">Runs compose, group, copy and cleanup. Required.</param>
/// <param name="StateScope">
/// The application-state scope the lifecycle applies and disposes (ADR-0020, R4.2). Required.
/// </param>
/// <param name="StagingWriter">
/// The staging-scheme shape write port; the staging cleanup owns what it creates. Required.
/// </param>
/// <param name="GroupingPort">The grouping port. Required.</param>
/// <param name="CopyPort">The verified-copy port. Required.</param>
/// <param name="ClipboardReader">The clipboard format reader the copy seam verifies with. Required.</param>
/// <param name="TechnicalRecord">
/// Receives the dropped-row record and every lifecycle cleanup record, or null to discard them.
/// </param>
internal sealed record CopyEditableRequest(
    IGanttTableReader? TableReader,
    IConfigCatalogueReader? ConfigReader,
    IPanelGridMeasurementPort? GridMeasurement,
    ExportRequestBuild? ExportRequestFactory,
    ExportBoundsResolve? BoundsResolver,
    StagingLifecycleRun? Lifecycle,
    IApplicationStateScope? StateScope,
    IShapeWritePort? StagingWriter,
    IShapeGroupingPort? GroupingPort,
    IShapeCopyPort? CopyPort,
    IClipboardFormatReader? ClipboardReader,
    Action<string>? TechnicalRecord = null);

/// <summary>
/// R6.8: the Copy Editable click. Runs the whole editable-export pipeline — read, validate,
/// measure, compose, stage, group, copy, verify and clean up — and reports one concise
/// success or refusal line.
/// </summary>
/// <remarks>
/// <para>
/// <b>The command owns the pipeline order; each stage owns its own contract.</b> The order is
/// pre-flight (read → configuration → validate → measurement → export request → selection →
/// bounds) then the R6.5 lifecycle (compose → group → copy → verified formats → cleanup).
/// The configuration precedes validation because the validator needs the workbook's style
/// registry to resolve a style-key row's capabilities; validating first silently dropped a
/// Custom Activity / Custom Milestone / Delineator row's own <c>LabelPosition</c> (and its
/// colour overrides), so the copy rendered the style's default position while Refresh
/// rendered the row's — the reported live-versus-copy label defect (2026-11-19).
/// Nothing in the pre-flight creates a shape, so every early refusal leaves the workbook
/// untouched and no cleanup is owed; the lifecycle's <c>finally</c> is the inner cleanup
/// guarantee for everything after the first create. A command test pins the order through the
/// delegate seams — the work item's named risk is ordering drift.
/// </para>
/// <para>
/// <b>Partial export, surfaced (owner ruling 2026-11-19).</b> The export scene is built from
/// validated events, and a blocking-error row has no defined chart geometry, so "copy the
/// table as-is, errors included" is not expressible through the landed pipeline. The owner's
/// stated fallback governs: the copy proceeds with the valid rows exactly as the live chart's
/// partial render does, and the omission is named in the feedback and in one technical record.
/// A table whose every row is invalid refuses through the existing typed refusals.
/// </para>
/// <para>
/// <b>No repairs on copy.</b> Unlike the refresh's preflight, the command performs no identity
/// repair and no column-presentation restore: a copy must not mutate the schedule. A damaged
/// row is dropped-and-named; the remedy is Fix or Refresh.
/// </para>
/// <para>
/// <b>Success is R6.6's verified copy, never a non-throwing call.</b> The command reads the
/// lifecycle's copy result, which is truthful only because <see cref="StagedShapeCopier"/>
/// verified the measured drawing formats first. The command adds no second verification and no
/// bypass, and a run that could not reach zero residue is a failure.
/// </para>
/// <para>
/// <b>D2 — one command, one state scope.</b> The live entry opens the process-shared ADR-0055
/// protection window around the run (a protected sheet refuses staging shape creation) and
/// creates the <see cref="ExcelApplicationStateScope"/>; the lifecycle applies its members at
/// the mutation boundary and disposes it in its <c>finally</c>, and the live entry's
/// <c>using</c> covers the pre-flight refusal paths (the scope guards double dispose). Both the
/// scope dispose and the window close restore the selection with <c>Range.Select()</c>, so the
/// add-in tab is restored after both (ADR-0058), exactly as the refresh paths do.
/// </para>
/// <para>
/// <b>Presentation.</b> Success writes the D1 sentence to the status bar (non-modal: the user's
/// next action is pasting into another application); when the user has hidden the status bar
/// (ADR-0020 D4) the sentence goes through the dialog presenter instead, so the confirmation is
/// never invisible. Refusals use the dialog presenter. Either presenter is contained: a
/// throwing presenter degrades to no dialog and never escapes into Excel.
/// </para>
/// </remarks>
internal static class CopyEditableCommand
{
    /// <summary>
    /// The settings key that selects the export composition (catalogue default <c>TRUE</c>).
    /// The key and its default are owned by <see cref="GanttCatalogues"/>; this constant names
    /// the one read, and an explicit case-insensitive <c>FALSE</c> is the only chart-only
    /// request.
    /// </summary>
    internal const string ExportIncludeDataPanelSettingKey =
        EditableExportPreflight.ExportIncludeDataPanelSettingKey;

    /// <summary>The one actionable line used when a required dependency is absent.</summary>
    internal const string ConstructionMessage =
        "The copy could not start. Run Refresh chart, then copy again.";

    /// <summary>
    /// Runs the live Copy Editable command in the current Excel session: resolves the live
    /// adapters, opens the protection window, runs once, and restores the add-in tab.
    /// </summary>
    internal static void RunForExcel()
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;
        Action<string> trace = CommandBoundary.Instance.WriteTechnicalRecord;
        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application);

        // D9-5: the status bar is the non-modal channel; a user who hid it gets the dialog.
        Action<string> successPresenter = SelectSuccessPresenter(
            RibbonStateService.Instance.IsStatusBarVisible(),
            RibbonStatusBar.Warn,
            CommandErrorDialog.Show);

        // ADR-0055: staging shape creation on a protected sheet is refused, so the window is
        // scoped to the pipeline only and the tab restore runs after it closes (ADR-0058).
        using (RefreshSheetCommand.OpenProtectionWindow())
        {
            // The lifecycle disposes this scope inside its finally; `using` is the
            // analyzer-satisfying double dispose, which ExcelApplicationStateScope guards,
            // and it covers every pre-flight refusal where the lifecycle never ran.
            using ExcelApplicationStateScope scope = new(application, trace);
            CopyEditableRequest request = new(
                new ExcelGanttTableReader(application),
                new ExcelConfigCatalogueReader(application),
                new ExcelPanelGridMeasurement(application),
                EditableExportRequest.Create,
                ExportBoundsResolver.TryResolve,
                StagingLifecycle.Run,
                scope,
                new ExcelShapeWriter(
                    application,
                    guard,
                    trace,
                    ShapeOwnershipScheme.Staging,
                    reuseTemplates: true),
                new ExcelShapeGrouper(application, guard),
                new ExcelShapeCopier(application, guard),
                new WindowsClipboardFormatReader(),
                trace);

            _ = Run(request, successPresenter, CommandErrorDialog.Show);
        }

        // Both the scope's dispose and the window's close restore the selection with
        // Range.Select(), which activates Excel's Table Design tab when the selection sits
        // inside tblGanttData (ADR-0058). The add-in tab is restored after both, and the call
        // is a no-op unless the surface is Ready on the Gantt sheet.
        RibbonStateService.Instance.ActivateGanttTab();
    }

    /// <summary>Runs one Copy Editable pipeline and presents its one-line result.</summary>
    /// <param name="request">The run's inputs; see <see cref="CopyEditableRequest"/>.</param>
    /// <param name="successPresenter">Receives the success sentence.</param>
    /// <param name="failurePresenter">Receives the refusal sentence.</param>
    /// <returns>The typed outcome; the presenters are the only side effect.</returns>
    /// <remarks>
    /// Exactly one presenter runs, exactly once. The stage order is the pipeline order and the
    /// first missing dependency names its stage, so a construction error is reported rather
    /// than dereferenced.
    /// </remarks>
    internal static CopyEditableOutcome Run(
        CopyEditableRequest? request,
        Action<string> successPresenter,
        Action<string> failurePresenter)
    {
        ArgumentNullException.ThrowIfNull(successPresenter);
        ArgumentNullException.ThrowIfNull(failurePresenter);

        if (request is null)
        {
            return Refuse(CopyEditableRefusal.NullRequest, ConstructionMessage, 0, failurePresenter);
        }

        // The command's own construction members (the lifecycle half) are checked before any
        // host read, so a missing dependency is reported, never dereferenced. The read-path
        // members are the shared pre-flight's checks (R7.6 D4), so both export commands share
        // one authority and each maps its typed refusal to its own message.
        if (request.Lifecycle is not { } lifecycleRun)
        {
            return Refuse(CopyEditableRefusal.MissingLifecycle, ConstructionMessage, 0, failurePresenter);
        }

        if (request.StateScope is not { } stateScope)
        {
            return Refuse(CopyEditableRefusal.MissingStateScope, ConstructionMessage, 0, failurePresenter);
        }

        if (request.StagingWriter is not { } stagingWriter)
        {
            return Refuse(CopyEditableRefusal.MissingStagingWriter, ConstructionMessage, 0, failurePresenter);
        }

        if (request.GroupingPort is not { } groupingPort)
        {
            return Refuse(CopyEditableRefusal.MissingGroupingPort, ConstructionMessage, 0, failurePresenter);
        }

        if (request.CopyPort is not { } copyPort)
        {
            return Refuse(CopyEditableRefusal.MissingCopyPort, ConstructionMessage, 0, failurePresenter);
        }

        if (request.ClipboardReader is not { } clipboardReader)
        {
            return Refuse(CopyEditableRefusal.MissingClipboardReader, ConstructionMessage, 0, failurePresenter);
        }

        // Steps 1-6 — the shared pre-flight owns read → configuration → validate → measure →
        // export request → selection → bounds, plus the post-resolution guard (R7.6 D4). Its
        // typed refusal is translated here, so every Copy Editable line stays byte-identical
        // and never names another command (the R7.6 review's F1 finding).
        EditableExportPreflightOutcome preflight = EditableExportPreflight.Run(new(
            request.TableReader,
            request.ConfigReader,
            request.GridMeasurement,
            request.ExportRequestFactory,
            request.BoundsResolver,
            request.TechnicalRecord));
        if (!preflight.Succeeded)
        {
            return RefusePreflight(preflight, failurePresenter);
        }

        EditableExportPreparation preparation = preflight.Preparation!;

        // Step 7 — the R6.5 lifecycle owns compose, group, copy, verified formats and cleanup.
        // The copy seam is the R6.6 verified stage, wired with the supplied ports and never
        // re-implemented here.
        StagingLifecycleRequest lifecycleRequest = new(
            stateScope,
            stagingWriter,
            groupingPort,
            preparation.Resolution,
            group => StagedShapeCopier.Copy(group, copyPort, clipboardReader),
            request.TechnicalRecord);
        StagingLifecycleOutcome lifecycle = lifecycleRun(lifecycleRequest);
        if (!lifecycle.Succeeded)
        {
            string refusalMessage = TranslateLifecycle(lifecycle);

            // The refusal reaches the user through a TaskDialog, whose text cannot be copied
            // out of, so the detail — the observed clipboard formats included — is recorded as
            // evidence. Without this a live copy failure was only describable by paraphrase.
            request.TechnicalRecord?.Invoke(
                "gantt.copy refused=" + (lifecycle.Refusal?.ToString() ?? "none")
                + " detail=" + refusalMessage);

            return Refuse(
                MapLifecycleRefusal(lifecycle.Refusal),
                refusalMessage,
                preflight.DroppedRowCount,
                failurePresenter);
        }

        // A clean lifecycle with no copy result is an empty composition: the seam never ran,
        // so nothing reached the clipboard. That is never reported as a zero-shape success.
        if (lifecycle.Copy is null)
        {
            return Refuse(
                CopyEditableRefusal.NothingCopied,
                "Nothing could be composed for the copy. Run Refresh chart, then copy again.",
                preflight.DroppedRowCount,
                failurePresenter);
        }

        int stagedCount = lifecycle.Composition?.StagedPrimitiveIds.Count ?? 0;
        string successMessage = SuccessMessage(stagedCount, preflight.DroppedRowCount);
        CommandBoundary.SafePresent(successPresenter, successMessage);
        return new CopyEditableOutcome(true, stagedCount, preflight.DroppedRowCount, null, successMessage);
    }

    /// <summary>Chooses the success channel from the measured status-bar visibility (D9-5).</summary>
    /// <param name="statusBarVisible">
    /// The measured visibility, or <see langword="null"/> when not determinable.
    /// </param>
    /// <param name="statusBarWriter">The non-modal writer used when the status bar is visible.</param>
    /// <param name="dialogPresenter">The dialog presenter used when the status bar is hidden.</param>
    /// <returns>The presenter to use; unknown visibility keeps the non-modal channel.</returns>
    /// <remarks>
    /// ADR-0020 D4 makes status-bar visibility the user's setting, so the D1 confirmation must
    /// not depend on a channel the user has hidden. The fail-open direction is the established
    /// non-modal one.
    /// </remarks>
    internal static Action<string> SelectSuccessPresenter(
        bool? statusBarVisible,
        Action<string> statusBarWriter,
        Action<string> dialogPresenter)
    {
        ArgumentNullException.ThrowIfNull(statusBarWriter);
        ArgumentNullException.ThrowIfNull(dialogPresenter);
        return (statusBarVisible ?? true) ? statusBarWriter : dialogPresenter;
    }

    /// <summary>Builds the D1 success sentence, including the dropped-row clause when needed.</summary>
    /// <param name="stagedShapeCount">The staged child count.</param>
    /// <param name="droppedRows">The blocking-error row count left out of the copy.</param>
    /// <returns>The user-facing sentence.</returns>
    internal static string SuccessMessage(int stagedShapeCount, int droppedRows)
    {
        string message = string.Create(
            CultureInfo.InvariantCulture,
            $"Copied the editable Gantt ({stagedShapeCount} {(stagedShapeCount == 1 ? "shape" : "shapes")}) to the clipboard.");
        return droppedRows == 0
            ? message
            : message + " " + string.Create(
                CultureInfo.InvariantCulture,
                $"{droppedRows} row(s) with errors were not copied — fix them and copy again.");
    }

    /// <summary>Translates a table-read refusal into one actionable line.</summary>
    /// <param name="refusal">The reader's typed reason.</param>
    /// <returns>The user-facing sentence.</returns>
    /// <remarks>
    /// The trailing arm covers undefined enum values, which a cast can always produce; a member
    /// added to <see cref="GanttTableReadRefusalReason"/> would land there silently, so a test
    /// asserts that every defined member has its own sentence and that Send to PowerPoint words
    /// it identically apart from its own voice.
    /// </remarks>
    internal static string TranslateRead(GanttTableReadRefusalReason refusal) => refusal switch
    {
        GanttTableReadRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open a workbook, then copy again.",
        GanttTableReadRefusalReason.TableMissing =>
            "Copy Editable could not find a table named tblGanttData on the active sheet. Run Initialise sheet first.",
        GanttTableReadRefusalReason.TableUnreadable =>
            "Copy Editable found tblGanttData but its columns no longer match the Gantt Creator schema. Run Fix to repair the table, then copy again.",
        GanttTableReadRefusalReason.DateSystemUnsupported =>
            "This workbook uses the 1904 date system, which Gantt Creator does not support. Use a 1900-date workbook and copy again.",
        _ => ConstructionMessage,
    };

    /// <summary>Translates a configuration-read refusal into one actionable line.</summary>
    /// <param name="refusal">The reader's typed reason.</param>
    /// <returns>The user-facing sentence.</returns>
    /// <remarks>
    /// The trailing arm covers undefined enum values only; a member added to
    /// <see cref="ConfigReadRefusalReason"/> would land there silently, so a test asserts that
    /// every defined member has its own sentence and that Send to PowerPoint words it identically
    /// apart from its own voice.
    /// </remarks>
    internal static string TranslateConfig(ConfigReadRefusalReason refusal) => refusal switch
    {
        ConfigReadRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open a workbook, then copy again.",
        ConfigReadRefusalReason.ConfigSheetMissing =>
            "The workbook has no Gantt Creator configuration. Run Initialise sheet first.",
        ConfigReadRefusalReason.TableMissing =>
            "The Gantt Creator configuration tables are missing. Run Initialise sheet, then copy again.",
        ConfigReadRefusalReason.HeaderMismatch or ConfigReadRefusalReason.RowCountMismatch =>
            "The Gantt Creator configuration is incomplete. Run Fix to repair it, then copy again.",
        ConfigReadRefusalReason.CatalogueMismatch =>
            "The Gantt Creator configuration does not match this add-in. Run Fix to repair it, then copy again.",
        ConfigReadRefusalReason.CatalogueHashMismatch =>
            "The Gantt Creator configuration is out of date. Run Fix to repair it, then copy again.",
        ConfigReadRefusalReason.ValueOutOfRange =>
            "The Gantt Creator configuration contains an invalid value. Run Fix to repair it, then copy again.",
        ConfigReadRefusalReason.BadColourFormat =>
            "A named-style colour in the configuration is invalid. Run Fix to repair it, then copy again.",
        _ => "The workbook configuration could not be read. Run Fix to repair it, then copy again.",
    };

    /// <summary>Translates a panel-grid refusal into one actionable line.</summary>
    /// <param name="refusal">The measurement port's typed reason.</param>
    /// <returns>The user-facing sentence.</returns>
    /// <remarks>
    /// The trailing arm covers undefined enum values only; a member added to
    /// <see cref="PanelGridRefusalReason"/> would land there silently, so a test asserts that
    /// every defined member has its own sentence and that Send to PowerPoint words it identically
    /// apart from its own voice.
    /// </remarks>
    internal static string TranslateMeasurement(PanelGridRefusalReason refusal) => refusal switch
    {
        PanelGridRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open a workbook, then copy again.",
        PanelGridRefusalReason.TableMissing =>
            "Copy Editable could not find tblGanttData on the active sheet. Run Initialise sheet first.",
        PanelGridRefusalReason.TargetProtected =>
            "The worksheet is protected, so the columns could not be measured. Unlock the sheet or run Refresh chart, then copy again.",
        PanelGridRefusalReason.InvalidMeasurement =>
            "The worksheet columns could not be measured. Run Refresh chart to re-measure them, then copy again.",
        _ => ConstructionMessage,
    };

    /// <summary>Translates an export-bounds refusal into one actionable line.</summary>
    /// <param name="resolution">The resolution carrying the bounds and scene refusals.</param>
    /// <returns>The user-facing sentence.</returns>
    private static string TranslateBounds(ExportBoundsOutcome resolution)
    {
        return resolution.Refusal switch
        {
            // Core owns this sentence; restating it here would be a second authority.
            ExportBoundsRefusal.EmptyVisibleTable => ExportBoundsRefusalMessages.EmptyVisibleTable,
            ExportBoundsRefusal.PanelThemeMissing =>
                "The copy could not be composed: the table styles were unavailable.",
            ExportBoundsRefusal.InvalidSelection =>
                "The copy could not start because the export selection is not recognised.",
            ExportBoundsRefusal.InvalidBounds =>
                "The chart has no area to copy. Adjust the chart dates, then copy again.",
            ExportBoundsRefusal.UnanchoredRequest =>
                "The copy could not be composed from the chart's rows.",
            ExportBoundsRefusal.SceneBuildRefused => resolution.SceneRefusal switch
            {
                SceneBuilderRefusal.EmptyEvents =>
                    "The table has no events to copy. Add rows or copy the chart only.",
                SceneBuilderRefusal.InvalidPlotRange =>
                    "The chart's date range is invalid. Adjust the chart dates, then copy again.",
                SceneBuilderRefusal.InvalidLayoutSettings =>
                    "The chart layout settings could not be resolved. Run Fix, then copy again.",
                SceneBuilderRefusal.ChartAboveSheet =>
                    "The chart would be drawn above the top of the sheet. Move the table down, then copy again.",
                SceneBuilderRefusal.UnresolvableStyle =>
                    "A row's style could not be resolved. Run Fix, then copy again.",
                SceneBuilderRefusal.UnresolvableProjection =>
                    "The parent/child structure could not be resolved. Fix the ParentId cells, then copy again.",
                SceneBuilderRefusal.NullRequest
                    or SceneBuilderRefusal.NullMetrics
                    or SceneBuilderRefusal.NullPanelGrid
                    or SceneBuilderRefusal.NullBounds
                    or SceneBuilderRefusal.LiveProfileCarriesPanel
                    or SceneBuilderRefusal.UnresolvableLaneAnchor => ConstructionMessage,
                _ => "The chart could not be composed. Fix the table errors marked on the rows, then copy again.",
            },
            ExportBoundsRefusal.NullRequest => ConstructionMessage,
            _ => ConstructionMessage,
        };
    }

    /// <summary>Translates a lifecycle refusal into one actionable line.</summary>
    /// <param name="outcome">The lifecycle outcome carrying the stage details.</param>
    /// <returns>The user-facing sentence.</returns>
    private static string TranslateLifecycle(StagingLifecycleOutcome outcome) => outcome.Refusal switch
    {
        StagingLifecycleRefusal.CompositionRefused => outcome.Composition?.Refusal switch
        {
            // Staging shapes are invisible to the live ownership scan, so a Refresh cannot
            // remove one. The only remedy that exists is deleting the named shapes.
            StagingCompositionRefusal.StagingAlreadyPopulated =>
                "Temporary copy shapes from an earlier attempt are still on the sheet, named GanttCreatorStaging. Select and delete them, then copy again.",
            StagingCompositionRefusal.HostRejected =>
                TranslateWriteRefusal(outcome.Composition?.WriteRefusal),
            // The resolution carries the real export refusal; restating it generically was
            // the same defect one layer up.
            StagingCompositionRefusal.ExportRefused => outcome.Composition?.Resolution is { } refusedResolution
                ? TranslateBounds(refusedResolution)
                : "The copy could not be composed from the export bounds. Run Fix, then copy again.",
            StagingCompositionRefusal.UnsupportedPrimitive =>
                "A chart element could not be copied. Run Fix, then copy again.",
            StagingCompositionRefusal.InvalidSubScene =>
                "The copy could not be composed. Run Fix, then copy again.",
            StagingCompositionRefusal.NullResolution
                or StagingCompositionRefusal.NullPort
                or StagingCompositionRefusal.MissingScene => ConstructionMessage,
            _ => "The copy could not be composed. Run Fix, then copy again.",
        },
        StagingLifecycleRefusal.GroupingRefused =>
            "The copy could not be grouped. Try again; if the problem persists, run Fix.",
        // The R6.6 stage's detail already carries the actionable sentence (and the reader's
        // diagnostic token); the typed member is never parsed out of it.
        StagingLifecycleRefusal.CopyRefused =>
            outcome.Copy?.Detail ?? "The staged copy could not be placed on the clipboard.",
        StagingLifecycleRefusal.CleanupResidue =>
            "The copy was made but temporary shapes could not be removed. Run Refresh chart to clean them up.",
        // R7.6 adds the transfer seam; Copy Editable supplies no hook, so this arm is
        // unreachable here. It is enumerated so a future refusal is a compile-time reminder.
        StagingLifecycleRefusal.TransferRefused => ConstructionMessage,
        StagingLifecycleRefusal.NullRequest
            or StagingLifecycleRefusal.NullStateScope
            or StagingLifecycleRefusal.NullStagingPort
            or StagingLifecycleRefusal.NullGroupingPort
            or StagingLifecycleRefusal.NullResolution => ConstructionMessage,
        _ => ConstructionMessage,
    };

    /// <summary>Translates the write refusal behind a refused composition.</summary>
    /// <param name="refusal">The port's typed refusal, when the composer carried one.</param>
    /// <returns>One sentence naming the reason, never a remedy that cannot work.</returns>
    /// <remarks>
    /// <para>
    /// The 2026-11-19 defect: every write refusal was reported as "Excel refused to create
    /// the copy. Run Refresh chart, then copy again." A Refresh cannot remove a staging
    /// shape (the live ownership scan never proves one) and cannot make a zero-area
    /// primitive creatable, so the advice sent the user in a circle. The composer already
    /// keeps the typed member on the outcome; this maps it.
    /// </para>
    /// <para>
    /// Every member is enumerated, so a refusal added later is a compile-time reminder
    /// rather than a silent fall-through to a generic sentence.
    /// </para>
    /// </remarks>
    private static string TranslateWriteRefusal(ShapeWriteRefusal? refusal) => refusal switch
    {
        ShapeWriteRefusal.TargetProtected =>
            "The worksheet is protected, so Excel refused the copy. Unprotect it, then copy again.",
        ShapeWriteRefusal.AlreadyExists =>
            "A temporary copy shape with the same name is already on the sheet. Select and delete the shapes named GanttCreatorStaging, then copy again.",
        ShapeWriteRefusal.InvalidGeometry =>
            "Part of the copy has no visible area, so Excel cannot create it. Run Fix, then copy again.",
        ShapeWriteRefusal.UnsupportedIdentifierForm or ShapeWriteRefusal.BlankIdentifier =>
            "The copy could not be composed from the table. Run Fix, then copy again.",
        ShapeWriteRefusal.NoActiveWorkbook or ShapeWriteRefusal.WorksheetMissing =>
            "The workbook with the Gantt table is not available. Open it, then copy again.",
        ShapeWriteRefusal.HostRejected
            or ShapeWriteRefusal.NotFound
            or ShapeWriteRefusal.InvalidOwnershipTag
            or ShapeWriteRefusal.NullRequest => "Excel refused to create the copy.",
        null => "Excel refused to create the copy.",
        _ => "Excel refused to create the copy.",
    };

    /// <summary>Maps the stored inclusion setting to the export selection.</summary>
    /// <param name="settings">The workbook's effective settings map.</param>
    /// <returns>Chart-only on an explicit case-insensitive <c>FALSE</c>; otherwise table-plus-chart.</returns>
    /// <remarks>
    /// R7.6 moved the rule into the shared pre-flight so both export commands read one
    /// authority; this forwarding member keeps the landed R6.8 reasoning and its integration
    /// harness source-compatible without a second implementation.
    /// </remarks>
    internal static ExportSelection SelectExportSelection(IReadOnlyDictionary<string, string> settings) =>
        EditableExportPreflight.SelectExportSelection(settings);

    /// <summary>Maps a lifecycle refusal to the command's typed refusal.</summary>
    /// <param name="refusal">The lifecycle's reason, or null when it succeeded.</param>
    /// <returns>The command-level refusal.</returns>
    private static CopyEditableRefusal MapLifecycleRefusal(StagingLifecycleRefusal? refusal) => refusal switch
    {
        StagingLifecycleRefusal.CompositionRefused => CopyEditableRefusal.CompositionRefused,
        StagingLifecycleRefusal.GroupingRefused => CopyEditableRefusal.GroupingRefused,
        StagingLifecycleRefusal.CopyRefused => CopyEditableRefusal.CopyRefused,
        StagingLifecycleRefusal.CleanupResidue => CopyEditableRefusal.CleanupResidue,
        // R7.6: unreachable for Copy Editable (no transfer hook), mapped to the generic arm.
        StagingLifecycleRefusal.TransferRefused => CopyEditableRefusal.LifecycleRefused,
        StagingLifecycleRefusal.NullRequest
            or StagingLifecycleRefusal.NullStateScope
            or StagingLifecycleRefusal.NullStagingPort
            or StagingLifecycleRefusal.NullGroupingPort
            or StagingLifecycleRefusal.NullResolution => CopyEditableRefusal.LifecycleRefused,
        _ => CopyEditableRefusal.LifecycleRefused,
    };

    /// <summary>Maps the shared pre-flight's typed refusal to one Copy Editable line.</summary>
    /// <param name="outcome">The pre-flight outcome carrying the underlying typed refusal.</param>
    /// <param name="presenter">The failure presenter.</param>
    /// <returns>The refused command outcome.</returns>
    /// <remarks>
    /// The pre-flight owns the chain, not the message: this translation is the command's own
    /// authority, so Copy Editable's established strings (including the two that name the
    /// command) are preserved exactly and can never be presented by another command.
    /// </remarks>
    private static CopyEditableOutcome RefusePreflight(
        EditableExportPreflightOutcome outcome,
        Action<string> presenter)
    {
        (CopyEditableRefusal refusal, string message) = outcome.Refusal switch
        {
            EditableExportPreflightRefusal.NullRequest =>
                (CopyEditableRefusal.NullRequest, ConstructionMessage),
            EditableExportPreflightRefusal.MissingTableReader =>
                (CopyEditableRefusal.MissingTableReader, ConstructionMessage),
            EditableExportPreflightRefusal.MissingConfigReader =>
                (CopyEditableRefusal.MissingConfigReader, ConstructionMessage),
            EditableExportPreflightRefusal.MissingGridMeasurement =>
                (CopyEditableRefusal.MissingGridMeasurement, ConstructionMessage),
            EditableExportPreflightRefusal.MissingExportRequestFactory =>
                (CopyEditableRefusal.MissingExportRequestFactory, ConstructionMessage),
            EditableExportPreflightRefusal.MissingBoundsResolver =>
                (CopyEditableRefusal.MissingBoundsResolver, ConstructionMessage),
            EditableExportPreflightRefusal.TableReadRefused =>
                (CopyEditableRefusal.TableReadRefused, outcome.TableReadRefusal is { } readRefusal
                    ? TranslateRead(readRefusal)
                    : ConstructionMessage),
            EditableExportPreflightRefusal.ConfigReadRefused =>
                (CopyEditableRefusal.ConfigReadRefused, outcome.ConfigRefusal is { } configRefusal
                    ? TranslateConfig(configRefusal)
                    : ConstructionMessage),
            EditableExportPreflightRefusal.MeasurementRefused =>
                (CopyEditableRefusal.MeasurementRefused, outcome.MeasurementRefusal is { } gridRefusal
                    ? TranslateMeasurement(gridRefusal)
                    : ConstructionMessage),
            EditableExportPreflightRefusal.ExportRequestRefused =>
                (CopyEditableRefusal.ExportRequestRefused, outcome.ExportRequestMessage ?? ConstructionMessage),
            EditableExportPreflightRefusal.BoundsRefused =>
                (CopyEditableRefusal.BoundsRefused, outcome.BoundsOutcome is { } bounds
                    ? TranslateBounds(bounds)
                    : ConstructionMessage),
            _ => (CopyEditableRefusal.LifecycleRefused, ConstructionMessage),
        };

        return Refuse(refusal, message, outcome.DroppedRowCount, presenter);
    }

    /// <summary>Presents a refusal and returns its typed outcome.</summary>
    /// <param name="refusal">The command-level refusal.</param>
    /// <param name="message">The one actionable line.</param>
    /// <param name="droppedRows">The blocking-error row count, when it is known.</param>
    /// <param name="presenter">The failure presenter.</param>
    /// <returns>The refused outcome.</returns>
    private static CopyEditableOutcome Refuse(
        CopyEditableRefusal refusal,
        string message,
        int droppedRows,
        Action<string> presenter)
    {
        CommandBoundary.SafePresent(presenter, message);
        return new CopyEditableOutcome(false, 0, droppedRows, refusal, message);
    }

}
