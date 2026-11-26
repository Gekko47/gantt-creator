using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>Why the Send to PowerPoint command produced no verified transfer (R7.6).</summary>
/// <remarks>
/// Every member is a typed, expected outcome rather than an exception. The construction
/// members exist so a dependency the command cannot run without is reported, never
/// dereferenced; the stage members name the first refusal in the pipeline order.
/// </remarks>
internal enum SendToPowerPointRefusal
{
    /// <summary>No command request was supplied.</summary>
    NullRequest = 0,

    /// <summary>No staging lifecycle was supplied.</summary>
    MissingLifecycle = 1,

    /// <summary>No application-state scope was supplied.</summary>
    MissingStateScope = 2,

    /// <summary>No staging shape write port was supplied.</summary>
    MissingStagingWriter = 3,

    /// <summary>No grouping port was supplied.</summary>
    MissingGroupingPort = 4,

    /// <summary>No shape copy port was supplied.</summary>
    MissingCopyPort = 5,

    /// <summary>No clipboard format reader was supplied.</summary>
    MissingClipboardReader = 6,

    /// <summary>No PowerPoint application port was supplied.</summary>
    MissingPowerPointApplication = 7,

    /// <summary>No PowerPoint paste port was supplied.</summary>
    MissingPowerPointPaste = 8,

    /// <summary>No PowerPoint placement port was supplied.</summary>
    MissingPowerPointPlacement = 9,

    /// <summary>No table reader was supplied.</summary>
    MissingTableReader = 10,

    /// <summary>No configuration reader was supplied.</summary>
    MissingConfigReader = 11,

    /// <summary>No panel-grid measurement port was supplied.</summary>
    MissingGridMeasurement = 12,

    /// <summary>No export-request factory was supplied.</summary>
    MissingExportRequestFactory = 13,

    /// <summary>No bounds resolver was supplied.</summary>
    MissingBoundsResolver = 14,

    /// <summary>The table could not be read, so nothing was sent.</summary>
    TableReadRefused = 15,

    /// <summary>The workbook configuration could not be read, so nothing was sent.</summary>
    ConfigReadRefused = 16,

    /// <summary>The live panel grid could not be measured, so nothing was sent.</summary>
    MeasurementRefused = 17,

    /// <summary>The export scene request was refused, so nothing was sent.</summary>
    ExportRequestRefused = 18,

    /// <summary>The export bounds or scene were refused, including the post-resolution guard.</summary>
    BoundsRefused = 19,

    /// <summary>The staging composition refused; its shapes were cleaned up.</summary>
    CompositionRefused = 20,

    /// <summary>The staged set could not be grouped; the children were restored.</summary>
    GroupingRefused = 21,

    /// <summary>The verified copy refused (see the copy result's own detail).</summary>
    CopyRefused = 22,

    /// <summary>The transfer refused; R7.5's typed failure carries the one actionable message.</summary>
    TransferRefused = 23,

    /// <summary>Cleanup reported residue, which is a failure and never a success.</summary>
    CleanupResidue = 24,

    /// <summary>The lifecycle ran cleanly but the transfer seam never ran.</summary>
    NothingTransferred = 25,

    /// <summary>The lifecycle refused for a construction reason after the command's checks.</summary>
    LifecycleRefused = 26,
}

/// <summary>The typed result of one Send to PowerPoint command run (R7.6).</summary>
/// <param name="Succeeded">Whether the verified group was pasted, validated and placed once.</param>
/// <param name="StagedShapeCount">The validated pasted child count on success, or zero on a refusal.</param>
/// <param name="DroppedRowCount">
/// How many body rows carried a blocking validation error and were left out of the partial
/// export, or zero when nothing was dropped.
/// </param>
/// <param name="PresentationCreated">
/// Whether the transfer created a presentation because none was open (D-R7.2). The creation is
/// announced in the success message; on a later refusal the technical record names it, because
/// the closed failure message contract cannot be extended (R7.5).
/// </param>
/// <param name="Refusal">The typed refusal, or <see langword="null"/> on success.</param>
/// <param name="Message">The one-line user feedback: the count-only success sentence or one actionable refusal line.</param>
internal sealed record SendToPowerPointOutcome(
    bool Succeeded,
    int StagedShapeCount,
    int DroppedRowCount,
    bool PresentationCreated,
    SendToPowerPointRefusal? Refusal,
    string Message);

/// <summary>The inputs one Send to PowerPoint run consumes (R7.6).</summary>
/// <param name="TableReader">Reads every body row into neutral DTOs. Required.</param>
/// <param name="ConfigReader">Reads the workbook's effective settings, styles and metrics. Required.</param>
/// <param name="GridMeasurement">Measures the live data panel. Required.</param>
/// <param name="ExportRequestFactory">Builds the export scene request. Required.</param>
/// <param name="BoundsResolver">Resolves the export bounds and scene. Required.</param>
/// <param name="Lifecycle">Runs compose, group, copy, transfer and cleanup. Required.</param>
/// <param name="StateScope">The application-state scope the lifecycle applies and disposes. Required.</param>
/// <param name="StagingWriter">The staging-scheme shape write port. Required.</param>
/// <param name="GroupingPort">The grouping port. Required.</param>
/// <param name="CopyPort">The verified-copy port. Required.</param>
/// <param name="ClipboardReader">The clipboard format reader the copy seam verifies with. Required.</param>
/// <param name="PowerPointApplication">The typed PowerPoint application port the target resolver drives. Required.</param>
/// <param name="PowerPointPaste">The typed PowerPoint paste port. Required.</param>
/// <param name="PowerPointPlacement">The typed PowerPoint placement port. Required.</param>
/// <param name="TechnicalRecord">Receives the dropped-row, failure and creation records, or null to discard them.</param>
internal sealed record SendToPowerPointRequest(
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
    IPowerPointApplicationPort? PowerPointApplication,
    IPowerPointPastePort? PowerPointPaste,
    IPowerPointPlacementPort? PowerPointPlacement,
    Action<string>? TechnicalRecord = null);

/// <summary>
/// R7.6: the Send to PowerPoint click. Runs the shared export pre-flight, the R6.5 staging
/// lifecycle with the R7.6 transfer seam, and reports one concise success or refusal line.
/// </summary>
/// <remarks>
/// <para>
/// <b>One command, the full pipeline (D1).</b> The pre-flight owns read → configuration →
/// validate → measure → export request → selection → bounds (and the post-resolution guard);
/// the lifecycle owns compose → group → verified copy → transfer → cleanup. The transfer runs
/// inside the lifecycle, after the verified copy and before the <c>finally</c>, so the
/// clipboard is live while the staged shapes exist and the <c>finally</c> remains the one
/// Excel cleanup authority.
/// </para>
/// <para>
/// <b>Ownership is structural (D2).</b> The PowerPoint ports expose no save, close or quit
/// member; the command holds ports, never proxies, and the live entry constructs exactly one
/// <see cref="PowerPointTransferAdapter"/> over one <see cref="PowerPointProxyOwnership"/> and
/// disposes it in a <c>using</c>. A transfer never closes or terminates a pre-existing host.
/// </para>
/// <para>
/// <b>A failed transfer never costs user data (D3).</b> The lifecycle's <c>finally</c> removes
/// the Excel staging; the transfer hook removes a pasted group when a later stage failed, and
/// the removal claim is presented only when the delete confirmed it (the mapper's
/// <c>MapResidueCleanup</c> substitutes a truthful host failure otherwise). The application
/// state scope is disposed even on a pre-flight refusal.
/// </para>
/// <para>
/// <b>Messages are this command's own.</b> The shared pre-flight carries typed refusals, not
/// text; Copy Editable's established strings are never presented here (R7.6 review F1). The
/// transfer failures present exactly R7.5's mapper message.
/// </para>
/// </remarks>
internal static class SendToPowerPointCommand
{
    /// <summary>The one actionable line used when a required dependency is absent.</summary>
    internal const string ConstructionMessage =
        "The send could not start. Run Refresh chart, then send again.";

    /// <summary>
    /// Runs the live Send to PowerPoint command in the current Excel session: resolves the
    /// live adapters, opens the protection window, runs once, and restores the add-in tab.
    /// </summary>
    /// <remarks>
    /// The single <see cref="PowerPointTransferAdapter"/> is constructed here and disposed
    /// before the state scope and the protection window close. Its dispose releases the
    /// application proxy it acquired and never terminates a pre-existing PowerPoint.
    /// </remarks>
    internal static void RunForExcel()
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;
        Action<string> trace = CommandBoundary.Instance.WriteTechnicalRecord;
        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application);

        // D9-5: the status bar is the non-modal channel; a user who hid it gets the dialog.
        Action<string> successPresenter = CopyEditableCommand.SelectSuccessPresenter(
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

            // R7.6 D2: one adapter implements all three PowerPoint ports and owns every proxy
            // it acquires; disposing it releases them exactly once in reverse order.
            using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership(), trace);

            SendToPowerPointRequest request = new(
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
                adapter,
                adapter,
                adapter,
                trace);

            _ = Run(request, successPresenter, CommandErrorDialog.Show);
        }

        RibbonStateService.Instance.ActivateGanttTab();
    }

    /// <summary>Runs one Send to PowerPoint pipeline and presents its one-line result.</summary>
    /// <param name="request">The run's inputs; see <see cref="SendToPowerPointRequest"/>.</param>
    /// <param name="successPresenter">Receives the success sentence.</param>
    /// <param name="failurePresenter">Receives the refusal sentence.</param>
    /// <returns>The typed outcome; the presenters are the only side effect.</returns>
    internal static SendToPowerPointOutcome Run(
        SendToPowerPointRequest? request,
        Action<string> successPresenter,
        Action<string> failurePresenter)
    {
        ArgumentNullException.ThrowIfNull(successPresenter);
        ArgumentNullException.ThrowIfNull(failurePresenter);

        if (request is null)
        {
            return Refuse(SendToPowerPointRefusal.NullRequest, ConstructionMessage, 0, false, failurePresenter);
        }

        // The command's own construction members are checked before any host read, so a
        // missing dependency is reported, never dereferenced. The read-path members are the
        // shared pre-flight's checks (R7.6 D4).
        if (request.Lifecycle is not { } lifecycleRun)
        {
            return Refuse(SendToPowerPointRefusal.MissingLifecycle, ConstructionMessage, 0, false, failurePresenter);
        }

        if (request.StateScope is not { } stateScope)
        {
            return Refuse(SendToPowerPointRefusal.MissingStateScope, ConstructionMessage, 0, false, failurePresenter);
        }

        if (request.StagingWriter is not { } stagingWriter)
        {
            return Refuse(SendToPowerPointRefusal.MissingStagingWriter, ConstructionMessage, 0, false, failurePresenter);
        }

        if (request.GroupingPort is not { } groupingPort)
        {
            return Refuse(SendToPowerPointRefusal.MissingGroupingPort, ConstructionMessage, 0, false, failurePresenter);
        }

        if (request.CopyPort is not { } copyPort)
        {
            return Refuse(SendToPowerPointRefusal.MissingCopyPort, ConstructionMessage, 0, false, failurePresenter);
        }

        if (request.ClipboardReader is not { } clipboardReader)
        {
            return Refuse(SendToPowerPointRefusal.MissingClipboardReader, ConstructionMessage, 0, false, failurePresenter);
        }

        if (request.PowerPointApplication is not { } powerPointApplication)
        {
            return Refuse(SendToPowerPointRefusal.MissingPowerPointApplication, ConstructionMessage, 0, false, failurePresenter);
        }

        if (request.PowerPointPaste is not { } powerPointPaste)
        {
            return Refuse(SendToPowerPointRefusal.MissingPowerPointPaste, ConstructionMessage, 0, false, failurePresenter);
        }

        if (request.PowerPointPlacement is not { } powerPointPlacement)
        {
            return Refuse(SendToPowerPointRefusal.MissingPowerPointPlacement, ConstructionMessage, 0, false, failurePresenter);
        }

        // The shared pre-flight owns the export chain and the post-resolution guard.
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
        bool presentationCreated = false;

        // Step 7 — the lifecycle owns compose, group, verified copy, the transfer seam and
        // cleanup. The transfer closes over the preparation's bounds and inset (D5: the same
        // EditableExport scene the copy composes).
        StagingLifecycleRequest lifecycleRequest = new(
            stateScope,
            stagingWriter,
            groupingPort,
            preparation.Resolution,
            group => StagedShapeCopier.Copy(group, copyPort, clipboardReader),
            request.TechnicalRecord,
            (group, _) => TransferStagedGroup(
                group,
                preparation,
                powerPointApplication,
                powerPointPaste,
                powerPointPlacement,
                request.TechnicalRecord,
                () => presentationCreated = true));

        StagingLifecycleOutcome lifecycle = lifecycleRun(lifecycleRequest);
        if (!lifecycle.Succeeded)
        {
            (SendToPowerPointRefusal refusal, string message) = TranslateLifecycle(lifecycle);

            // One command-level record per non-transfer refusal; a transfer refusal already
            // carries the mapper's richer single record from the hook (checklist K).
            if (lifecycle.Refusal != StagingLifecycleRefusal.TransferRefused)
            {
                request.TechnicalRecord?.Invoke(
                    "gantt.send refused=" + (lifecycle.Refusal?.ToString() ?? "none")
                    + " detail=" + message);
            }

            // The closed failure message cannot name a created deck (R7.5), so the record does.
            if (presentationCreated)
            {
                request.TechnicalRecord?.Invoke(
                    "gantt.send presentation-created=true outcome=refused");
            }

            return Refuse(refusal, message, preflight.DroppedRowCount, presentationCreated, failurePresenter);
        }

        if (lifecycle.Transfer is not { } transfer)
        {
            return Refuse(
                SendToPowerPointRefusal.NothingTransferred,
                "Nothing was sent. Run Refresh chart, then send again.",
                preflight.DroppedRowCount,
                presentationCreated,
                failurePresenter);
        }

        string successMessage = SuccessMessage(
            transfer.StagedShapeCount,
            preflight.DroppedRowCount,
            presentationCreated);
        CommandBoundary.SafePresent(successPresenter, successMessage);
        return new SendToPowerPointOutcome(
            true,
            transfer.StagedShapeCount,
            preflight.DroppedRowCount,
            presentationCreated,
            null,
            successMessage);
    }

    /// <summary>
    /// The transfer chain: target decision → announced creation → paste with the verified
    /// expectation → one placement → truthful residue cleanup on any post-paste failure.
    /// </summary>
    /// <param name="group">The verified staged group the copy seam reported.</param>
    /// <param name="preparation">The shared pre-flight preparation (bounds and inset).</param>
    /// <param name="application">The PowerPoint application port.</param>
    /// <param name="pastePort">The PowerPoint paste port.</param>
    /// <param name="placementPort">The PowerPoint placement port.</param>
    /// <param name="technicalRecord">Receives one record per failure.</param>
    /// <param name="onPresentationCreated">Invoked once when a presentation was created.</param>
    /// <returns>The typed transfer result.</returns>
    private static StagingTransferResult TransferStagedGroup(
        StagedGroupOutcome group,
        EditableExportPreparation preparation,
        IPowerPointApplicationPort application,
        IPowerPointPastePort pastePort,
        IPowerPointPlacementPort placementPort,
        Action<string>? technicalRecord,
        Action onPresentationCreated)
    {
        // 1. The landed R7.2 decision table owns proceed/create/refuse.
        PowerPointTargetDecision decision = new PowerPointTargetResolver(application).Resolve();
        PowerPointTransferFailure? failure = PowerPointTransferFailureMapper.MapTargetDecision(decision);
        if (failure is not null)
        {
            return Refused(failure, technicalRecord);
        }

        // 2. Create only when no presentation is open (D-R7.2), announced by the caller.
        if (decision.Action == PowerPointTargetAction.CreatePresentation)
        {
            PowerPointPresentationCreation creation = application.CreatePresentationWithBlankSlide();
            failure = PowerPointTransferFailureMapper.MapCreation(creation);
            if (failure is not null)
            {
                return Refused(failure, technicalRecord);
            }

            onPresentationCreated();
        }

        // 3. The expectation is assembled from the verified staged group, exactly as the R7.3
        // type's contract states: the child count and the verified envelope.
        RectD envelope = group.Bounds ?? preparation.Resolution.Bounds!.Bounds;
        PowerPointPasteExpectation expectation = new(
            group.ChildOrder.Count,
            envelope.Width,
            envelope.Height);

        PowerPointPasteStageResult paste = new PowerPointPasteStage(pastePort).Paste(expectation);
        failure = PowerPointTransferFailureMapper.MapPasteStage(paste);
        if (failure is not null)
        {
            return RefusedAfterPaste(failure, paste.GroupId, pastePort, technicalRecord);
        }

        // A validated paste always located exactly one group, so the identity is present; a
        // missing one is a construction/sequence defect, never a silent success.
        if (paste.GroupId is not { } groupId)
        {
            return Refused(
                PowerPointTransferFailureMapper.SequenceViolated("paste-group-identity"),
                technicalRecord);
        }

        // 4. The single placement: the zero-origin export bounds and the themed inset (D-R7.4).
        PowerPointPlacementResult placement = new PowerPointPlacementStage(placementPort, technicalRecord)
            .Place(groupId, preparation.Resolution.Bounds!.Bounds, preparation.InsetPt);
        failure = PowerPointTransferFailureMapper.MapPlacement(placement);
        if (failure is not null)
        {
            return RefusedAfterPaste(failure, groupId, pastePort, technicalRecord);
        }

        PowerPointPlacementObservation observation = placement.Observation!;
        RectD placed = new(
            observation.LeftPt,
            observation.TopPt,
            observation.WidthPt,
            observation.HeightPt);
        return StagingTransferResult.Success(group.ChildOrder.Count, placed);
    }

    /// <summary>Records and returns a pre-paste transfer failure.</summary>
    /// <param name="failure">The typed failure.</param>
    /// <param name="technicalRecord">The record sink.</param>
    /// <returns>The refused transfer result.</returns>
    private static StagingTransferResult Refused(
        PowerPointTransferFailure failure,
        Action<string>? technicalRecord)
    {
        technicalRecord?.Invoke(failure.TechnicalRecord);
        return StagingTransferResult.Refused(failure);
    }

    /// <summary>
    /// Records a failure that happened after a paste materialised, removes the pasted group,
    /// and substitutes a truthful host failure when the removal was refused.
    /// </summary>
    /// <param name="failure">The stage failure.</param>
    /// <param name="groupId">The pasted group's identifier, or null when none materialised.</param>
    /// <param name="pastePort">The port that performs the residue delete.</param>
    /// <param name="technicalRecord">The record sink.</param>
    /// <returns>The refused transfer result whose message never claims an unconfirmed removal.</returns>
    /// <remarks>
    /// The R7.5 <c>PasteValidationFailed</c> message claims the pasted shapes were removed. A
    /// refused delete would make that claim false, so the mapper's residue mapping presents the
    /// host failure instead and the helper's own record names the leftover (R7.6 review F2).
    /// </remarks>
    private static StagingTransferResult RefusedAfterPaste(
        PowerPointTransferFailure failure,
        int? groupId,
        IPowerPointPastePort pastePort,
        Action<string>? technicalRecord)
    {
        technicalRecord?.Invoke(failure.TechnicalRecord);

        PowerPointGroupDeleteResult? cleanup = PowerPointPasteResidueCleanup.Cleanup(
            groupId,
            pastePort,
            technicalRecord);
        PowerPointTransferFailure? residue = cleanup is null
            ? null
            : PowerPointTransferFailureMapper.MapResidueCleanup(cleanup);
        if (residue is null)
        {
            return StagingTransferResult.Refused(failure);
        }

        // The removal claim cannot be made: the substituted host failure is the presented
        // text, and its record names the cleanup refusal with the taxonomy member.
        technicalRecord?.Invoke(residue.TechnicalRecord);
        return StagingTransferResult.Refused(residue);
    }

    /// <summary>Maps the shared pre-flight's typed refusal to one Send line.</summary>
    /// <param name="outcome">The pre-flight outcome carrying the underlying typed refusal.</param>
    /// <param name="presenter">The failure presenter.</param>
    /// <returns>The refused command outcome.</returns>
    private static SendToPowerPointOutcome RefusePreflight(
        EditableExportPreflightOutcome outcome,
        Action<string> presenter)
    {
        (SendToPowerPointRefusal refusal, string message) = outcome.Refusal switch
        {
            EditableExportPreflightRefusal.NullRequest =>
                (SendToPowerPointRefusal.NullRequest, ConstructionMessage),
            EditableExportPreflightRefusal.MissingTableReader =>
                (SendToPowerPointRefusal.MissingTableReader, ConstructionMessage),
            EditableExportPreflightRefusal.MissingConfigReader =>
                (SendToPowerPointRefusal.MissingConfigReader, ConstructionMessage),
            EditableExportPreflightRefusal.MissingGridMeasurement =>
                (SendToPowerPointRefusal.MissingGridMeasurement, ConstructionMessage),
            EditableExportPreflightRefusal.MissingExportRequestFactory =>
                (SendToPowerPointRefusal.MissingExportRequestFactory, ConstructionMessage),
            EditableExportPreflightRefusal.MissingBoundsResolver =>
                (SendToPowerPointRefusal.MissingBoundsResolver, ConstructionMessage),
            EditableExportPreflightRefusal.TableReadRefused =>
                (SendToPowerPointRefusal.TableReadRefused, outcome.TableReadRefusal is { } readRefusal
                    ? TranslateRead(readRefusal)
                    : ConstructionMessage),
            EditableExportPreflightRefusal.ConfigReadRefused =>
                (SendToPowerPointRefusal.ConfigReadRefused, outcome.ConfigRefusal is { } configRefusal
                    ? TranslateConfig(configRefusal)
                    : ConstructionMessage),
            EditableExportPreflightRefusal.MeasurementRefused =>
                (SendToPowerPointRefusal.MeasurementRefused, outcome.MeasurementRefusal is { } gridRefusal
                    ? TranslateMeasurement(gridRefusal)
                    : ConstructionMessage),
            EditableExportPreflightRefusal.ExportRequestRefused =>
                (SendToPowerPointRefusal.ExportRequestRefused, outcome.ExportRequestMessage ?? ConstructionMessage),
            EditableExportPreflightRefusal.BoundsRefused =>
                (SendToPowerPointRefusal.BoundsRefused, outcome.BoundsOutcome is { } bounds
                    ? TranslateBounds(bounds)
                    : ConstructionMessage),
            _ => (SendToPowerPointRefusal.LifecycleRefused, ConstructionMessage),
        };

        return Refuse(refusal, message, outcome.DroppedRowCount, false, presenter);
    }

    /// <summary>Maps a lifecycle refusal to the command's typed refusal and one line.</summary>
    /// <param name="outcome">The lifecycle outcome carrying the stage details.</param>
    /// <returns>The command refusal and its actionable message.</returns>
    private static (SendToPowerPointRefusal Refusal, string Message) TranslateLifecycle(
        StagingLifecycleOutcome outcome) => outcome.Refusal switch
        {
            StagingLifecycleRefusal.CompositionRefused => (SendToPowerPointRefusal.CompositionRefused, TranslateComposition(outcome)),
            StagingLifecycleRefusal.GroupingRefused => (SendToPowerPointRefusal.GroupingRefused, "The chart could not be grouped for the transfer. Try again; if the problem persists, run Fix."),
            StagingLifecycleRefusal.CopyRefused => (SendToPowerPointRefusal.CopyRefused, TranslateCopy(outcome)),
            StagingLifecycleRefusal.TransferRefused => (SendToPowerPointRefusal.TransferRefused, outcome.Transfer?.Failure?.Message ?? "The transfer to PowerPoint failed. Try again."),
            StagingLifecycleRefusal.CleanupResidue => (SendToPowerPointRefusal.CleanupResidue, "The transfer ran but temporary shapes could not be removed. Run Refresh chart to clean them up."),
            StagingLifecycleRefusal.NullRequest
                or StagingLifecycleRefusal.NullStateScope
                or StagingLifecycleRefusal.NullStagingPort
                or StagingLifecycleRefusal.NullGroupingPort
                or StagingLifecycleRefusal.NullResolution => (SendToPowerPointRefusal.LifecycleRefused, ConstructionMessage),
            _ => (SendToPowerPointRefusal.LifecycleRefused, ConstructionMessage),
        };

    /// <summary>Translates a copy refusal, including the transfer taxonomy's clipboard arm.</summary>
    /// <param name="outcome">The lifecycle outcome carrying the copy result.</param>
    /// <returns>The user-facing sentence.</returns>
    /// <remarks>
    /// R7.5's <c>MapCopyResult</c> exists for this command: a busy or unavailable clipboard is
    /// a transfer failure with its own actionable line. Every other copy refusal keeps the
    /// verified copy stage's own detail, so no message has a second authority.
    /// </remarks>
    private static string TranslateCopy(StagingLifecycleOutcome outcome)
    {
        PowerPointTransferFailure? clipboard = PowerPointTransferFailureMapper.MapCopyResult(outcome.Copy);
        return clipboard?.Message
            ?? outcome.Copy?.Detail
            ?? "The staged copy could not be placed on the clipboard.";
    }

    /// <summary>Translates a refused composition into one line.</summary>
    /// <param name="outcome">The lifecycle outcome carrying the composition details.</param>
    /// <returns>The user-facing sentence.</returns>
    private static string TranslateComposition(StagingLifecycleOutcome outcome) =>
        outcome.Composition?.Refusal switch
        {
            StagingCompositionRefusal.StagingAlreadyPopulated => "Temporary copy shapes from an earlier attempt are still on the sheet, named GanttCreatorStaging. Select and delete them, then send again.",
            StagingCompositionRefusal.HostRejected => TranslateWriteRefusal(outcome.Composition?.WriteRefusal),
            StagingCompositionRefusal.ExportRefused => outcome.Composition?.Resolution is { } refused
                ? TranslateBounds(refused)
                : "The transfer could not be prepared from the export bounds. Run Fix, then send again.",
            StagingCompositionRefusal.UnsupportedPrimitive or StagingCompositionRefusal.InvalidSubScene => "A chart element could not be prepared for the transfer. Run Fix, then send again.",
            StagingCompositionRefusal.NullResolution
                or StagingCompositionRefusal.NullPort
                or StagingCompositionRefusal.MissingScene => ConstructionMessage,
            _ => "The transfer could not be prepared. Run Fix, then send again.",
        };

    /// <summary>Translates a write refusal behind a refused composition.</summary>
    /// <param name="refusal">The port's typed refusal.</param>
    /// <returns>One sentence naming the reason, never a remedy that cannot work.</returns>
    private static string TranslateWriteRefusal(ShapeWriteRefusal? refusal) => refusal switch
    {
        ShapeWriteRefusal.TargetProtected => "The worksheet is protected, so Excel refused to prepare the transfer. Unprotect it, then send again.",
        ShapeWriteRefusal.AlreadyExists => "A temporary shape with the same name is already on the sheet. Select and delete the shapes named GanttCreatorStaging, then send again.",
        ShapeWriteRefusal.InvalidGeometry => "Part of the chart has no visible area, so Excel cannot prepare it. Run Fix, then send again.",
        ShapeWriteRefusal.UnsupportedIdentifierForm or ShapeWriteRefusal.BlankIdentifier => "The transfer could not be prepared from the table. Run Fix, then send again.",
        ShapeWriteRefusal.NoActiveWorkbook or ShapeWriteRefusal.WorksheetMissing => "The workbook with the Gantt table is not available. Open it, then send again.",
        ShapeWriteRefusal.HostRejected
            or ShapeWriteRefusal.NotFound
            or ShapeWriteRefusal.InvalidOwnershipTag
            or ShapeWriteRefusal.NullRequest => "Excel refused to prepare the transfer.",
        null => "Excel refused to prepare the transfer.",
        _ => "Excel refused to prepare the transfer.",
    };

    /// <summary>Translates a table-read refusal into one actionable line.</summary>
    /// <param name="refusal">The reader's typed reason.</param>
    /// <returns>The user-facing sentence.</returns>
    private static string TranslateRead(GanttTableReadRefusalReason refusal) => refusal switch
    {
        GanttTableReadRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open a workbook, then send again.",
        GanttTableReadRefusalReason.TableMissing =>
            "Send to PowerPoint could not find a table named tblGanttData on the active sheet. Run Initialise sheet first.",
        GanttTableReadRefusalReason.TableUnreadable =>
            "Send to PowerPoint found tblGanttData but its columns no longer match the Gantt Creator schema. Run Fix to repair the table, then send again.",
        GanttTableReadRefusalReason.DateSystemUnsupported =>
            "This workbook uses the 1904 date system, which Gantt Creator does not support. Use a 1900-date workbook and send again.",
        _ => ConstructionMessage,
    };

    /// <summary>Translates a configuration-read refusal into one actionable line.</summary>
    /// <param name="refusal">The reader's typed reason.</param>
    /// <returns>The user-facing sentence.</returns>
    private static string TranslateConfig(ConfigReadRefusalReason refusal) => refusal switch
    {
        ConfigReadRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open a workbook, then send again.",
        ConfigReadRefusalReason.ConfigSheetMissing =>
            "The workbook has no Gantt Creator configuration. Run Initialise sheet first.",
        ConfigReadRefusalReason.TableMissing =>
            "The Gantt Creator configuration tables are missing. Run Initialise sheet, then send again.",
        ConfigReadRefusalReason.HeaderMismatch or ConfigReadRefusalReason.RowCountMismatch =>
            "The Gantt Creator configuration is incomplete. Run Fix to repair it, then send again.",
        ConfigReadRefusalReason.CatalogueMismatch
            or ConfigReadRefusalReason.CatalogueHashMismatch
            or ConfigReadRefusalReason.ValueOutOfRange
            or ConfigReadRefusalReason.BadColourFormat =>
            "The Gantt Creator configuration is out of date or invalid. Run Fix to repair it, then send again.",
        _ => "The workbook configuration could not be read. Run Fix to repair it, then send again.",
    };

    /// <summary>Translates a panel-grid refusal into one actionable line.</summary>
    /// <param name="refusal">The measurement port's typed reason.</param>
    /// <returns>The user-facing sentence.</returns>
    private static string TranslateMeasurement(PanelGridRefusalReason refusal) => refusal switch
    {
        PanelGridRefusalReason.NoActiveWorkbook =>
            "Gantt Creator needs an open workbook. Open a workbook, then send again.",
        PanelGridRefusalReason.TableMissing =>
            "Send to PowerPoint could not find tblGanttData on the active sheet. Run Initialise sheet first.",
        PanelGridRefusalReason.TargetProtected =>
            "The worksheet is protected, so the columns could not be measured. Unlock the sheet or run Refresh chart, then send again.",
        PanelGridRefusalReason.InvalidMeasurement =>
            "The worksheet columns could not be measured. Run Refresh chart to re-measure them, then send again.",
        _ => "The worksheet columns could not be measured. Run Refresh chart to re-measure them, then send again.",
    };

    /// <summary>Translates an export-bounds refusal into one actionable line.</summary>
    /// <param name="resolution">The resolution carrying the bounds and scene refusals.</param>
    /// <returns>The user-facing sentence.</returns>
    private static string TranslateBounds(ExportBoundsOutcome resolution) => resolution.Refusal switch
    {
        // Core owns this sentence; restating it here would be a second authority.
        ExportBoundsRefusal.EmptyVisibleTable => ExportBoundsRefusalMessages.EmptyVisibleTable,
        ExportBoundsRefusal.PanelThemeMissing => "The transfer could not be prepared: the table styles were unavailable.",
        ExportBoundsRefusal.InvalidSelection => "The transfer could not start because the export selection is not recognised.",
        ExportBoundsRefusal.InvalidBounds => "The chart has no area to send. Adjust the chart dates, then send again.",
        ExportBoundsRefusal.UnanchoredRequest => "The transfer could not be prepared from the chart's rows.",
        ExportBoundsRefusal.SceneBuildRefused => resolution.SceneRefusal switch
        {
            SceneBuilderRefusal.EmptyEvents => "The table has no events to send. Add rows or send the chart only.",
            SceneBuilderRefusal.InvalidPlotRange => "The chart's date range is invalid. Adjust the chart dates, then send again.",
            SceneBuilderRefusal.InvalidLayoutSettings => "The chart layout settings could not be resolved. Run Fix, then send again.",
            SceneBuilderRefusal.ChartAboveSheet => "The chart would be drawn above the top of the sheet. Move the table down, then send again.",
            SceneBuilderRefusal.UnresolvableStyle => "A row's style could not be resolved. Run Fix, then send again.",
            SceneBuilderRefusal.UnresolvableProjection => "The parent/child structure could not be resolved. Fix the ParentId cells, then send again.",
            SceneBuilderRefusal.NullRequest
                or SceneBuilderRefusal.NullMetrics
                or SceneBuilderRefusal.NullPanelGrid
                or SceneBuilderRefusal.NullBounds
                or SceneBuilderRefusal.LiveProfileCarriesPanel
                or SceneBuilderRefusal.UnresolvableLaneAnchor => ConstructionMessage,
            _ => ConstructionMessage,
        },
        ExportBoundsRefusal.NullRequest => ConstructionMessage,
        _ => ConstructionMessage,
    };

    /// <summary>Builds the count-only success sentence and its optional clauses.</summary>
    /// <param name="stagedShapeCount">The validated pasted child count.</param>
    /// <param name="droppedRows">The blocking-error row count left out of the transfer.</param>
    /// <param name="presentationCreated">Whether the transfer created the presentation (D-R7.2 announcement).</param>
    /// <returns>The user-facing sentence.</returns>
    internal static string SuccessMessage(
        int stagedShapeCount,
        int droppedRows,
        bool presentationCreated)
    {
        string message = string.Create(
            CultureInfo.InvariantCulture,
            $"Sent {stagedShapeCount} {(stagedShapeCount == 1 ? "shape" : "shapes")} to the active slide.");
        if (presentationCreated)
        {
            message += " PowerPoint was not open, so a new presentation was created.";
        }

        return droppedRows == 0
            ? message
            : message + " " + string.Create(
                CultureInfo.InvariantCulture,
                $"{droppedRows} row(s) with errors were not sent — fix them and send again.");
    }

    /// <summary>Presents a refusal and returns its typed outcome.</summary>
    /// <param name="refusal">The command-level refusal.</param>
    /// <param name="message">The one actionable line.</param>
    /// <param name="droppedRows">The blocking-error row count, when known.</param>
    /// <param name="presentationCreated">Whether a presentation was created before the refusal.</param>
    /// <param name="presenter">The failure presenter.</param>
    /// <returns>The refused outcome.</returns>
    private static SendToPowerPointOutcome Refuse(
        SendToPowerPointRefusal refusal,
        string message,
        int droppedRows,
        bool presentationCreated,
        Action<string> presenter)
    {
        CommandBoundary.SafePresent(presenter, message);
        return new SendToPowerPointOutcome(false, 0, droppedRows, presentationCreated, refusal, message);
    }

}
