using System.Globalization;
using GanttCreator.Core.Export;

namespace GanttCreator.Office;

/// <summary>
/// R7.5 D1: consolidates every stage outcome into the closed <see
/// cref="PowerPointTransferFailureKind"/> taxonomy, with one actionable message and one
/// technical record per failure.
/// </summary>
/// <remarks>
/// <para>
/// <b>One total function per source.</b> The Core decision table, the R6.6 copy result, the
/// R7.3 paste stage and the R7.4 placement stage each map through their own method; every
/// deployment-table refusal arm is consumed (the two non-refusal arms — proceed and announced
/// creation — are not failures).
/// </para>
/// <para>
/// <b>The mapper only refines a catch-all.</b> A specific stage classification is never
/// overridden (a <c>PastePayloadMissing</c> input stays <c>PastePayloadMissing</c>); only a
/// <c>HostRejected</c> HRESULT is refined into <c>UnsupportedBuild</c>, <c>HostBusy</c> or
/// <c>HostRejected</c>.
/// </para>
/// <para>
/// <b>Host classification is HRESULT membership, not exhaustion</b> (PD-7): a code in the
/// measured transient set is <see cref="PowerPointTransferFailureKind.HostBusy"/> wherever it
/// surfaced; <c>DISP_E_MEMBERNOTFOUND</c> is <see cref="PowerPointTransferFailureKind.UnsupportedBuild"/>;
/// anything else is <see cref="PowerPointTransferFailureKind.HostRejected"/>.
/// </para>
/// </remarks>
public static class PowerPointTransferFailureMapper
{
    /// <summary>
    /// The <c>DISP_E_MEMBERNOTFOUND</c> HRESULT (<c>0x80020006</c>), measured on the reference
    /// host in R7.1: an unsupported build, never a transient retry.
    /// </summary>
    public const int DispMemberNotFound = unchecked((int)0x80020006);

    /// <summary>
    /// Classifies a COM failure code into the host-failure taxonomy members (PD-7).
    /// </summary>
    /// <param name="hresult">The COM failure code, or <see langword="null"/>.</param>
    /// <param name="policy">The transient filter, or <see langword="null"/> for the documented candidates.</param>
    /// <returns>The host-failure kind.</returns>
    public static PowerPointTransferFailureKind ClassifyHostFailure(
        int? hresult,
        PowerPointTransientRetryPolicy? policy = null)
    {
        if (hresult is not int code)
        {
            return PowerPointTransferFailureKind.HostRejected;
        }

        if (code == DispMemberNotFound)
        {
            return PowerPointTransferFailureKind.UnsupportedBuild;
        }

        PowerPointTransientRetryPolicy effective = policy ?? PowerPointTransientRetryPolicy.DocumentedCandidates;
        return effective.IsTransient(code)
            ? PowerPointTransferFailureKind.HostBusy
            : PowerPointTransferFailureKind.HostRejected;
    }

    /// <summary>Maps the R7.2 target decision's refusal arm into the taxonomy.</summary>
    /// <param name="decision">The resolved target decision.</param>
    /// <param name="policy">The transient filter, or <see langword="null"/> for the documented candidates.</param>
    /// <returns>The typed failure, or <see langword="null"/> when the decision did not refuse.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="decision"/> is <see langword="null"/>.</exception>
    public static PowerPointTransferFailure? MapTargetDecision(
        PowerPointTargetDecision decision,
        PowerPointTransientRetryPolicy? policy = null)
    {
        ArgumentNullException.ThrowIfNull(decision);

        return decision.Action == PowerPointTargetAction.Refuse && decision.Refusal is { } refusal
            ? MapTargetRefusal(refusal, decision.HResult, policy)
            : null;
    }

    private static PowerPointTransferFailure MapTargetRefusal(
        PowerPointTargetRefusal refusal,
        int? hresult,
        PowerPointTransientRetryPolicy? policy) =>
        refusal switch
        {
            PowerPointTargetRefusal.AttachRefused => HostFailure(hresult, policy, "attach"),
            PowerPointTargetRefusal.FactsUnreadable => HostFailure(hresult, policy, "target-read"),
            PowerPointTargetRefusal.SlideShowActive => Simple(PowerPointTransferFailureKind.SlideShowActive, null, "target"),
            PowerPointTargetRefusal.ProtectedView => Simple(PowerPointTransferFailureKind.PresentationProtectedView, null, "target"),
            PowerPointTargetRefusal.ReadOnlyPresentation => Simple(PowerPointTransferFailureKind.PresentationReadOnly, null, "target"),
            PowerPointTargetRefusal.NoWindow => Simple(PowerPointTransferFailureKind.NoWindow, null, "target"),
            PowerPointTargetRefusal.UnsupportedView => Simple(PowerPointTransferFailureKind.UnsupportedView, null, "target"),
            PowerPointTargetRefusal.NoSlide => Simple(PowerPointTransferFailureKind.NoSlide, null, "target"),
            PowerPointTargetRefusal.NoActiveSlide => Simple(PowerPointTransferFailureKind.NoSlide, null, "target"),
            _ => Simple(PowerPointTransferFailureKind.HostRejected, null, "target"),
        };

    /// <summary>
    /// Maps an R7.6 creation observation into the taxonomy: a refused creation is
    /// <see cref="PowerPointTransferFailureKind.NoPresentation"/> (PD-2).
    /// </summary>
    /// <param name="creation">The creation observation.</param>
    /// <returns>The typed failure, or <see langword="null"/> when the creation succeeded.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="creation"/> is <see langword="null"/>.</exception>
    public static PowerPointTransferFailure? MapCreation(PowerPointPresentationCreation creation)
    {
        ArgumentNullException.ThrowIfNull(creation);

        return creation.Created
            ? null
            : new PowerPointTransferFailure(
                PowerPointTransferFailureKind.NoPresentation,
                creation.HResult,
                MessageFor(PowerPointTransferFailureKind.NoPresentation),
                BuildRecord(PowerPointTransferFailureKind.NoPresentation, "create", creation.HResult, null));
    }

    /// <summary>
    /// Maps the R6.6 copy result into the taxonomy. Only a busy or unavailable clipboard is a
    /// transfer failure; the other copy refusals happen before PowerPoint is touched and stay
    /// the R6.8 command's Excel-side messages (out of R7.5 scope).
    /// </summary>
    /// <param name="result">The copy result, or <see langword="null"/>.</param>
    /// <returns>The typed failure, or <see langword="null"/> when the copy never ran or was not
    /// a clipboard failure.</returns>
    public static PowerPointTransferFailure? MapCopyResult(StagingCopyResult? result) =>
        result is null || result.Succeeded || result.Refusal != StagedCopyRefusal.ClipboardUnavailable
            ? null
            : new PowerPointTransferFailure(
                PowerPointTransferFailureKind.ClipboardUnavailable,
                null,
                MessageFor(PowerPointTransferFailureKind.ClipboardUnavailable),
                BuildRecord(
                    PowerPointTransferFailureKind.ClipboardUnavailable,
                    "copy",
                    null,
                    result.Detail));

    /// <summary>Maps the R7.3 paste stage result into the taxonomy.</summary>
    /// <param name="result">The paste stage result.</param>
    /// <param name="policy">The transient filter, or <see langword="null"/> for the documented candidates.</param>
    /// <returns>The typed failure, or <see langword="null"/> when the paste validated.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="result"/> is <see langword="null"/>.</exception>
    public static PowerPointTransferFailure? MapPasteStage(
        PowerPointPasteStageResult result,
        PowerPointTransientRetryPolicy? policy = null)
    {
        ArgumentNullException.ThrowIfNull(result);

        if (result.Succeeded)
        {
            return null;
        }

        if (result.Refusal is { } refusal)
        {
            return refusal switch
            {
                PowerPointPasteStageRefusal.PastePayloadMissing => Simple(
                    PowerPointTransferFailureKind.PastePayloadMissing,
                    result.HResult,
                    "paste"),
                PowerPointPasteStageRefusal.PartialDocument => Simple(
                    PowerPointTransferFailureKind.PartialDocument,
                    result.HResult,
                    "paste"),
                PowerPointPasteStageRefusal.NoSlide => Simple(
                    PowerPointTransferFailureKind.NoSlide,
                    null,
                    "paste"),
                PowerPointPasteStageRefusal.HostRejected => HostFailure(result.HResult, policy, "paste"),
                _ => HostFailure(result.HResult, policy, "paste"),
            };
        }

        if (result.Validation is { Succeeded: false } validation)
        {
            return new PowerPointTransferFailure(
                PowerPointTransferFailureKind.PasteValidationFailed,
                null,
                MessageFor(PowerPointTransferFailureKind.PasteValidationFailed),
                BuildRecord(
                    PowerPointTransferFailureKind.PasteValidationFailed,
                    "paste-validation",
                    null,
                    validation.Refusal?.ToString()));
        }

        // Defensive: a refused result with neither a stage refusal nor a failed validation is a
        // host rejection with no further detail.
        return Simple(PowerPointTransferFailureKind.HostRejected, null, "paste");
    }

    /// <summary>Maps the R7.4 placement stage result into the taxonomy.</summary>
    /// <param name="result">The placement stage result.</param>
    /// <param name="policy">The transient filter, or <see langword="null"/> for the documented candidates.</param>
    /// <returns>The typed failure, or <see langword="null"/> when the placement was applied and verified.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="result"/> is <see langword="null"/>.</exception>
    public static PowerPointTransferFailure? MapPlacement(
        PowerPointPlacementResult result,
        PowerPointTransientRetryPolicy? policy = null)
    {
        ArgumentNullException.ThrowIfNull(result);

        if (result.Succeeded)
        {
            return null;
        }

        if (result.FitRefusal is { } fit)
        {
            // A degenerate fit reaches the stage only as a caller defect: R7.6's pre-flight
            // refuses an invalid resolved Bounds before staging (R7.4 Step 0).
            return new PowerPointTransferFailure(
                PowerPointTransferFailureKind.InvalidTransferRequest,
                null,
                MessageFor(PowerPointTransferFailureKind.InvalidTransferRequest),
                BuildRecord(
                    PowerPointTransferFailureKind.InvalidTransferRequest,
                    "placement-fit",
                    null,
                    fit.ToString()));
        }

        return result.Refusal switch
        {
            PowerPointPlacementRefusal.AlreadyApplied => Simple(
                PowerPointTransferFailureKind.TransferSequenceViolated,
                null,
                "placement"),
            PowerPointPlacementRefusal.GroupMissing => Simple(
                PowerPointTransferFailureKind.HostRejected,
                null,
                "placement-group"),
            PowerPointPlacementRefusal.ReadBackMismatch => Simple(
                PowerPointTransferFailureKind.HostRejected,
                null,
                "placement-readback"),
            PowerPointPlacementRefusal.NoSlide => Simple(
                PowerPointTransferFailureKind.NoSlide,
                null,
                "placement"),
            PowerPointPlacementRefusal.SlideSizeUnavailable => result.HResult is null
                ? Simple(PowerPointTransferFailureKind.TransferSequenceViolated, null, "placement-size")
                : HostFailure(result.HResult, policy, "placement-size"),
            PowerPointPlacementRefusal.HostRejected => HostFailure(result.HResult, policy, "placement"),
            _ => Simple(PowerPointTransferFailureKind.HostRejected, null, "placement"),
        };
    }

    /// <summary>Builds the construction-member failure for an invalid request or argument.</summary>
    /// <param name="guard">The guard that refused (a code-owned constant).</param>
    /// <returns>The typed failure.</returns>
    /// <exception cref="ArgumentException"><paramref name="guard"/> is blank.</exception>
    public static PowerPointTransferFailure InvalidRequest(string guard)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(guard);
        return new PowerPointTransferFailure(
            PowerPointTransferFailureKind.InvalidTransferRequest,
            null,
            MessageFor(PowerPointTransferFailureKind.InvalidTransferRequest),
            BuildRecord(PowerPointTransferFailureKind.InvalidTransferRequest, "construction", null, guard));
    }

    /// <summary>Builds the construction-member failure for an out-of-order transfer.</summary>
    /// <param name="reason">The sequence defect (a code-owned constant).</param>
    /// <returns>The typed failure.</returns>
    /// <exception cref="ArgumentException"><paramref name="reason"/> is blank.</exception>
    public static PowerPointTransferFailure SequenceViolated(string reason)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(reason);
        return new PowerPointTransferFailure(
            PowerPointTransferFailureKind.TransferSequenceViolated,
            null,
            MessageFor(PowerPointTransferFailureKind.TransferSequenceViolated),
            BuildRecord(PowerPointTransferFailureKind.TransferSequenceViolated, "sequence", null, reason));
    }

    /// <summary>
    /// Builds the ownership failure for a refused release on the teardown path (PD-6): the
    /// operation already answered, so this is reported, never thrown.
    /// </summary>
    /// <param name="hresult">The release refusal code, or <see langword="null"/>.</param>
    /// <returns>The typed failure.</returns>
    public static PowerPointTransferFailure OwnershipFailed(int? hresult) =>
        new(
            PowerPointTransferFailureKind.OwnershipFailed,
            hresult,
            MessageFor(PowerPointTransferFailureKind.OwnershipFailed),
            BuildRecord(PowerPointTransferFailureKind.OwnershipFailed, "ownership", hresult, null));

    private static PowerPointTransferFailure HostFailure(
        int? hresult,
        PowerPointTransientRetryPolicy? policy,
        string stage) =>
        Simple(ClassifyHostFailure(hresult, policy), hresult, stage);

    private static PowerPointTransferFailure Simple(
        PowerPointTransferFailureKind kind,
        int? hresult,
        string stage) =>
        new(kind, hresult, MessageFor(kind), BuildRecord(kind, stage, hresult, null));

    private static string BuildRecord(
        PowerPointTransferFailureKind kind,
        string stage,
        int? hresult,
        string? detail)
    {
        string head = string.Create(
            CultureInfo.InvariantCulture,
            $"r7.5 transfer failure kind={kind} stage={stage}");
        string codePart = hresult is int code
            ? string.Create(CultureInfo.InvariantCulture, $" hresult=0x{code:X8}")
            : string.Empty;
        string detailPart = string.IsNullOrEmpty(detail) ? string.Empty : " detail=" + detail;
        return head + codePart + detailPart;
    }

    private static string MessageFor(PowerPointTransferFailureKind kind) => kind switch
    {
        PowerPointTransferFailureKind.NoPresentation =>
            "PowerPoint had no presentation open and a new one could not be created. Open or create a presentation and try again.",
        PowerPointTransferFailureKind.PresentationReadOnly =>
            "The active presentation is read-only or password-protected. Open an editable copy and try again.",
        PowerPointTransferFailureKind.PresentationProtectedView =>
            "The document is in Protected View. Enable editing on the document and try again.",
        PowerPointTransferFailureKind.NoWindow =>
            "The presentation is open without a document window. Restore its window and try again.",
        PowerPointTransferFailureKind.NoSlide =>
            "The presentation has no active slide. Open a slide and try again.",
        PowerPointTransferFailureKind.UnsupportedView =>
            "The current view cannot receive pasted shapes. Switch to Normal, Slide or Notes page view and try again.",
        PowerPointTransferFailureKind.SlideShowActive =>
            "A slide show is running. End the show and try again.",
        PowerPointTransferFailureKind.ClipboardUnavailable =>
            "The clipboard is busy or unavailable. Close the application holding the clipboard and try again.",
        PowerPointTransferFailureKind.PastePayloadMissing =>
            "The clipboard holds no shape data to paste. Refresh the chart, copy it again, then send it.",
        PowerPointTransferFailureKind.PartialDocument =>
            "The source content is only partially available. Wait for it to finish downloading and try again.",
        PowerPointTransferFailureKind.PasteValidationFailed =>
            "PowerPoint pasted shapes that did not match the chart. The pasted shapes were removed; try again.",
        PowerPointTransferFailureKind.HostBusy =>
            "PowerPoint is busy. Close any open dialog and try again.",
        PowerPointTransferFailureKind.HostRejected =>
            "PowerPoint refused the operation. Try again.",
        PowerPointTransferFailureKind.UnsupportedBuild =>
            "This PowerPoint version does not support the operation. Update Office and try again.",
        PowerPointTransferFailureKind.OwnershipFailed =>
            "A PowerPoint object could not be released. Restart PowerPoint before trying again.",
        PowerPointTransferFailureKind.InvalidTransferRequest =>
            "The transfer request was invalid. Refresh the chart and try again.",
        PowerPointTransferFailureKind.TransferSequenceViolated =>
            "The transfer ran in an unexpected order. Refresh the chart and try again.",
        _ => "The transfer to PowerPoint failed. Try again.",
    };
}
