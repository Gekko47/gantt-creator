namespace GanttCreator.Office;

/// <summary>
/// The closed typed transfer-failure taxonomy (R7.5 D1). Every transfer failure mode maps to
/// exactly one member; the two members beyond D1's recorded list — <see
/// cref="InvalidTransferRequest"/> and <see cref="TransferSequenceViolated"/> — are the work
/// item's "construction members" (null/argument guards and sequence defects).
/// </summary>
/// <remarks>
/// <para>
/// <b>One member per user-visible mode.</b> A member is a stable identity for a failure the
/// command translates to one actionable message; the message text travels on the failure
/// record, and R7.6 presents it. Undefined visual or host behaviour is never inferred into a
/// new member — the list is closed (work item D1).
/// </para>
/// <para>
/// <b>Non-silent by construction.</b> A failure always carries a technical record naming the
/// stage, the taxonomy member and, where available, the COM HRESULT; the record is written to
/// the redacting rolling log by the caller.
/// </para>
/// </remarks>
public enum PowerPointTransferFailureKind
{
    /// <summary>No presentation was open and the announced creation was refused.</summary>
    NoPresentation = 0,

    /// <summary>The resolved active presentation is read-only or password-protected.</summary>
    PresentationReadOnly = 1,

    /// <summary>The document is in Protected View (foreground, or the only surface with no deck).</summary>
    PresentationProtectedView = 2,

    /// <summary>A presentation is open but has no document window, so no active slide exists.</summary>
    NoWindow = 3,

    /// <summary>The presentation has no slides, or the view exposes no active slide.</summary>
    NoSlide = 4,

    /// <summary>The active view is not slide, notes page or normal.</summary>
    UnsupportedView = 5,

    /// <summary>A slide-show window is active, so a shape paste is not a valid target.</summary>
    SlideShowActive = 6,

    /// <summary>The clipboard is busy or unavailable.</summary>
    ClipboardUnavailable = 7,

    /// <summary>The clipboard holds no payload the host could paste.</summary>
    PastePayloadMissing = 8,

    /// <summary>The source content is only partially downloaded.</summary>
    PartialDocument = 9,

    /// <summary>The pasted range did not match the staged chart (count, groups, type or bounds).</summary>
    PasteValidationFailed = 10,

    /// <summary>PowerPoint rejected a call with a measured transient HRESULT, after bounded retries.</summary>
    HostBusy = 11,

    /// <summary>PowerPoint refused a call with an unrecognised COM failure.</summary>
    HostRejected = 12,

    /// <summary>The host reports a missing member (<c>DISP_E_MEMBERNOTFOUND</c>) — an unsupported build.</summary>
    UnsupportedBuild = 13,

    /// <summary>An owned PowerPoint proxy could not be released on the teardown path.</summary>
    OwnershipFailed = 14,

    /// <summary>A transfer request or argument was invalid (a construction member).</summary>
    InvalidTransferRequest = 15,

    /// <summary>The transfer ran out of order, e.g. a disposed adapter or a consumed placement (a construction member).</summary>
    TransferSequenceViolated = 16,
}

/// <summary>
/// One typed transfer failure: the taxonomy member, the COM failure code where available, an
/// actionable user message, and the technical record for the rolling log (R7.5 D1).
/// </summary>
/// <param name="Kind">The closed taxonomy member.</param>
/// <param name="HResult">The COM failure code, or <see langword="null"/> when the failure carried none.</param>
/// <param name="Message">
/// The actionable user message. R7.6 presents exactly this text; the failure never shows a
/// dialog itself.
/// </param>
/// <param name="TechnicalRecord">
/// The redaction-ready technical record naming the stage, the taxonomy member and the HRESULT.
/// Never contains schedule content.
/// </param>
public sealed record PowerPointTransferFailure(
    PowerPointTransferFailureKind Kind,
    int? HResult,
    string Message,
    string TechnicalRecord);
