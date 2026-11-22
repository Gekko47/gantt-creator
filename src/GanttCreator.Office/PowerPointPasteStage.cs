using GanttCreator.Core.Export;

namespace GanttCreator.Office;

/// <summary>
/// The typed stage failures a <c>PasteSpecial(ppPasteShape)</c> call can produce before any
/// range exists to validate (R7.3 D4), named identically to R7.5's recorded taxonomy so
/// R7.5 consolidates instead of renaming.
/// </summary>
/// <remarks>
/// These are stage failures, not paste-validation failures: the validator only ever sees a
/// returned range. A returned range that fails validation is carried by
/// <see cref="PowerPointPasteStageResult.Validation"/>, and R7.5 maps that outcome to its
/// <c>PasteValidationFailed</c> failure.
/// </remarks>
public enum PowerPointPasteStageRefusal
{
    /// <summary>
    /// The clipboard carried no payload the host could paste — an empty or unusable clipboard.
    /// Microsoft Learn documents an error when there is no data on the Clipboard.
    /// </summary>
    PastePayloadMissing = 0,

    /// <summary>
    /// The source content is not fully downloaded (a partial document), which Microsoft Learn
    /// documents as a paste failure. This member gains an HRESULT filter only from measured
    /// evidence; until a build exhibits it, it is unreachable and unmeasured COM failures
    /// map to <see cref="HostRejected"/>.
    /// </summary>
    PartialDocument = 1,

    /// <summary>The host rejected the paste with a COM failure the classifier does not recognise.</summary>
    HostRejected = 2,

    /// <summary>
    /// The paste target was unavailable — the application was not attached, or the active
    /// slide could not be resolved. R7.2's decision table owns the pre-paste slide cases;
    /// this member covers the slide vanishing (or the port never being attached) at paste time.
    /// </summary>
    NoSlide = 3,
}

/// <summary>
/// The typed outcome of the R7.3 paste stage.
/// </summary>
/// <param name="Succeeded">Whether the paste materialised a range that passed validation.</param>
/// <param name="Refusal">
/// The typed stage failure when the paste itself did not produce a range (D4); otherwise
/// <see langword="null"/>. A returned range that fails validation leaves this
/// <see langword="null"/> and carries the refusal on <paramref name="Validation"/> — a
/// non-throwing call is not success (D3).
/// </param>
/// <param name="HResult">The COM failure code carried by a stage refusal, or <see langword="null"/>.</param>
/// <param name="Validation">The Core validation outcome when a range was returned, or <see langword="null"/> on a stage refusal.</param>
/// <param name="GroupName">
/// The pasted group's name, forwarded for R7.4 (D2), or <see langword="null"/> when no group
/// was located.
/// </param>
/// <param name="GroupId">
/// The pasted group's identifier, forwarded for R7.4 (D2), or <see langword="null"/> when no
/// group was located.
/// </param>
public sealed record PowerPointPasteStageResult(
    bool Succeeded,
    PowerPointPasteStageRefusal? Refusal,
    int? HResult,
    PowerPointPasteValidation? Validation,
    string? GroupName,
    int? GroupId);

/// <summary>
/// Classifies a refused paste observation into R7.3 D4's typed stage failures before
/// validation.
/// </summary>
/// <remarks>
/// <para>
/// <b>Evidence-first.</b> The clipboard-empty mapping is the documented candidate HRESULT for
/// "Invalid request. Clipboard is empty or contains data which may not be pasted here." — the
/// same <c>0x80048240</c> family the R7.1/R7.2 probes measured for typed absence. It remains
/// pending the R7.3 Step 0 live confirmation and must be updated only from a measured run; an
/// unrecognised COM failure is <see cref="PowerPointPasteStageRefusal.HostRejected"/> with
/// its HRESULT still carried on the result.
/// </para>
/// <para>
/// <b>No HRESULT means no target.</b> The adapter reports a refused observation without a COM
/// code when the application or its active slide could not be resolved, so a null code maps
/// to <see cref="PowerPointPasteStageRefusal.NoSlide"/>.
/// </para>
/// </remarks>
internal static class PowerPointPasteFailureClassifier
{
    /// <summary>
    /// The documented candidate HRESULT for an empty or unusable clipboard payload. Measured
    /// R7.1/R7.2 absence probes reported the same code family; the R7.3 Step 0 probe confirms
    /// or corrects this value from evidence.
    /// </summary>
    internal const int PastePayloadMissingCandidateHResult = unchecked((int)0x80048240);

    /// <summary>
    /// Builds the refused observation the adapter returns from a caught COM failure. This is
    /// the host-free seam the contract tests drive: a <c>PasteSpecial</c> exception becomes
    /// this observation, and the stage maps the observation to the D4 refusal.
    /// </summary>
    /// <param name="hresult">The COM failure code, or <see langword="null"/> for an unavailable target.</param>
    /// <returns>The refused observation.</returns>
    internal static PowerPointPasteObservation RefusedObservation(int? hresult) =>
        new(
            PasteSucceeded: false,
            HResult: hresult,
            Count: 0,
            GroupCount: 0,
            GroupShapeType: null,
            GroupChildCount: 0,
            GroupName: null,
            GroupId: null,
            WidthPt: 0,
            HeightPt: 0);

    /// <summary>Classifies a refused observation's HRESULT into the typed stage failure.</summary>
    /// <param name="hresult">The observation's COM failure code, or <see langword="null"/>.</param>
    /// <returns>The typed stage failure.</returns>
    internal static PowerPointPasteStageRefusal Classify(int? hresult) =>
        hresult is null
            ? PowerPointPasteStageRefusal.NoSlide
            : hresult.Value == PastePayloadMissingCandidateHResult
                ? PowerPointPasteStageRefusal.PastePayloadMissing
                : PowerPointPasteStageRefusal.HostRejected;
}

/// <summary>
/// R7.3: calls the paste port, classifies a stage failure before validation, and validates the
/// returned range's facts against the staged expectation.
/// </summary>
/// <remarks>
/// <para>
/// <b>Classification before validation (D4).</b> A refused observation becomes one typed
/// stage failure and the Core validator never runs; only a returned range reaches
/// <see cref="PowerPointPasteValidator"/>.
/// </para>
/// <para>
/// <b>Non-throwing is not success (D3).</b> A returned range that fails any rule yields
/// <see cref="PowerPointPasteStageResult.Succeeded"/> false with the refusal on the
/// validation, even though no COM call raised.
/// </para>
/// <para>
/// <b>Success forwards the facts (D2).</b> The result carries the validated facts through the
/// validation result and the group identity through <see cref="PowerPointPasteStageResult.GroupId"/>
/// and <see cref="PowerPointPasteStageResult.GroupName"/>, so R7.4 positions the group without
/// re-reading state the stage already has.
/// </para>
/// <para>
/// <b>Residue is R7.5's policy.</b> A validation refusal reports the typed failure and leaves
/// the pasted shapes in place; R7.5's outcome owns cleanup on every path, and ADR-0066 D3's
/// group-delete mechanic (delete the group, never the range) is recorded for it. No
/// user-facing path exists before R7.6 lands.
/// </para>
/// </remarks>
public sealed class PowerPointPasteStage
{
    private readonly IPowerPointPastePort _port;

    /// <summary>Creates a stage over the PowerPoint paste port.</summary>
    /// <param name="port">The port the stage drives.</param>
    /// <exception cref="ArgumentNullException"><paramref name="port"/> is <see langword="null"/>.</exception>
    public PowerPointPasteStage(IPowerPointPastePort port)
    {
        ArgumentNullException.ThrowIfNull(port);
        _port = port;
    }

    /// <summary>
    /// Pastes the clipboard onto the active slide and validates the returned range.
    /// </summary>
    /// <param name="expectation">The staged expectation the result must satisfy.</param>
    /// <returns>The typed stage outcome; every expected host state is a result, never an exception.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="expectation"/> is <see langword="null"/>.</exception>
    public PowerPointPasteStageResult Paste(PowerPointPasteExpectation expectation)
    {
        ArgumentNullException.ThrowIfNull(expectation);

        PowerPointPasteObservation observation = _port.PasteSpecialShape();
        if (!observation.PasteSucceeded)
        {
            return new PowerPointPasteStageResult(
                Succeeded: false,
                Refusal: PowerPointPasteFailureClassifier.Classify(observation.HResult),
                HResult: observation.HResult,
                Validation: null,
                GroupName: observation.GroupName,
                GroupId: observation.GroupId);
        }

        PowerPointPasteFacts facts = new(
            Count: observation.Count,
            GroupCount: observation.GroupCount,
            GroupShapeType: observation.GroupShapeType,
            WidthPt: observation.WidthPt,
            HeightPt: observation.HeightPt);

        PowerPointPasteValidation validation =
            PowerPointPasteValidator.Validate(facts, expectation);

        return new PowerPointPasteStageResult(
            Succeeded: validation.Succeeded,
            Refusal: null,
            HResult: null,
            Validation: validation,
            GroupName: observation.GroupName,
            GroupId: observation.GroupId);
    }
}
