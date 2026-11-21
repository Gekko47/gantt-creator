using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// The typed reasons the R7.4 placement stage refuses an application (work item D1, audit A-11).
/// </summary>
/// <remarks>
/// The members are named for R7.5's recorded taxonomy where the two overlap
/// (<c>GroupMissing</c>, <c>HostRejected</c>, <c>NoSlide</c>, <c>AlreadyApplied</c>) so R7.5
/// consolidates instead of renaming — the <see cref="PowerPointPasteStageRefusal"/> precedent.
/// Every member has a positive test.
/// </remarks>
public enum PowerPointPlacementRefusal
{
    /// <summary>
    /// The placement had already been consumed — the work item D1 once-rule. A placement is
    /// single-use; a retry re-runs the fit.
    /// </summary>
    AlreadyApplied = 0,

    /// <summary>The forwarded group identifier does not name an <c>msoGroup</c> on the active slide.</summary>
    GroupMissing = 1,

    /// <summary>The host refused a write or a read with a COM failure; the HRESULT is carried.</summary>
    HostRejected = 2,

    /// <summary>The applied geometry differs from the placement by more than the repository tolerance.</summary>
    ReadBackMismatch = 3,

    /// <summary>The application or the active slide was unavailable.</summary>
    NoSlide = 4,

    /// <summary>The active slide's size could not be read, so no fit could be computed.</summary>
    SlideSizeUnavailable = 5,
}

/// <summary>The measured per-member difference between the applied and requested placement.</summary>
/// <param name="LeftPt">The applied left minus the requested left, in points.</param>
/// <param name="TopPt">The applied top minus the requested top, in points.</param>
/// <param name="WidthPt">The applied width minus the requested width, in points.</param>
/// <param name="HeightPt">The applied height minus the requested height, in points.</param>
public sealed record PowerPointPlacementDeltas(double LeftPt, double TopPt, double WidthPt, double HeightPt)
{
    /// <summary>Gets the largest absolute difference across the four members, in points.</summary>
    public double MaxPt => Math.Max(
        Math.Max(Math.Abs(LeftPt), Math.Abs(TopPt)),
        Math.Max(Math.Abs(WidthPt), Math.Abs(HeightPt)));
}

/// <summary>The typed outcome of the R7.4 placement stage.</summary>
/// <param name="Succeeded">Whether the placement was applied and verified.</param>
/// <param name="Refusal">The typed placement refusal, or <see langword="null"/>.</param>
/// <param name="HResult">The COM failure code carried by a <see cref="PowerPointPlacementRefusal.HostRejected"/> refusal, or <see langword="null"/>.</param>
/// <param name="FitRefusal">The typed Core fit refusal when no placement could be computed, or <see langword="null"/>.</param>
/// <param name="Placement">The single placement that was attempted, or <see langword="null"/>.</param>
/// <param name="Observation">The port's measured observation, or <see langword="null"/> on a pre-write refusal.</param>
/// <param name="Deltas">The measured read-back deltas, or <see langword="null"/> on a pre-verification refusal.</param>
public sealed record PowerPointPlacementResult(
    bool Succeeded,
    PowerPointPlacementRefusal? Refusal,
    int? HResult,
    SlideFitRefusal? FitRefusal,
    SlidePlacement? Placement,
    PowerPointPlacementObservation? Observation,
    PowerPointPlacementDeltas? Deltas);

/// <summary>
/// R7.4: fits the resolved bounds inside the slide's margins, consumes the single placement once,
/// applies it to the pasted group, and verifies the read-back.
/// </summary>
/// <remarks>
/// <para>
/// <b>Once, and only once (D1).</b> <see cref="Apply"/> consumes the placement before the port
/// call, so a second attempt is refused even when the first failed. <see cref="SlidePlacement.TryConsume"/>
/// is the Core half; this stage is the enforcer.
/// </para>
/// <para>
/// <b>Non-throwing is not success.</b> A returned observation whose applied geometry differs from
/// the placement by more than <see cref="HostOwnedGeometry.GeometryMatchTolerancePt"/> is a
/// <see cref="PowerPointPlacementRefusal.ReadBackMismatch"/> refusal, not a quiet success.
/// </para>
/// <para>
/// <b>One technical record per refusal.</b> The optional sink receives exactly one redaction-ready
/// line per refused case (the AddIn wires the redacting rolling log); success carries the
/// measured deltas on the result instead.
/// </para>
/// <para>
/// <b>The inset is a caller input (PD-6).</b> The D-R7.4 margin policy is an owed product
/// decision; this stage takes <c>insetPt</c> so the policy constant can land without changing
/// this type.
/// </para>
/// </remarks>
public sealed class PowerPointPlacementStage
{
    private readonly IPowerPointPlacementPort _port;
    private readonly Action<string>? _technicalRecord;

    /// <summary>Creates a stage over the PowerPoint placement port.</summary>
    /// <param name="port">The port the stage drives.</param>
    /// <param name="technicalRecord">Receives one technical record per refusal, or <see langword="null"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="port"/> is <see langword="null"/>.</exception>
    public PowerPointPlacementStage(IPowerPointPlacementPort port, Action<string>? technicalRecord = null)
    {
        ArgumentNullException.ThrowIfNull(port);
        _port = port;
        _technicalRecord = technicalRecord;
    }

    /// <summary>
    /// Reads the slide size, fits <paramref name="groupBounds"/> inside the policy margins, and
    /// applies the resulting placement once.
    /// </summary>
    /// <param name="groupId">The pasted group's identifier, forwarded by the R7.3 paste stage.</param>
    /// <param name="groupBounds">The resolved zero-origin <see cref="ExportBounds.Bounds"/>.</param>
    /// <param name="insetPt">The policy margin in points, supplied by the caller (PD-6).</param>
    /// <returns>The typed stage outcome; every expected host state is a result, never an exception.</returns>
    public PowerPointPlacementResult Place(int groupId, RectD groupBounds, double insetPt)
    {
        PowerPointSlideSizeFacts size = _port.ReadSlideSize();
        if (!size.ReadSucceeded)
        {
            Record("r7.4 placement refused slide-size-unavailable", size.HResult);
            return Refused(PowerPointPlacementRefusal.SlideSizeUnavailable, size.HResult);
        }

        SlideFitOutcome fit = SlideFit.Fit(groupBounds, size.WidthPt, size.HeightPt, insetPt);
        if (!fit.Succeeded)
        {
            Record("r7.4 placement refused fit=" + fit.Refusal, null);
            return new PowerPointPlacementResult(
                Succeeded: false,
                Refusal: null,
                HResult: null,
                FitRefusal: fit.Refusal,
                Placement: null,
                Observation: null,
                Deltas: null);
        }

        return Apply(groupId, fit.Placement!);
    }

    /// <summary>
    /// Applies an already-fitted placement exactly once and verifies the read-back.
    /// </summary>
    /// <param name="groupId">The pasted group's identifier, forwarded by the R7.3 paste stage.</param>
    /// <param name="placement">The single placement produced by <see cref="SlideFit.Fit"/>.</param>
    /// <returns>The typed stage outcome; a second application of the same placement is refused.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="placement"/> is <see langword="null"/>.</exception>
    public PowerPointPlacementResult Apply(int groupId, SlidePlacement placement)
    {
        ArgumentNullException.ThrowIfNull(placement);

        if (!placement.TryConsume())
        {
            Record("r7.4 placement refused already-applied", null);
            return Refused(PowerPointPlacementRefusal.AlreadyApplied, null);
        }

        PowerPointPlacementObservation observation = _port.ApplyPlacement(groupId, placement);

        PowerPointPlacementRefusal? refusal = observation.Outcome switch
        {
            PowerPointPlacementApplyOutcome.Applied => null,
            PowerPointPlacementApplyOutcome.NoSlide => PowerPointPlacementRefusal.NoSlide,
            PowerPointPlacementApplyOutcome.GroupMissing => PowerPointPlacementRefusal.GroupMissing,
            PowerPointPlacementApplyOutcome.HostRejected => PowerPointPlacementRefusal.HostRejected,
            _ => PowerPointPlacementRefusal.HostRejected,
        };

        if (refusal is not null)
        {
            Record("r7.4 placement refused " + refusal.Value, observation.HResult);
            return Refused(refusal.Value, observation.HResult);
        }

        PowerPointPlacementDeltas deltas = new(
            LeftPt: observation.LeftPt - placement.Left,
            TopPt: observation.TopPt - placement.Top,
            WidthPt: observation.WidthPt - placement.Width,
            HeightPt: observation.HeightPt - placement.Height);

        if (deltas.MaxPt > HostOwnedGeometry.GeometryMatchTolerancePt)
        {
            Record(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"r7.4 placement refused read-back-mismatch maxdelta={deltas.MaxPt:0.###}"),
                null);
            return new PowerPointPlacementResult(
                Succeeded: false,
                Refusal: PowerPointPlacementRefusal.ReadBackMismatch,
                HResult: null,
                FitRefusal: null,
                Placement: placement,
                Observation: observation,
                Deltas: deltas);
        }

        return new PowerPointPlacementResult(
            Succeeded: true,
            Refusal: null,
            HResult: null,
            FitRefusal: null,
            Placement: placement,
            Observation: observation,
            Deltas: deltas);
    }

    private static PowerPointPlacementResult Refused(
        PowerPointPlacementRefusal refusal,
        int? hresult) =>
        new(
            Succeeded: false,
            Refusal: refusal,
            HResult: hresult,
            FitRefusal: null,
            Placement: null,
            Observation: null,
            Deltas: null);

    private void Record(string message, int? hresult)
    {
        if (_technicalRecord is null)
        {
            return;
        }

        _technicalRecord(hresult is null
            ? message
            : string.Create(CultureInfo.InvariantCulture, $"{message} hresult=0x{hresult.Value:X8}"));
    }
}
