using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Export;

/// <summary>
/// The typed reasons the R7.4 slide fit refuses an input (work item CR-9, CR-10).
/// </summary>
/// <remarks>
/// <para>
/// Every member is one typed refusal with a positive test that constructs the bad input. The
/// members are ordered to mirror the total precedence <see cref="SlideFit.Fit"/> applies, so a
/// result names the first failing rule deterministically.
/// </para>
/// <para>
/// <b>Why a refusal rather than an exception on the group bounds.</b> <see cref="RectD"/> already
/// refuses a non-finite or negative extent when it is constructed, so the reachable bad group is
/// a zero extent (<c>default(RectD)</c>); the fit repeats the strictly-positive rule so a
/// degenerate group can never become a degenerate placement (CR-10). The slide size and the
/// inset are raw doubles, so their non-finite and non-positive arms are directly reachable.
/// </para>
/// </remarks>
public enum SlideFitRefusal
{
    /// <summary>The group bounds have a zero width or height.</summary>
    InvalidGroupBounds = 0,

    /// <summary>The slide width or height is not finite or not strictly positive.</summary>
    InvalidSlideSize = 1,

    /// <summary>The inset is not finite or is negative.</summary>
    InvalidInset = 2,

    /// <summary>The inset leaves no positive area inside the slide (the fit frame is degenerate).</summary>
    DegenerateFitFrame = 3,

    /// <summary>
    /// The computed scale is not finite or not positive, or the resulting geometry is not
    /// representable at display precision (it would collapse to zero).
    /// </summary>
    InvalidScale = 4,
}

/// <summary>The typed outcome of one <see cref="SlideFit.Fit"/> call (R7.4).</summary>
/// <param name="Succeeded">Whether a placement was produced.</param>
/// <param name="Placement">The single placement on success, or <see langword="null"/> on a refusal.</param>
/// <param name="Refusal">The typed refusal on failure, or <see langword="null"/> on success.</param>
public sealed record SlideFitOutcome(bool Succeeded, SlidePlacement? Placement, SlideFitRefusal? Refusal);

/// <summary>
/// The pure R7.4 fit: scales the resolved export bounds once to fit inside the slide's policy
/// margins, preserving the aspect ratio (work item D2, PD-2, PD-8).
/// </summary>
/// <remarks>
/// <para>
/// <b>The aspect source is the caller's group bounds, which is the zero-origin
/// <see cref="ExportBounds.Bounds"/></b> — never <c>ChartBounds</c> and never the pasted range
/// (work item D6). <c>PlacementDelta</c> is already consumed by the staging composition and is
/// not re-applied here or downstream.
/// </para>
/// <para>
/// <b>The margin is scene-space, so it scales with the group (D-R7.4, recorded 2026-11-22).</b>
/// The inset is the chart's own outer padding
/// (<see cref="ChartPaddingPt.Uniform"/>, 4 pt at scene scale), not a slide-space number: the
/// chart keeps the margin it was composed with, and the slide-space margin is
/// <c>inset × scale</c>. The closed form therefore puts the inset on the group's side of the
/// ratio —
/// <c>Scale = min(slideWidth / (groupWidth + 2·inset), slideHeight / (groupHeight + 2·inset))</c> —
/// and the placement is centred:
/// <c>Left = (slideWidth − groupWidth × Scale) / 2</c>, which equals <c>inset × Scale</c> on the
/// binding axis. A slide-space inset would instead grow the margin's share as the chart shrinks,
/// which is why it is not the recorded policy.
/// </para>
/// <para>
/// <b>Ratio preserved by construction.</b> One scale drives both axes, so width and height
/// cannot be computed independently and the aspect cannot drift.
/// </para>
/// <para>
/// <b>One rounding policy.</b> The four values that reach PowerPoint are snapped through
/// <see cref="GeometryMath.SnapToDisplayPrecision"/> once each, at the boundary;
/// <see cref="SlidePlacement.Scale"/> stays the exact mathematical factor. A raw geometry that
/// cannot be represented at display precision is a <see cref="SlideFitRefusal.InvalidScale"/>
/// refusal rather than an escaping exception.
/// </para>
/// <para>
/// <b>Total precedence.</b> Group bounds, slide size, inset, frame and scale are checked in that
/// order; the first failing rule names the refusal. The method is a pure function of its
/// arguments and therefore deterministic for identical inputs.
/// </para>
/// </remarks>
public static class SlideFit
{
    /// <summary>
    /// Fits <paramref name="groupBounds"/> inside the slide's margins, preserving its aspect ratio.
    /// </summary>
    /// <param name="groupBounds">
    /// The group bounds to fit — the resolved zero-origin <see cref="ExportBounds.Bounds"/>.
    /// </param>
    /// <param name="slideWidthPt">The active slide's width in points.</param>
    /// <param name="slideHeightPt">The active slide's height in points.</param>
    /// <param name="insetPt">
    /// The policy margin in the group's own (scene) coordinate system, applied to every side and
    /// scaled with the group — <see cref="ChartPaddingPt.Uniform"/> of the chart's outer padding.
    /// </param>
    /// <returns>The single placement, or a typed refusal; no exception escapes for a bad input.</returns>
    public static SlideFitOutcome Fit(
        RectD groupBounds,
        double slideWidthPt,
        double slideHeightPt,
        double insetPt)
    {
        // RectD's constructor already refuses a non-finite or negative extent, so the reachable
        // bad group is the zero extent; the strictly-positive rule is repeated here (CR-10) so a
        // degenerate group can never become a degenerate placement.
        if (groupBounds.Width <= 0 || groupBounds.Height <= 0)
        {
            return Refuse(SlideFitRefusal.InvalidGroupBounds);
        }

        if (!IsFiniteAndPositive(slideWidthPt) || !IsFiniteAndPositive(slideHeightPt))
        {
            return Refuse(SlideFitRefusal.InvalidSlideSize);
        }

        if (!double.IsFinite(insetPt) || insetPt < 0)
        {
            return Refuse(SlideFitRefusal.InvalidInset);
        }

        // The inset is scene-space, so an inset at or beyond half a slide dimension leaves no
        // room for the chart however small the group's share of it is. The frame subtraction is
        // therefore a policy sanity guard, not the fit itself.
        double frameWidthPt = slideWidthPt - (2 * insetPt);
        double frameHeightPt = slideHeightPt - (2 * insetPt);
        if (!IsFiniteAndPositive(frameWidthPt) || !IsFiniteAndPositive(frameHeightPt))
        {
            return Refuse(SlideFitRefusal.DegenerateFitFrame);
        }

        // The inset is added on the BRACKET side so the emitted slide-space margin is
        // inset × scale, which is the recorded D-R7.4 policy (see the type remarks).
        double scale = Math.Min(
            slideWidthPt / (groupBounds.Width + (2 * insetPt)),
            slideHeightPt / (groupBounds.Height + (2 * insetPt)));
        if (!IsFiniteAndPositive(scale))
        {
            return Refuse(SlideFitRefusal.InvalidScale);
        }

        // Both applied axes derive from the one scale, so the ratio cannot drift (D2), and the
        // group is centred, so the margin is symmetric and equals inset × scale on the binding axis.
        double rawWidthPt = groupBounds.Width * scale;
        double rawHeightPt = groupBounds.Height * scale;
        double rawLeftPt = (slideWidthPt - rawWidthPt) / 2;
        double rawTopPt = (slideHeightPt - rawHeightPt) / 2;

        if (!IsRepresentableAtDisplayPrecision(rawLeftPt)
            || !IsRepresentableAtDisplayPrecision(rawTopPt)
            || !IsRepresentableAtDisplayPrecision(rawWidthPt)
            || !IsRepresentableAtDisplayPrecision(rawHeightPt))
        {
            return Refuse(SlideFitRefusal.InvalidScale);
        }

        double leftPt = GeometryMath.SnapToDisplayPrecision(rawLeftPt);
        double topPt = GeometryMath.SnapToDisplayPrecision(rawTopPt);
        double widthPt = GeometryMath.SnapToDisplayPrecision(rawWidthPt);
        double heightPt = GeometryMath.SnapToDisplayPrecision(rawHeightPt);

        // A legitimate size can still collapse to zero at display precision; a zero-size shape
        // is the degenerate emission CR-9 forbids, so it is a typed refusal.
        bool collapsed = !IsFiniteAndPositive(widthPt) || !IsFiniteAndPositive(heightPt);

        return collapsed
            ? Refuse(SlideFitRefusal.InvalidScale)
            : new SlideFitOutcome(
                Succeeded: true,
                Placement: new SlidePlacement(leftPt, topPt, scale, widthPt, heightPt),
                Refusal: null);
    }

    private static SlideFitOutcome Refuse(SlideFitRefusal refusal) =>
        new(Succeeded: false, Placement: null, Refusal: refusal);

    private static bool IsFiniteAndPositive(double value) => double.IsFinite(value) && value > 0;

    private static bool IsRepresentableAtDisplayPrecision(double value) =>
        double.IsFinite(value) && value >= -float.MaxValue && value <= float.MaxValue;
}
