namespace GanttCreator.Core.Export;

/// <summary>
/// The single placement the R7.4 fit computed for the pasted group: absolute slide-point
/// coordinates, one uniform scale, and the resulting size (work item D2, PD-2).
/// </summary>
/// <remarks>
/// <para>
/// <b>The data is immutable; the once-latch is a consumption guard.</b> The four geometry
/// members and the scale are get-only and validated in the constructor. The single-use latch is
/// behavioural state, not placement data: it exists so the once-rule (work item D1) is
/// structural — a second application is refused even when no host call raised. This is the audit
/// A-3 reconciliation: the members cannot drift, and the latch cannot be confused for a mutable
/// member.
/// </para>
/// <para>
/// <b>Ratio preserved by construction.</b> <see cref="Width"/> and <see cref="Height"/> derive
/// from the same <see cref="Scale"/>, so the applied aspect equals the group's aspect.
/// </para>
/// <para>
/// <b>Display precision.</b> The four geometry members are already snapped through
/// <see cref="GeometryMath.SnapToDisplayPrecision"/>, so writing them to PowerPoint's
/// <c>Single</c>-typed properties is exact and the read-back comparison is honest.
/// </para>
/// </remarks>
public sealed class SlidePlacement
{
    private int _consumed;

    /// <summary>Creates a placement.</summary>
    /// <param name="leftPt">The absolute slide-point left edge.</param>
    /// <param name="topPt">The absolute slide-point top edge.</param>
    /// <param name="scale">The uniform scale applied to the group bounds.</param>
    /// <param name="widthPt">The applied width in points.</param>
    /// <param name="heightPt">The applied height in points.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when a coordinate is not finite, or the scale, width or height is not finite and
    /// strictly positive. The guards are deliberate: a degenerate placement must never reach the
    /// host (CR-9). Prefer <see cref="SlideFit.Fit"/>, which returns a typed refusal instead.
    /// </exception>
    public SlidePlacement(double leftPt, double topPt, double scale, double widthPt, double heightPt)
    {
        if (!double.IsFinite(leftPt))
        {
            throw new ArgumentOutOfRangeException(
                nameof(leftPt),
                leftPt,
                "The left edge must be finite.");
        }

        if (!double.IsFinite(topPt))
        {
            throw new ArgumentOutOfRangeException(
                nameof(topPt),
                topPt,
                "The top edge must be finite.");
        }

        if (!double.IsFinite(scale) || scale <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(scale),
                scale,
                "The scale must be finite and strictly positive.");
        }

        if (!double.IsFinite(widthPt) || widthPt <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(widthPt),
                widthPt,
                "The width must be finite and strictly positive.");
        }

        if (!double.IsFinite(heightPt) || heightPt <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(heightPt),
                heightPt,
                "The height must be finite and strictly positive.");
        }

        Left = leftPt;
        Top = topPt;
        Scale = scale;
        Width = widthPt;
        Height = heightPt;
    }

    /// <summary>Gets the absolute slide-point left edge.</summary>
    public double Left { get; }

    /// <summary>Gets the absolute slide-point top edge.</summary>
    public double Top { get; }

    /// <summary>Gets the uniform scale applied to the group bounds.</summary>
    public double Scale { get; }

    /// <summary>Gets the applied width in points.</summary>
    public double Width { get; }

    /// <summary>Gets the applied height in points.</summary>
    public double Height { get; }

    /// <summary>Gets whether this placement has been consumed by an application attempt.</summary>
    public bool IsConsumed => Volatile.Read(ref _consumed) != 0;

    /// <summary>
    /// Attempts to consume the placement exactly once. A second call is refused, which is the
    /// work item D1 once-rule made structural.
    /// </summary>
    /// <returns><see langword="true"/> on the first call; <see langword="false"/> on every later call.</returns>
    /// <remarks>
    /// The placement is consumed when an application attempt <em>starts</em>, so a failed attempt
    /// still exhausts it. R7.5 owns retry policy; a retry re-runs the fit rather than re-applying
    /// the same placement.
    /// </remarks>
    public bool TryConsume() => Interlocked.CompareExchange(ref _consumed, 1, 0) == 0;
}
