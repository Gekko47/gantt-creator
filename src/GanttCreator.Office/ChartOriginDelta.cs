using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// The uniform translation a renderer applies to move a scene whose chart bounds
/// may have a negative origin onto the worksheet, where Excel can express it.
/// </summary>
/// <remarks>
/// <para>
/// Entity guide section 1 requires a renderer to translate the chart bounds to a
/// zero origin and carry the delta. The rule was provisional pending the R4.3
/// Step-0 probe, and that probe is why this type is load-bearing rather than
/// cosmetic: a live host <strong>silently clamps</strong> a negative shape
/// <c>Left</c>/<c>Top</c> to zero instead of rejecting it. A renderer that passed
/// scene coordinates straight through would therefore produce a visibly wrong
/// chart with no error anywhere to catch it.
/// </para>
/// <para>
/// Negative coordinates remain legal in the scene - <see cref="RectD"/> permits
/// them deliberately - and are legal nowhere in the host. The translation is
/// uniform across every primitive kind: a rect's bounds, a line's
/// <c>From</c>/<c>To</c>, a polygon's points, and a text's bounds all shift by
/// this same constant, so no relationship inside the scene changes, including
/// whether a label sits outside the chart.
/// </para>
/// <para>
/// The delta is also carried into export bounds (R6.1). Because empty
/// whitespace outside the bounds is not exported, dropping it on the way to a
/// crop would silently remove exactly <c>ChartOuterPaddingPt</c> of margin.
/// </para>
/// </remarks>
public readonly record struct ChartOriginDelta
{
    /// <summary>Initialises a translation from its two constant components.</summary>
    /// <param name="dx">The horizontal shift in points.</param>
    /// <param name="dy">The vertical shift in points.</param>
    public ChartOriginDelta(double dx, double dy)
    {
        if (!double.IsFinite(dx) || !double.IsFinite(dy))
        {
            throw new ArgumentOutOfRangeException(
                nameof(dx),
                double.IsFinite(dx) ? dy : dx,
                "A chart origin delta must be finite on both axes.");
        }

        Dx = dx;
        Dy = dy;
    }

    /// <summary>Gets the horizontal shift in points.</summary>
    public double Dx { get; }

    /// <summary>Gets the vertical shift in points.</summary>
    public double Dy { get; }

    /// <summary>The translation that leaves every coordinate unchanged.</summary>
    public static ChartOriginDelta Identity => new(0d, 0d);

    /// <summary>
    /// Computes the translation that moves <paramref name="chartBounds"/> onto
    /// the worksheet origin.
    /// </summary>
    /// <param name="chartBounds">The scene's chart bounds, which may be negative.</param>
    /// <returns>The delta that makes the bounds' top-left corner the origin.</returns>
    /// <remarks>
    /// The negation is the whole computation. Because the chart bounds are
    /// derived by expanding the content union by <c>ChartOuterPaddingPt</c>, a
    /// chart whose content starts at the origin has a negative bounds origin and
    /// a positive delta.
    /// </remarks>
    public static ChartOriginDelta ForChartBounds(RectD chartBounds) =>
        new(-chartBounds.X, -chartBounds.Y);

    /// <summary>Applies this translation to a point.</summary>
    /// <param name="point">The scene point.</param>
    /// <returns>The worksheet point.</returns>
    public PointD Apply(PointD point) => point.Offset(Dx, Dy);

    /// <summary>Applies this translation to a rectangle's origin, leaving its extents unchanged.</summary>
    /// <param name="rect">The scene rectangle.</param>
    /// <returns>The worksheet rectangle.</returns>
    public RectD Apply(RectD rect) => new(rect.X + Dx, rect.Y + Dy, rect.Width, rect.Height);
}
