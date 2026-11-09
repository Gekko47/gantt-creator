namespace GanttCreator.Core.Scene;

/// <summary>
/// Identifies the primitives whose vertical extent is owned by the Excel host rather than by the
/// scene (ADR-0052 D6). A host-owned primitive spans the plot's full vertical extent, so Excel's
/// <c>xlMoveAndSize</c> placement can resize it live on collapse/expand without a refresh.
/// </summary>
/// <remarks>
/// The class is defined from scene geometry alone, with no schema field and no per-type list, so
/// the predicate stays pure and the decision is testable in Core. A horizontal line is never in the
/// class: it is row-anchored and <c>xlMove</c> translation is already what it needs.
/// </remarks>
public static class HostOwnedGeometry
{
    /// <summary>
    /// The tolerance, in points, applied to a host-owned shape's height when verifying that the live
    /// shape still matches the scene (ADR-0052 D8). Excel resizes <c>xlMoveAndSize</c> shapes
    /// proportionally, so repeated collapse/expand cycles can drift by a fraction of a point; a
    /// difference inside this tolerance is a match and no write is issued, a difference outside it
    /// is an update because the scene's measured height is the authority again.
    /// </summary>
    /// <remarks>
    /// Stated, not tuned: 0.5 pt is above Excel's proportional-resize rounding and below any height
    /// a real row change produces. It is wider than <see cref="GeometryMatchTolerancePt"/> because
    /// it covers host rounding, not floating-point representation.
    /// </remarks>
    public const double HostResizeTolerancePt = 0.5;

    /// <summary>
    /// The tolerance, in points, applied to the position and non-host-owned dimensions when
    /// verifying that a live shape still matches the scene (ADR-0052). Above <c>float</c> rounding,
    /// below display resolution.
    /// </summary>
    public const double GeometryMatchTolerancePt = 0.01;

    /// <summary>
    /// Returns <see langword="true"/> when <paramref name="primitive"/> spans the plot's full
    /// vertical extent and is therefore host-owned.
    /// </summary>
    /// <param name="primitive">The scene primitive to classify.</param>
    /// <param name="plotBounds">The plot bounds the primitive is measured against.</param>
    /// <returns>
    /// <see langword="true"/> for a vertical <see cref="SceneLine"/> or a <see cref="SceneRect"/>
    /// whose vertical extent <em>covers</em> the plot's; otherwise <see langword="false"/>.
    /// </returns>
    /// <remarks>
    /// <para>
    /// <b>Coverage, not equality, and the difference is not cosmetic.</b> The builders never
    /// emit plot-bounds geometry for these primitives:
    /// <see cref="PlotSpanGeometry.LiftTopIntoHeader"/> moves a band's top edge up into the
    /// header row by <c>PlotBandHeaderOverlapPt</c> and grows its height (ADR-0037), and
    /// <see cref="PlotSpanGeometry.ClosingLineBottomPt"/> extends a delineator's bottom through
    /// the reserved anchor row by <c>anchorRow + boundaryWidth/2</c> (ADR-0038). A predicate
    /// that demanded equality with <see cref="RectD.Top"/> and <see cref="RectD.Bottom"/>
    /// therefore rejected every real band and every real delineator and left them on
    /// <c>xlMove</c> — the classification silently did nothing, which no amount of idealised
    /// bounds in a unit test could reveal.
    /// </para>
    /// <para>
    /// Containment also keeps the class honest in the other direction: a bar occupies one lane,
    /// a label its measured text height, and a horizontal rule has no vertical extent at all, so
    /// none of them can cover the plot however the rows are sized.
    /// </para>
    /// </remarks>
    public static bool IsHostOwned(ScenePrimitive primitive, RectD plotBounds)
    {
        ArgumentNullException.ThrowIfNull(primitive);

        return primitive switch
        {
            SceneLine line => IsFullHeightLine(line, plotBounds),
            SceneRect rect => IsFullHeightRect(rect, plotBounds),
            _ => false,
        };
    }

    private static bool IsFullHeightLine(SceneLine line, RectD plotBounds) =>
        NearlyEqual(line.From.X, line.To.X)
        && Covers(
            Math.Min(line.From.Y, line.To.Y),
            Math.Max(line.From.Y, line.To.Y),
            plotBounds);

    private static bool IsFullHeightRect(SceneRect rect, RectD plotBounds) =>
        Covers(rect.Bounds.Top, rect.Bounds.Bottom, plotBounds);

    /// <summary>
    /// Whether a vertical interval spans at least the plot's own top-to-bottom extent.
    /// </summary>
    /// <param name="topPt">The interval's top edge.</param>
    /// <param name="bottomPt">The interval's bottom edge.</param>
    /// <param name="plotBounds">The plot bounds.</param>
    /// <returns><see langword="true"/> when the interval contains the plot's extent.</returns>
    private static bool Covers(double topPt, double bottomPt, RectD plotBounds) =>
        topPt <= plotBounds.Top + GeometryMatchTolerancePt
        && bottomPt >= plotBounds.Bottom - GeometryMatchTolerancePt;

    private static bool NearlyEqual(double left, double right) =>
        Math.Abs(left - right) <= GeometryMatchTolerancePt;
}
