namespace GanttCreator.Core.Scene;

/// <summary>
/// Identifies the primitives whose extent is owned by the Excel host rather than by the scene
/// (ADR-0052 D6, amended 2026-11-13). A host-owned primitive spans the plot's full vertical extent,
/// or the plot's full width as a horizontal rule, so Excel's <c>xlMoveAndSize</c> placement can
/// resize it live on a collapse/expand or a column resize without a refresh.
/// </summary>
/// <remarks>
/// <para>
/// The class is defined from scene geometry alone, with no schema field and no per-type list, so
/// the predicate stays pure and the decision is testable in Core.
/// </para>
/// <para>
/// <b>Horizontal rules are in the class (owner ruling 2026-11-13).</b> This paragraph used to say
/// the opposite — "a horizontal line is never in the class: it is row-anchored and
/// <c>xlMove</c> translation is already what it needs" — and the live chart proved it wrong for one
/// case. <c>xlMove</c> translates a shape and never resizes it, so the chart's upper and lower frame
/// rules stayed at their old width when a worksheet column was widened, and Excel offers no
/// two-edge anchor that could stretch them another way. A rule spanning the plot's width is
/// therefore host-sized exactly like a band spanning its height. A bar-length rule, a
/// label-width rule, and a rule that stops short of the plot are all still excluded: they are
/// row-anchored and have nothing to stretch.
/// </para>
/// </remarks>
public static class HostOwnedGeometry
{
    /// <summary>
    /// The tolerance, in points, applied to a host-owned shape's host-resized extent when verifying
    /// that the live shape still matches the scene (ADR-0052 D8). Excel resizes
    /// <c>xlMoveAndSize</c> shapes proportionally, so repeated collapse/expand cycles - and a column
    /// resize, for a horizontal rule - can drift by a fraction of a point; a difference inside this
    /// tolerance is a match and no write is issued, a difference outside it is an update because the
    /// scene's measured extent is the authority again.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Stated, not tuned: 0.5 pt is above Excel's proportional-resize rounding and below any height
    /// a real row change produces. It is wider than <see cref="GeometryMatchTolerancePt"/> because
    /// it covers host rounding, not floating-point representation.
    /// </para>
    /// <para>
    /// It applies to the dimension the host actually stretches: the <em>height</em> of a band or a
    /// vertical delineator, and the <em>width</em> of a horizontal rule. The other dimension keeps
    /// the strict tolerance, because the host has no reason to have moved it and a mismatch there is
    /// a real difference.
    /// </para>
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
    /// vertical extent, or its full width as a horizontal rule, and is therefore host-owned.
    /// </summary>
    /// <param name="primitive">The scene primitive to classify.</param>
    /// <param name="plotBounds">The plot bounds the primitive is measured against.</param>
    /// <returns>
    /// <see langword="true"/> for a vertical <see cref="SceneLine"/>, a horizontal
    /// <see cref="SceneLine"/> whose horizontal extent <em>covers</em> the plot's, or a
    /// <see cref="SceneRect"/> whose vertical extent covers the plot's; otherwise
    /// <see langword="false"/>.
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
            SceneLine line => IsFullHeightLine(line, plotBounds) || IsFullWidthLine(line, plotBounds),
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

    /// <summary>
    /// Whether a horizontal line spans at least the plot's own left-to-right extent, which is the
    /// horizontal-rule member of the class.
    /// </summary>
    /// <param name="line">The line to classify.</param>
    /// <param name="plotBounds">The plot bounds the line is measured against.</param>
    /// <returns><see langword="true"/> when the line is horizontal and covers the plot's width.</returns>
    /// <remarks>
    /// Coverage, not equality, for the same reason the vertical rule uses it: the frame rules are
    /// built on the content bounds, which include the panel and can be wider than the plot, so a
    /// predicate demanding equality would classify nothing and the carve-out would be inert again.
    /// A near-zero-height line cannot cover the plot VERTICALLY, so the two tests cannot both fire
    /// on one line.
    /// </remarks>
    private static bool IsFullWidthLine(SceneLine line, RectD plotBounds) =>
        NearlyEqual(line.From.Y, line.To.Y)
        && CoversHorizontally(
            Math.Min(line.From.X, line.To.X),
            Math.Max(line.From.X, line.To.X),
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

    /// <summary>
    /// Whether a horizontal interval spans at least the plot's own left-to-right extent.
    /// </summary>
    /// <param name="leftPt">The interval's left edge.</param>
    /// <param name="rightPt">The interval's right edge.</param>
    /// <param name="plotBounds">The plot bounds.</param>
    /// <returns><see langword="true"/> when the interval contains the plot's width.</returns>
    private static bool CoversHorizontally(double leftPt, double rightPt, RectD plotBounds) =>
        leftPt <= plotBounds.Left + GeometryMatchTolerancePt
        && rightPt >= plotBounds.Right - GeometryMatchTolerancePt;

    private static bool NearlyEqual(double left, double right) =>
        Math.Abs(left - right) <= GeometryMatchTolerancePt;
}
