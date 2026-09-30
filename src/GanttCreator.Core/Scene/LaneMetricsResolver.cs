namespace GanttCreator.Core.Scene;

/// <summary>
/// The result of resolving measured worksheet row heights into lane metrics
/// (R4.7D, ADR-0026 D2 and D5).
/// </summary>
/// <param name="Metrics">The resolved metrics, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record LaneMetricsResolution(LaneLayoutMetrics? Metrics, LaneMetricsRefusal? Refusal)
{
    /// <summary>Gets whether metrics were resolved.</summary>
    public bool Succeeded => Metrics is not null;
}

/// <summary>Why measured row heights could not become lane metrics.</summary>
public enum LaneMetricsRefusal
{
    /// <summary>The measured grid was null.</summary>
    NullGrid = 0,

    /// <summary>The grid carried no body row heights.</summary>
    NoRows = 1,

    /// <summary>A measured row height was non-finite or non-positive.</summary>
    InvalidRowHeight = 2,
}

/// <summary>
/// Turns the <em>measured</em> worksheet row heights into the lane metrics, so a
/// lane's height comes from the worksheet rather than from its content.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is a separate step.</b> R4.7D removed lane auto-growth, which means
/// the lane height is no longer something Core can choose: it is the managed row
/// height. Core cannot measure a worksheet, so the measurement arrives as a
/// <see cref="PanelCellGrid"/> from the Office layer, and this turns it into the
/// <see cref="LaneLayoutMetrics"/> that <c>LaneLayoutBuilder</c> consumes. The
/// decision stays in Core and is therefore testable without a live host, while
/// Excel remains the only thing that measures.
/// </para>
/// <para>
/// <b>Which row height wins.</b> Managed rows are normalised to a single
/// <c>GanttRowHeightPt</c>, so in a correctly normalised worksheet every measured
/// height is that same value. A worksheet the user has dragged is therefore
/// measured as-is, and the <em>most common</em> height is used — with equal
/// counts resolved by the first height in worksheet order, so the result is
/// deterministic rather than dependent on enumeration.
/// </para>
/// <para>
/// <b>Why the most common and not the maximum.</b> A single dragged row must not
/// change every lane in the chart, and the minimum would let one compressed row
/// shrink them all. The mode is the single value the worksheet is actually
/// expressing.
/// </para>
/// <para>
/// <b>Structural rows keep their own heights.</b> <c>SplitterHeightPt</c> and
/// <c>SpacerHeightPt</c> are separate tokens and are not derived from the measured
/// body, because a structural row is not a normal managed row.
/// </para>
/// </remarks>
public static class LaneMetricsResolver
{
    /// <summary>
    /// Resolves lane metrics from a measured grid and the other lane tokens.
    /// </summary>
    /// <param name="grid">The measured grid, from the Office measurement port.</param>
    /// <param name="lanePaddingTopPt">Space above the first stack slot.</param>
    /// <param name="lanePaddingBottomPt">Space below the last stack slot.</param>
    /// <param name="stackGapPt">Space between distinct effective slots.</param>
    /// <param name="splitterHeightPt">The fixed splitter lane height.</param>
    /// <param name="spacerHeightPt">The fixed spacer lane height.</param>
    /// <returns>The resolved metrics, or a typed refusal.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="grid"/> is null.</exception>
    public static LaneMetricsResolution Resolve(
        PanelCellGrid? grid,
        double lanePaddingTopPt,
        double lanePaddingBottomPt,
        double stackGapPt,
        double splitterHeightPt,
        double spacerHeightPt)
    {
        if (grid is null)
        {
            return new LaneMetricsResolution(null, LaneMetricsRefusal.NullGrid);
        }

        IReadOnlyList<double> heights = grid.RowHeightsPt;
        if (heights.Count == 0)
        {
            return new LaneMetricsResolution(null, LaneMetricsRefusal.NoRows);
        }

        foreach (var height in heights)
        {
            if (!double.IsFinite(height) || height <= 0)
            {
                return new LaneMetricsResolution(null, LaneMetricsRefusal.InvalidRowHeight);
            }
        }

        var laneHeightPt = MostCommon(heights);

        return new LaneMetricsResolution(
            new LaneLayoutMetrics(
                laneHeightPt,
                lanePaddingTopPt,
                lanePaddingBottomPt,
                stackGapPt,
                splitterHeightPt,
                spacerHeightPt),
            null);
    }

    /// <summary>
    /// The most frequent height, ties broken by the first occurrence in worksheet
    /// order so the result never depends on enumeration order.
    /// </summary>
    private static double MostCommon(IReadOnlyList<double> heights)
    {
        Dictionary<double, int> counts = [];
        foreach (var height in heights)
        {
            _ = counts.TryGetValue(height, out var seen);
            counts[height] = seen + 1;
        }

        var best = heights[0];
        var bestCount = 0;
        foreach (var height in heights)
        {
            if (counts[height] > bestCount)
            {
                best = height;
                bestCount = counts[height];
            }
        }

        return best;
    }
}
