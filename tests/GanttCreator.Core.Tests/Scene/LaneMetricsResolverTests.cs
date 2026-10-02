using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for <see cref="LaneMetricsResolver"/>: the lane height is the MEASURED
/// worksheet row height, not a value Core chooses from content (R4.7D,
/// ADR-0026 D2).
/// </summary>
public sealed class LaneMetricsResolverTests
{
    private const double PaddingTop = 3;
    private const double PaddingBottom = 3;
    private const double StackGap = 2;
    private const double Splitter = 18;
    private const double Spacer = 9;

    private static PanelCellGrid Grid(params double[] heightsPt) =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Id", 40), new PanelColumn("Description", 160)],
            heightsPt,
            10,
            ["Id", "Description"]).Grid!;

    private static LaneMetricsResolution Resolve(params double[] heightsPt) =>
        LaneMetricsResolver.Resolve(Grid(heightsPt), PaddingTop, PaddingBottom, StackGap, Splitter, Spacer);

    /// <summary>
    /// The core promise: a uniformly measured worksheet produces a lane height
    /// equal to that measured row height, and Core consumes it rather than
    /// choosing its own.
    /// </summary>
    [Fact]
    public void A_uniform_measurement_becomes_the_lane_height()
    {
        LaneMetricsResolution resolution = Resolve(18, 18, 18);

        Assert.True(resolution.Succeeded);
        Assert.Equal(18, resolution.Metrics!.LaneHeightPt);
    }

    /// <summary>
    /// A non-default measured height is used verbatim. This is the difference
    /// between deriving the lane height and measuring it: a 30pt worksheet row
    /// yields 30pt lanes, not the 18pt the catalogue default would give.
    /// </summary>
    [Fact]
    public void A_measured_height_is_used_verbatim_not_defaulted()
    {
        LaneMetricsResolution resolution = Resolve(30, 30, 30);

        Assert.Equal(30, resolution.Metrics!.LaneHeightPt);
    }

    /// <summary>
    /// The other lane tokens pass through unchanged; only the height is derived
    /// from the measurement.
    /// </summary>
    [Fact]
    public void The_other_lane_tokens_pass_through()
    {
        LaneMetricsResolution resolution = Resolve(18, 18);

        LaneLayoutMetrics metrics = resolution.Metrics!;
        Assert.Equal(PaddingTop, metrics.LanePaddingTopPt);
        Assert.Equal(PaddingBottom, metrics.LanePaddingBottomPt);
        Assert.Equal(StackGap, metrics.StackGapPt);
        Assert.Equal(Splitter, metrics.SplitterHeightPt);
        Assert.Equal(Spacer, metrics.SpacerHeightPt);
    }

    /// <summary>
    /// One dragged row must not change every lane in the chart, so the most common
    /// measured height wins.
    /// </summary>
    [Fact]
    public void The_most_common_measured_height_wins()
    {
        LaneMetricsResolution resolution = Resolve(18, 18, 18, 45);

        Assert.Equal(18, resolution.Metrics!.LaneHeightPt);
    }

    /// <summary>
    /// ...and the minimum must not win either: one compressed row must not shrink
    /// every lane.
    /// </summary>
    [Fact]
    public void A_single_short_row_does_not_shrink_the_lane_height()
    {
        LaneMetricsResolution resolution = Resolve(18, 18, 18, 6);

        Assert.Equal(18, resolution.Metrics!.LaneHeightPt);
    }

    /// <summary>
    /// Determinism: the same measured heights resolve the same way regardless of
    /// the order they are supplied in, so a re-sorted worksheet does not move the
    /// chart.
    /// </summary>
    [Fact]
    public void Resolution_is_deterministic_regardless_of_row_order()
    {
        Assert.Equal(Resolve(18, 18, 45).Metrics!.LaneHeightPt, Resolve(45, 18, 18).Metrics!.LaneHeightPt);
        Assert.Equal(Resolve(18, 18, 45).Metrics!.LaneHeightPt, Resolve(18, 45, 18).Metrics!.LaneHeightPt);
    }

    /// <summary>
    /// Structural rows keep their own tokens rather than inheriting a measured body
    /// height: a structural row is not a normal managed row. Measured at 30pt so the
    /// assertion is not satisfied by accident through the 18pt splitter token.
    /// </summary>
    [Fact]
    public void Structural_heights_are_not_taken_from_the_body()
    {
        LaneMetricsResolution resolution = Resolve(30, 30);

        Assert.Equal(30, resolution.Metrics!.LaneHeightPt);
        Assert.Equal(Splitter, resolution.Metrics.SplitterHeightPt);
        Assert.Equal(Spacer, resolution.Metrics.SpacerHeightPt);
    }

    [Fact]
    public void A_null_grid_is_refused()
    {
        LaneMetricsResolution resolution = LaneMetricsResolver.Resolve(null, PaddingTop, PaddingBottom, StackGap, Splitter, Spacer);

        Assert.False(resolution.Succeeded);
        Assert.Equal(LaneMetricsRefusal.NullGrid, resolution.Refusal);
    }

    /// <summary>
    /// An empty body is refused at the boundary. <see cref="PanelCellGrid"/> rejects
    /// an empty height list first, so the resolver's own <c>NoRows</c> branch is
    /// unreachable through that path -- it exists for defence in depth against a
    /// grid arriving from a caller that does not go through the factory. This test
    /// pins the reachable half honestly rather than asserting a branch it cannot
    /// reach.
    /// </summary>
    [Fact]
    public void An_empty_body_is_refused_by_the_grid()
    {
        PanelCellGridCreationOutcome outcome = PanelCellGrid.TryCreate(
            [new PanelColumn("Id", 40), new PanelColumn("Description", 160)],
            [],
            10,
            ["Id", "Description"]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PanelCellGridRefusal.NoRows, outcome.Refusal);
    }

    /// <summary>
    /// A non-positive measured height is refused at the boundary. A defaulted height
    /// here would silently put the chart out of alignment with the worksheet, which
    /// is the failure this whole row exists to remove.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The refusal is <b>the grid factory's</b>, not the resolver's:
    /// <see cref="PanelCellGrid.TryCreate"/> rejects a non-positive or non-finite
    /// height before a grid ever reaches <see cref="LaneMetricsResolver.Resolve"/>,
    /// so the resolver's own <c>InvalidRowHeight</c> branch is unreachable through
    /// the factory -- exactly the situation the empty-body test above pins honestly
    /// rather than asserting a branch it cannot reach.
    /// </para>
    /// <para>
    /// The previous form of this test called <c>Resolve</c> with a grid built from
    /// the bad height. That grid is <see langword="null"/> (the factory refused it),
    /// so the call reported <c>NullGrid</c> and the test passed without ever
    /// exercising a height check. Asserting the specific refusal code is what makes
    /// the test prove the height is what was rejected.
    /// </para>
    /// </remarks>
    [Theory]
    [InlineData(0d)]
    [InlineData(-5d)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void An_invalid_measured_height_is_refused(double badHeight)
    {
        PanelCellGridCreationOutcome outcome = PanelCellGrid.TryCreate(
            [new PanelColumn("Id", 40), new PanelColumn("Description", 160)],
            [badHeight, 18],
            10,
            ["Id", "Description"]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PanelCellGridRefusal.NonPositiveRowHeight, outcome.Refusal);
    }
}
