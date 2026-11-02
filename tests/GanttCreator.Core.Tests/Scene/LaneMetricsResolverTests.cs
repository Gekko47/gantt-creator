using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the measured-grid to lane-metrics resolution, including every refusal.
/// The refusals are the contract that keeps a chart from being drawn from a grid it
/// cannot trust: no grid, no rows, a negative height, or a body whose rows are all
/// hidden.
/// </summary>
public sealed class LaneMetricsResolverTests
{
    private static LaneMetricsResolution Resolve(PanelCellGrid? grid) =>
        LaneMetricsResolver.Resolve(grid, lanePaddingTopPt: 1.0, lanePaddingBottomPt: 2.0, stackGapPt: 3.0, splitterHeightPt: 20.0, spacerHeightPt: 10.0);

    private static PanelCellGrid Grid(IReadOnlyList<double> rowHeightsPt, bool allowEmptyRows = false) =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Type", 100)],
            rowHeightsPt,
            headerHeightPt: 18.0,
            requiredColumns: ["Type"],
            allowEmptyRows: allowEmptyRows).Grid!;

    [Fact]
    public void A_null_grid_is_refused()
    {
        Assert.Equal(LaneMetricsRefusal.NullGrid, Resolve(null).Refusal);
    }

    [Fact]
    public void A_grid_with_no_rows_is_refused()
    {
        Assert.Equal(LaneMetricsRefusal.NoRows, Resolve(Grid([], allowEmptyRows: true)).Refusal);
    }

    [Fact]
    public void A_negative_row_height_never_reaches_the_resolver()
    {
        // PanelCellGrid already refuses a negative row height, so the resolver's own
        // InvalidRowHeight arm cannot fire through the public grid type. Pinned here so
        // the redundancy is visible instead of turning up later as a coverage gap.
        Assert.Equal(
            PanelCellGridRefusal.NegativeRowHeight,
            PanelCellGrid.TryCreate(
                [new PanelColumn("Type", 100)],
                [10.0, -1.0],
                headerHeightPt: 18.0,
                requiredColumns: ["Type"]).Refusal);
    }

    [Fact]
    public void A_body_of_only_hidden_rows_is_refused()
    {
        // A hidden row measures zero. All-zero means no visible row to derive a lane
        // from, which is a refusal rather than a zero-height chart nobody explains.
        Assert.Equal(LaneMetricsRefusal.NoVisibleRows, Resolve(Grid([0.0, 0.0])).Refusal);
    }

    [Fact]
    public void A_visible_row_resolves_the_metrics()
    {
        // Hidden rows keep their place in the grid but contribute no lane, so the
        // resolver must skip them and still succeed on the visible remainder.
        LaneMetricsResolution outcome = Resolve(Grid([0.0, 12.0, 12.0, 0.0]));

        Assert.Null(outcome.Refusal);
        Assert.NotNull(outcome.Metrics);
    }
}
