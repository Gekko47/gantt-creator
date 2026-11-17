namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R5.1a: the refresh outcome carries the inclusive plot range the successful pass
/// rendered, so the ribbon can repaint an automatic end's displayed date from the
/// scene it just drew instead of re-reading the table.
/// </summary>
public sealed class GanttRefreshOutcomePlotRangeTests
{
    [Fact]
    public void A_successful_outcome_carries_the_range_the_scene_used()
    {
        GanttRefreshOutcome outcome = GanttRefreshOutcome.Ok(
            [],
            shapesWritten: 1,
            durationCellsWritten: 0,
            styleKeyCellsWritten: 0,
            plotStart: new DateOnly(2026, 3, 1),
            plotFinish: new DateOnly(2026, 4, 30));

        Assert.Equal(new DateOnly(2026, 3, 1), outcome.PlotStart);
        Assert.Equal(new DateOnly(2026, 4, 30), outcome.PlotFinish);
    }

    [Fact]
    public void An_outcome_without_a_range_reports_nulls()
    {
        // The optional parameters keep every existing caller compiling; no range
        // means "unknown", which the ribbon treats as "leave the text alone".
        GanttRefreshOutcome outcome = GanttRefreshOutcome.Ok(
            [],
            shapesWritten: 0,
            durationCellsWritten: 0,
            styleKeyCellsWritten: 0);

        Assert.Null(outcome.PlotStart);
        Assert.Null(outcome.PlotFinish);
    }

    [Fact]
    public void A_refused_outcome_carries_no_range()
    {
        GanttRefreshOutcome outcome = GanttRefreshOutcome.Refused(
            GanttRefreshRefusal.PartialReconciliation,
            "the chart could not be refreshed");

        Assert.Null(outcome.PlotStart);
        Assert.Null(outcome.PlotFinish);
    }
}
