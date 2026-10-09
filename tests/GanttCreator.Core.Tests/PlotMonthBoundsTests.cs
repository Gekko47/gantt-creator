namespace GanttCreator.Core.Tests;

/// <summary>
/// Contract tests for <see cref="PlotMonthBounds"/>: the single authority
/// for the chart's month-snapped plot bounds shared by the Refresh render
/// pipeline and the Ribbon's AUTO display derivation.
/// </summary>
public sealed class PlotMonthBoundsTests
{
    [Fact]
    public void SnapExtent_snaps_outward_to_whole_months()
    {
        // 10 Jan pads inside January and snaps back to 01/01; 20 Jun snaps
        // out to the end of June.
        (DateOnly start, DateOnly finish) = PlotMonthBounds.SnapExtent(
            new DateOnly(2026, 1, 10),
            new DateOnly(2026, 6, 20));

        Assert.Equal(new DateOnly(2026, 1, 1), start);
        Assert.Equal(new DateOnly(2026, 6, 30), finish);
    }

    [Fact]
    public void SnapExtent_snaps_a_same_month_extent_to_the_full_month()
    {
        // Jan 1 and Jan 10 both land in January, so the chart spans the
        // whole month. (The degenerate widen inside SnapExtent guards
        // inverted inputs only; month-snap outputs always satisfy
        // month-start < month-end.)
        (DateOnly start, DateOnly finish) = PlotMonthBounds.SnapExtent(
            new DateOnly(2026, 1, 1),
            new DateOnly(2026, 1, 10));

        Assert.Equal(new DateOnly(2026, 1, 1), start);
        Assert.Equal(new DateOnly(2026, 1, 31), finish);
    }

    [Fact]
    public void SnapToMonthEnd_handles_december_without_overflow() =>
        Assert.Equal(new DateOnly(2026, 12, 31), PlotMonthBounds.SnapToMonthEnd(new DateOnly(2026, 12, 15)));

    [Fact]
    public void SnapToMonthEnd_handles_february_in_a_leap_year()
    {
        Assert.Equal(new DateOnly(2024, 2, 29), PlotMonthBounds.SnapToMonthEnd(new DateOnly(2024, 2, 10)));
        Assert.Equal(new DateOnly(2026, 2, 28), PlotMonthBounds.SnapToMonthEnd(new DateOnly(2026, 2, 10)));
    }
}
