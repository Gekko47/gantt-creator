using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the deterministic text-metrics seam Core ships. Every label, band, and
/// truncation decision in the scene depends on this type reporting a finite,
/// non-negative width or failing honestly, so its failure arms are part of the
/// contract rather than defensive padding.
/// </summary>
public sealed class FakeTextMetricsTests
{
    [Fact]
    public void The_default_table_measures_four_points_per_character()
    {
        FakeTextMetrics metrics = new();

        Assert.True(metrics.TryMeasure("abcde", out TextMeasurement? measurement));
        Assert.Equal(20.0, measurement!.WidthPt);
    }

    [Fact]
    public void A_null_text_is_a_failed_measurement()
    {
        FakeTextMetrics metrics = new();

        Assert.False(metrics.TryMeasure(null!, out TextMeasurement? measurement));
        Assert.Null(measurement);
    }

    [Fact]
    public void A_non_finite_or_negative_advance_is_a_failed_measurement()
    {
        // A seam that reports a bad advance must fail the measurement rather than
        // contribute a zero width: a zero would silently "fit" text into no space and
        // the label would be drawn over whatever it was supposed to avoid.
        FakeTextMetrics infinite = new(_ => double.PositiveInfinity, 10.0);
        FakeTextMetrics negative = new(_ => -1.0, 10.0);

        Assert.False(infinite.TryMeasure("ab", out TextMeasurement? infiniteMeasurement));
        Assert.Null(infiniteMeasurement);
        Assert.False(negative.TryMeasure("ab", out TextMeasurement? negativeMeasurement));
        Assert.Null(negativeMeasurement);
    }

    [Fact]
    public void The_seam_refuses_a_line_height_it_could_not_report()
    {
        Assert.Throws<ArgumentNullException>(() => new FakeTextMetrics(null!, 10.0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new FakeTextMetrics(_ => 4.0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new FakeTextMetrics(_ => 4.0, double.NaN));
    }
}
