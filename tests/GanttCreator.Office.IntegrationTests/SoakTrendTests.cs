namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Host-free positives for the R7.7 D1 trend evaluators. No Category trait: these run in
/// the quick gate alongside the other non-live Office harness tests.
/// </summary>
public sealed class SoakTrendTests
{
    [Fact]
    public void A_flat_handle_series_is_within_tolerance()
    {
        SoakTrendOutcome outcome = SoakTrend.Handles([100, 101, 100, 102, 99, 100, 101], tolerance: 5);

        Assert.True(outcome.WithinTolerance);
        Assert.Equal(99, outcome.Baseline);
        Assert.Equal(102, outcome.End);
        Assert.Equal(3, outcome.Delta);
        Assert.True(outcome.Baseline <= outcome.End);
    }

    [Fact]
    public void A_monotonically_growing_handle_series_exceeds_tolerance()
    {
        // The measured span is the steady region's, so it reads 6 rather than the whole
        // series' 10: the first two samples are declared warm-up. The verdict is unchanged,
        // because growth through the steady region is what a leak looks like.
        SoakTrendOutcome outcome = SoakTrend.Handles([100, 102, 104, 106, 108, 110], tolerance: 5);

        Assert.False(outcome.WithinTolerance);
        Assert.Equal(6, outcome.Delta);
        Assert.True(outcome.Delta > 5);
    }

    [Fact]
    public void The_measured_attach_series_passes_once_its_warm_up_is_excluded()
    {
        // The series the live gate recorded on 2026-11-25, which failed the whole-series rule:
        // 1892 on the first pass, then an allocation step to 1951/1956, then a steady
        // 1920-1956. Two warm-up samples make the gated region span 36 against a tolerance
        // of 50, which is the measurement the tolerance was always meant to describe.
        SoakTrendOutcome outcome = SoakTrend.Handles(
            [1892, 1951, 1956, 1949, 1937, 1941, 1939, 1938, 1940, 1938, 1939, 1940, 1936, 1920, 1922],
            tolerance: 50);

        Assert.True(outcome.WithinTolerance);
        Assert.Equal(1920, outcome.Baseline);
        Assert.Equal(1956, outcome.End);
        Assert.Equal(36, outcome.Delta);
        Assert.Equal(2, SoakTrend.HandleWarmUpSamples);
    }

    [Fact]
    public void A_leak_that_starts_inside_the_warm_up_still_fails_the_trend()
    {
        // A climb from the very first sample cannot hide in the excluded warm-up: the steady
        // region keeps climbing, so the gated span still exceeds the tolerance.
        SoakTrendOutcome outcome = SoakTrend.Handles([100, 130, 160, 190, 220, 250], tolerance: 50);

        Assert.False(outcome.WithinTolerance);
        Assert.Equal(90, outcome.Delta);
    }

    [Fact]
    public void A_steady_region_too_short_to_prove_anything_cannot_pass()
    {
        // Fail closed, like the empty series: excluding the warm-up must not shrink the
        // measurement to one or two samples and call that flat.
        SoakTrendOutcome outcome = SoakTrend.Handles([100, 100, 100, 100], tolerance: int.MaxValue);

        Assert.False(outcome.WithinTolerance);

        SoakTrendOutcome tooShort = SoakTrend.Handles([100, 100], tolerance: int.MaxValue);

        Assert.False(tooShort.WithinTolerance);
    }

    [Fact]
    public void A_single_handle_spike_outside_tolerance_fails_the_trend()
    {
        // max - min is the gate: one 50-handle spike is a real event, not noise, until the
        // measured reference series says otherwise.
        SoakTrendOutcome outcome = SoakTrend.Handles([100, 100, 150, 100, 100], tolerance: 20);

        Assert.False(outcome.WithinTolerance);
        Assert.Equal(50, outcome.Delta);
    }

    [Fact]
    public void An_empty_handle_series_cannot_prove_a_flat_trend()
    {
        // Fail closed: an unmeasured trend is not a flat one (the exit-5 discipline).
        SoakTrendOutcome outcome = SoakTrend.Handles([], tolerance: 10);

        Assert.False(outcome.WithinTolerance);
    }

    [Fact]
    public void A_working_set_that_returns_to_baseline_is_within_tolerance_despite_a_mid_run_spike()
    {
        // First third [100, 100] and last third [100, 100] both median to 100, so the
        // mid-series 500 spike cannot trip the robust statistic.
        SoakTrendOutcome outcome = SoakTrend.WorkingSet([100, 100, 500, 500, 100, 100], toleranceBytes: 10);

        Assert.True(outcome.WithinTolerance);
        Assert.Equal(100, outcome.Baseline);
        Assert.Equal(100, outcome.End);
        Assert.Equal(0, outcome.Delta);
    }

    [Fact]
    public void A_monotonically_growing_working_set_exceeds_tolerance()
    {
        // First third [100, 110] medians to 105; last third [140, 150] medians to 145.
        SoakTrendOutcome outcome = SoakTrend.WorkingSet([100, 110, 120, 130, 140, 150], toleranceBytes: 30);

        Assert.False(outcome.WithinTolerance);
        Assert.Equal(105, outcome.Baseline);
        Assert.Equal(145, outcome.End);
        Assert.Equal(40, outcome.Delta);
    }

    [Fact]
    public void A_short_working_set_series_cannot_prove_a_flat_trend()
    {
        SoakTrendOutcome outcome = SoakTrend.WorkingSet([100, 100], toleranceBytes: long.MaxValue);

        Assert.False(outcome.WithinTolerance);
    }
}
