using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Export;

/// <summary>
/// R7.4: the pure slide fit. The matrix covers 4:3, 16:9 and custom slides against representative
/// group aspects; every refusal has a positive test that constructs the bad input; the ratio,
/// containment and single-use rules are pinned.
/// </summary>
public sealed class SlideFitTests
{
    /// <summary>The repository's host read-back tolerance (0.01 pt).</summary>
    private const double _tolerancePt = HostOwnedGeometry.GeometryMatchTolerancePt;

    /// <summary>
    /// slideWidth, slideHeight, groupWidth, groupHeight, inset. 4:3 (720x540), 16:9 (960x540)
    /// and custom wide/tall slides against wider, taller, small, near-exact and extreme groups.
    /// </summary>
    public static TheoryData<double, double, double, double, double> SlidesAndGroups => new()
    {
        { 720, 540, 800, 200, 4 },
        { 720, 540, 200, 800, 4 },
        { 720, 540, 100, 50, 4 },
        { 720, 540, 700, 500, 4 },
        { 960, 540, 800, 200, 4 },
        { 960, 540, 200, 800, 4 },
        { 960, 540, 100, 50, 4 },
        { 1200, 400, 2000, 100, 4 },
        { 400, 1200, 100, 2000, 4 },
    };

    [Theory]
    [MemberData(nameof(SlidesAndGroups))]
    public void A_fit_is_within_the_margins_fills_one_axis_and_keeps_the_aspect(
        double slideWidthPt,
        double slideHeightPt,
        double groupWidthPt,
        double groupHeightPt,
        double insetPt)
    {
        SlideFitOutcome outcome = SlideFit.Fit(
            new RectD(0, 0, groupWidthPt, groupHeightPt),
            slideWidthPt,
            slideHeightPt,
            insetPt);

        Assert.True(outcome.Succeeded, $"Refused: {outcome.Refusal}");
        SlidePlacement placement = Assert.IsType<SlidePlacement>(outcome.Placement);

        double fitWidthPt = slideWidthPt - (2 * insetPt);
        double fitHeightPt = slideHeightPt - (2 * insetPt);

        // Inset containment: the whole placement is inside the policy margins.
        Assert.True(placement.Left >= insetPt - _tolerancePt);
        Assert.True(placement.Top >= insetPt - _tolerancePt);
        Assert.True(placement.Left + placement.Width <= slideWidthPt - insetPt + _tolerancePt);
        Assert.True(placement.Top + placement.Height <= slideHeightPt - insetPt + _tolerancePt);

        // The limiting axis fills the frame, which is what "fit" means.
        Assert.True(
            Math.Abs(placement.Width - fitWidthPt) <= _tolerancePt
                || Math.Abs(placement.Height - fitHeightPt) <= _tolerancePt,
            "The fitted placement fills neither axis of the frame.");

        // Ratio preserved by construction (D2): one scale, both axes.
        double expectedAspect = groupWidthPt / groupHeightPt;
        double appliedAspect = placement.Width / placement.Height;
        Assert.True(
            Math.Abs(appliedAspect - expectedAspect) <= _tolerancePt,
            $"Aspect drifted: expected {expectedAspect}, applied {appliedAspect}.");

        // Centred within the frame on both axes.
        double expectedLeftPt = insetPt + ((fitWidthPt - placement.Width) / 2);
        double expectedTopPt = insetPt + ((fitHeightPt - placement.Height) / 2);
        Assert.True(Math.Abs(placement.Left - expectedLeftPt) <= _tolerancePt);
        Assert.True(Math.Abs(placement.Top - expectedTopPt) <= _tolerancePt);
    }

    [Fact]
    public void A_group_larger_than_the_slide_is_scaled_down()
    {
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 1440, 1080), 720, 540, 4);

        SlidePlacement placement = Assert.IsType<SlidePlacement>(outcome.Placement);
        Assert.True(placement.Scale < 1);
        Assert.True(placement.Width <= 720 - 8 + _tolerancePt);
        Assert.True(placement.Height <= 540 - 8 + _tolerancePt);
    }

    [Fact]
    public void A_group_smaller_than_the_frame_is_scaled_up()
    {
        // D2's formula carries no upscale cap; fill-to-frame is the binding reading (M-7).
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 100, 50), 720, 540, 4);

        SlidePlacement placement = Assert.IsType<SlidePlacement>(outcome.Placement);
        Assert.True(placement.Scale > 1);
    }

    [Fact]
    public void A_zero_inset_makes_the_frame_the_whole_slide()
    {
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 800, 200), 720, 540, 0);

        SlidePlacement placement = Assert.IsType<SlidePlacement>(outcome.Placement);
        Assert.True(placement.Left >= -_tolerancePt);
        Assert.True(placement.Top >= -_tolerancePt);
        Assert.True(placement.Left + placement.Width <= 720 + _tolerancePt);
        Assert.True(placement.Top + placement.Height <= 540 + _tolerancePt);
    }

    [Theory]
    [InlineData(double.NaN, 540)]
    [InlineData(double.PositiveInfinity, 540)]
    [InlineData(0, 540)]
    [InlineData(-720, 540)]
    [InlineData(720, double.NaN)]
    [InlineData(720, double.PositiveInfinity)]
    [InlineData(720, 0)]
    [InlineData(720, -540)]
    public void An_invalid_slide_size_is_refused(double slideWidthPt, double slideHeightPt)
    {
        SlideFitOutcome outcome = SlideFit.Fit(
            new RectD(0, 0, 800, 200),
            slideWidthPt,
            slideHeightPt,
            4);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SlideFitRefusal.InvalidSlideSize, outcome.Refusal);
        Assert.Null(outcome.Placement);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(-4)]
    public void An_invalid_inset_is_refused(double insetPt)
    {
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 800, 200), 720, 540, insetPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SlideFitRefusal.InvalidInset, outcome.Refusal);
    }

    [Fact]
    public void A_default_group_is_refused()
    {
        SlideFitOutcome outcome = SlideFit.Fit(default, 720, 540, 4);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SlideFitRefusal.InvalidGroupBounds, outcome.Refusal);
    }

    [Fact]
    public void A_zero_width_group_is_refused()
    {
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 0, 200), 720, 540, 4);

        Assert.Equal(SlideFitRefusal.InvalidGroupBounds, outcome.Refusal);
    }

    [Fact]
    public void A_zero_height_group_is_refused()
    {
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 800, 0), 720, 540, 4);

        Assert.Equal(SlideFitRefusal.InvalidGroupBounds, outcome.Refusal);
    }

    [Fact]
    public void An_inset_at_half_the_slide_width_leaves_no_frame()
    {
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 800, 200), 720, 540, 360);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SlideFitRefusal.DegenerateFitFrame, outcome.Refusal);
    }

    [Fact]
    public void An_inset_beyond_half_the_slide_width_leaves_no_frame()
    {
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 800, 200), 720, 540, 400);

        Assert.Equal(SlideFitRefusal.DegenerateFitFrame, outcome.Refusal);
    }

    [Fact]
    public void An_inset_at_half_the_slide_height_leaves_no_frame()
    {
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 800, 200), 720, 540, 270);

        Assert.Equal(SlideFitRefusal.DegenerateFitFrame, outcome.Refusal);
    }

    [Fact]
    public void A_non_finite_scale_is_refused()
    {
        // Both ratios overflow to infinity, so min() is infinite (CR-9).
        SlideFitOutcome outcome = SlideFit.Fit(
            new RectD(0, 0, double.Epsilon, double.Epsilon),
            double.MaxValue,
            double.MaxValue,
            0);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SlideFitRefusal.InvalidScale, outcome.Refusal);
    }

    [Fact]
    public void A_size_that_collapses_to_zero_at_display_precision_is_refused()
    {
        // Scale is 1, but the height snaps to zero at float precision; a zero-size shape is the
        // degenerate emission CR-9 forbids.
        SlideFitOutcome outcome = SlideFit.Fit(new RectD(0, 0, 720, double.Epsilon), 720, 540, 0);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SlideFitRefusal.InvalidScale, outcome.Refusal);
    }

    [Fact]
    public void Identical_inputs_produce_identical_placements()
    {
        SlideFitOutcome first = SlideFit.Fit(new RectD(0, 0, 800, 200), 960, 540, 4);
        SlideFitOutcome second = SlideFit.Fit(new RectD(0, 0, 800, 200), 960, 540, 4);

        SlidePlacement a = Assert.IsType<SlidePlacement>(first.Placement);
        SlidePlacement b = Assert.IsType<SlidePlacement>(second.Placement);
        Assert.Equal(a.Left, b.Left);
        Assert.Equal(a.Top, b.Top);
        Assert.Equal(a.Scale, b.Scale);
        Assert.Equal(a.Width, b.Width);
        Assert.Equal(a.Height, b.Height);
    }

    [Fact]
    public void A_placement_can_be_consumed_only_once()
    {
        SlidePlacement placement = Assert.IsType<SlidePlacement>(
            SlideFit.Fit(new RectD(0, 0, 800, 200), 720, 540, 4).Placement);

        Assert.False(placement.IsConsumed);
        Assert.True(placement.TryConsume());
        Assert.True(placement.IsConsumed);
        Assert.False(placement.TryConsume());
    }

    [Fact]
    public void A_valid_placement_round_trips_its_members()
    {
        SlidePlacement placement = new(10, 20, 0.5, 100, 50);

        Assert.Equal(10, placement.Left);
        Assert.Equal(20, placement.Top);
        Assert.Equal(0.5, placement.Scale);
        Assert.Equal(100, placement.Width);
        Assert.Equal(50, placement.Height);
    }

    [Fact]
    public void A_non_finite_left_is_refused()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new SlidePlacement(double.NaN, 0, 1, 100, 50));
    }

    [Fact]
    public void A_non_finite_top_is_refused()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new SlidePlacement(0, double.PositiveInfinity, 1, 100, 50));
    }

    [Fact]
    public void A_zero_scale_is_refused()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new SlidePlacement(0, 0, 0, 100, 50));
    }

    [Fact]
    public void A_negative_scale_is_refused()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new SlidePlacement(0, 0, -1, 100, 50));
    }

    [Fact]
    public void A_zero_width_is_refused()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new SlidePlacement(0, 0, 1, 0, 50));
    }

    [Fact]
    public void A_non_finite_height_is_refused()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new SlidePlacement(0, 0, 1, 100, double.NaN));
    }
}
