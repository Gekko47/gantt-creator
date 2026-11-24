namespace GanttCreator.Raster.Tests;

/// <summary>
/// The points-to-pixels conversion at 300 DPI (R8.1, re-scoped 2026-11-23). The PNG is
/// the chart's own size, so this is the only sizing boundary: no width or unit input, and
/// each dimension is rounded exactly once.
/// </summary>
/// <remarks>
/// The recorded R8.4 golden dimensions are reproduced here from the composition's points,
/// which is the proof that withdrawing the cm/in/px grammar changes no output.
/// </remarks>
public sealed class ExportSizeTests
{
    /// <summary>Converts a centimetre width to the points the composition carries.</summary>
    private static double PointsFromCm(double cm) => cm / 2.54 * ExportSize.PointsPerInch;

    [Fact]
    public void A_72_point_composition_is_300_pixels_square()
    {
        PixelDimensions px = ExportSize.ToPixelsFromPoints(72, 72);

        Assert.Equal(300, px.PixelWidth);
        Assert.Equal(300, px.PixelHeight);
    }

    [Fact]
    public void A_ten_centimetre_composition_converts_to_the_recorded_1181_pixels()
    {
        // The withdrawn 10 cm arm's number: 10 / 2.54 * 300 = 1181.10… → 1181.
        PixelDimensions px = ExportSize.ToPixelsFromPoints(PointsFromCm(10), 360);

        Assert.Equal(1181, px.PixelWidth);
    }

    [Fact]
    public void A_four_inch_composition_converts_to_the_recorded_1200_pixels()
    {
        // 4 in = 288 pt; 288 / 72 * 300 = 1200 exactly.
        PixelDimensions px = ExportSize.ToPixelsFromPoints(288, 144);

        Assert.Equal(1200, px.PixelWidth);
        Assert.Equal(600, px.PixelHeight);
    }

    [Fact]
    public void An_a4_landscape_composition_reproduces_the_ratified_golden_dimensions()
    {
        // The R8.4 ratified width: 29.7 cm = 841.8897637795276 pt → 3508 px. The scene's
        // aspect yields the recorded 1906 px height (1906 * 72 / 300 = 457.44 pt).
        const double widthPt = 841.8897637795276;
        const double heightPt = 457.44;

        PixelDimensions px = ExportSize.ToPixelsFromPoints(widthPt, heightPt);

        Assert.Equal(3508, px.PixelWidth);
        Assert.Equal(1906, px.PixelHeight);
    }

    [Fact]
    public void The_dimensions_preserve_the_compositions_aspect()
    {
        // 720 x 360 pt is 2:1, so the pixels keep a 2:1 ratio.
        PixelDimensions px = ExportSize.ToPixelsFromPoints(720, 360);

        Assert.Equal(3000, px.PixelWidth);
        Assert.Equal(1500, px.PixelHeight);
    }

    [Fact]
    public void A_pixel_target_converts_to_points_and_back_unchanged()
    {
        // 800 px at 300 DPI is 800 * 72 / 300 = 192 pt; the conversion round-trips.
        PixelDimensions px = ExportSize.ToPixelsFromPoints(192, 96);

        Assert.Equal(800, px.PixelWidth);
        Assert.Equal(400, px.PixelHeight);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void A_non_positive_width_is_refused(double widthPt)
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(widthPt, 360));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void A_non_positive_height_is_refused(double heightPt)
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(720, heightPt));
    }

    [Fact]
    public void A_nan_width_is_refused()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(double.NaN, 360));
    }

    [Fact]
    public void An_infinite_height_is_refused()
    {
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(720, double.PositiveInfinity));
    }

    [Fact]
    public void A_width_near_the_per_axis_maximum_is_accepted()
    {
        // 15,728 pt is 65,533.3 px — inside the 65,535 px per-axis limit.
        PixelDimensions px = ExportSize.ToPixelsFromPoints(15_728, 16);

        Assert.Equal(65_533, px.PixelWidth);
    }

    [Fact]
    public void A_width_above_the_per_axis_limit_is_refused()
    {
        // 15,728.5 pt is 65,535.4 px, above the per-axis limit before rounding.
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(15_728.5, 16));
    }

    [Fact]
    public void A_height_above_the_per_axis_limit_is_refused()
    {
        // 72,000 pt is 300,000 px, well above the practical per-axis limit.
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(1000, 72_000));
    }

    [Fact]
    public void The_total_budget_allows_a_square_at_the_boundary()
    {
        // 2,400 pt is 10,000 px; 10,000 x 10,000 = 100,000,000.
        PixelDimensions px = ExportSize.ToPixelsFromPoints(2400, 2400);

        Assert.Equal(10_000, px.PixelWidth);
        Assert.Equal(10_000, px.PixelHeight);
    }

    [Fact]
    public void The_total_budget_rejects_a_square_beyond_the_boundary()
    {
        // 2,401 pt is 10,004.17 px → 10,004² = 100,080,016 > 100,000,000.
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(2401, 2401));
    }

    [Fact]
    public void A_sub_pixel_width_that_rounds_to_zero_is_refused()
    {
        // 0.1 pt is 0.4167 px: positive and finite pre-rounding, zero after.
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(0.1, 100));
    }

    [Fact]
    public void A_sub_pixel_height_that_rounds_to_zero_is_refused()
    {
        // The width is fine; the height rounds to a zero-pixel export.
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixelsFromPoints(100, 0.1));
    }
}
