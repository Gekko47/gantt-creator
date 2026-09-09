using System.Globalization;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// Tests for <see cref="ExportSize"/>: width/unit parsing and the
/// scene-aspect pixel conversion. Pure tests; no SkiaSharp or file I/O.
/// </summary>
public class ExportSizeTests
{
    [Theory]
    [InlineData("10cm", ExportUnit.Centimetres, 10.0)]
    [InlineData("4in", ExportUnit.Inches, 4.0)]
    [InlineData("800px", ExportUnit.Pixels, 800.0)]
    [InlineData("  15 cm  ", ExportUnit.Centimetres, 15.0)]
    [InlineData("2.5IN", ExportUnit.Inches, 2.5)]
    public void ParseWidth_accepts_valid_formats(string input, ExportUnit expectedUnit, double expectedValue)
    {
        WidthRequest result = ExportSize.ParseWidth(input);
        Assert.Equal(expectedUnit, result.Unit);
        Assert.Equal(expectedValue, result.Value);
    }

    [Theory]
    [InlineData("10")]
    [InlineData("cm")]
    [InlineData("")]
    [InlineData("-5in")]
    [InlineData("metres")]
    public void ParseWidth_rejects_invalid_format(string input) => _ = Assert.Throws<FormatException>(() => ExportSize.ParseWidth(input));

    [Theory]
    [InlineData("NaNcm")]
    [InlineData("Infinitycm")]
    [InlineData("-Infinitycm")]
    [InlineData("1e999cm")]
    [InlineData("NaNpx")]
    [InlineData("Infinityin")]
    public void ParseWidth_rejects_non_finite_values(string input) => _ = Assert.Throws<FormatException>(() => ExportSize.ParseWidth(input));

    [Fact]
    public void ParseWidth_null_throws() => _ = Assert.Throws<ArgumentNullException>(() => ExportSize.ParseWidth(null!));

    [Fact]
    public void ParseWidth_is_culture_invariant()
    {
        // "2.5in" must parse identically regardless of the current
        // culture's decimal separator. This is the core offline/localisation
        // guarantee from the entity guide.
        CultureInfo previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            WidthRequest result = ExportSize.ParseWidth("2.5in");
            Assert.Equal(2.5, result.Value);
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    [Fact]
    public void ToPixels_cm_uses_correct_conversion()
    {
        // 10 cm at 300 DPI = 10/2.54 * 300 ≈ 1181 px
        var req = new WidthRequest(ExportUnit.Centimetres, 10.0);
        PixelDimensions px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(1181, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_in_uses_correct_conversion()
    {
        // 4 in at 300 DPI = 1200 px
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        PixelDimensions px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(1200, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_px_preserves_requested_width()
    {
        var req = new WidthRequest(ExportUnit.Pixels, 600.0);
        PixelDimensions px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(600, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_height_respects_aspect_ratio()
    {
        // 720 pt wide, 360 pt tall → 2:1 ratio. Height = half of width.
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        PixelDimensions px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);

        Assert.Equal(1200, px.PixelWidth);
        Assert.Equal(600, px.PixelHeight);
    }

    [Fact]
    public void ToPixels_zero_scene_width_throws()
    {
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 0.0, sceneHeightPt: 360.0));
    }

    [Fact]
    public void ToPixels_negative_request_value_throws()
    {
        var req = new WidthRequest(ExportUnit.Pixels, -5.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0));
    }

    [Fact]
    public void ToPixels_nan_request_value_throws()
    {
        var req = new WidthRequest(ExportUnit.Pixels, double.NaN);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0));
    }

    [Fact]
    public void ToPixels_infinity_request_value_throws()
    {
        var req = new WidthRequest(ExportUnit.Pixels, double.PositiveInfinity);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0));
    }

    [Fact]
    public void ToPixels_oversized_pixel_width_throws()
    {
        // (int.MaxValue + 1) pixels exceeds the int range for PixelDimensions.
        var req = new WidthRequest(ExportUnit.Pixels, int.MaxValue + 1.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0));
    }

    [Fact]
    public void ToPixels_nan_scene_height_throws()
    {
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: double.NaN));
    }

    [Fact]
    public void ToPixels_nan_scene_width_throws()
    {
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: double.NaN, sceneHeightPt: 360.0));
    }

    [Theory]
    [InlineData("0px")]
    [InlineData("0cm")]
    [InlineData("0.000in")]
    [InlineData("-0.001cm")]
    public void ParseWidth_rejects_non_positive_width(string input) =>
        _ = Assert.Throws<FormatException>(() => ExportSize.ParseWidth(input));

    [Fact]
    public void ToPixels_zero_scene_height_throws()
    {
        // A zero scene height yields a zero computed pixel height, which the
        // height guard must reject so ToPixels cannot return a PixelDimensions
        // with a zero height.
        var req = new WidthRequest(ExportUnit.Pixels, 600.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 0.0));
    }

    [Fact]
    public void ToPixels_zero_request_value_throws()
    {
        var req = new WidthRequest(ExportUnit.Pixels, 0.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0));
    }

    [Fact]
    public void ToPixels_max_dimension_is_accepted()
    {
        // 4:1 scene keeps the total within MaxTotalPixels (65535x16384 ≈ 1.07B < 2B)
        // while still exercising the per-axis MaxPixelDimension boundary.
        var req = new WidthRequest(ExportUnit.Pixels, ExportSize.MaxPixelDimension);
        PixelDimensions px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 180.0);
        Assert.Equal(65_535, px.PixelWidth);
        Assert.Equal(16_384, px.PixelHeight);
    }

    [Fact]
    public void ToPixels_width_above_limit_throws()
    {
        var req = new WidthRequest(ExportUnit.Pixels, ExportSize.MaxPixelDimension + 1);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0));
    }

    [Fact]
    public void ToPixels_height_above_limit_throws()
    {
        // 1000 px wide on a 720 x 72,000 pt scene → height = 1000 * 72000 / 720
        // = 100,000 px, above the practical export limit.
        var req = new WidthRequest(ExportUnit.Pixels, 1000.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 72_000.0));
    }

    [Fact]
    public void ToPixels_total_budget_allows_square_at_boundary()
    {
        // 44,721 * 44,721 = 1,999,967,841 < MaxTotalPixels (2,000,000,000).
        // Square scene (1:1) so height == width.
        var req = new WidthRequest(ExportUnit.Pixels, 44_721.0);
        PixelDimensions px = ExportSize.ToPixels(req, sceneWidthPt: 100.0, sceneHeightPt: 100.0);
        Assert.Equal(44_721, px.PixelWidth);
        Assert.Equal(44_721, px.PixelHeight);
    }

    [Fact]
    public void ToPixels_total_budget_rejects_square_beyond_boundary()
    {
        // 44,722 * 44,722 = 2,000,057,284 > MaxTotalPixels (2,000,000,000).
        var req = new WidthRequest(ExportUnit.Pixels, 44_722.0);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 100.0, sceneHeightPt: 100.0));
    }

    [Fact]
    public void ToPixels_subpixel_width_rounding_to_zero_is_rejected()
    {
        // 0.4 px is positive and finite pre-rounding but rounds to 0.
        var req = new WidthRequest(ExportUnit.Pixels, 0.4);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 100.0, sceneHeightPt: 100.0));
    }

    [Fact]
    public void ToPixels_subpixel_height_rounding_to_zero_is_rejected()
    {
        // Width 0.6 rounds to 1 (OK), but height = 0.6 * (1/100) = 0.006 rounds to 0.
        var req = new WidthRequest(ExportUnit.Pixels, 0.6);
        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 100.0, sceneHeightPt: 1.0));
    }
}
