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
        var result = ExportSize.ParseWidth(input);
        Assert.Equal(expectedUnit, result.Unit);
        Assert.Equal(expectedValue, result.Value);
    }

    [Theory]
    [InlineData("10")]
    [InlineData("cm")]
    [InlineData("")]
    [InlineData("-5in")]
    [InlineData("metres")]
    public void ParseWidth_rejects_invalid_format(string input)
    {
        Assert.Throws<FormatException>(() => ExportSize.ParseWidth(input));
    }

    [Fact]
    public void ParseWidth_null_throws()
    {
        Assert.Throws<ArgumentNullException>(() => ExportSize.ParseWidth(null!));
    }

    [Fact]
    public void ParseWidth_is_culture_invariant()
    {
        // "2.5in" must parse identically regardless of the current
        // culture's decimal separator. This is the core offline/localisation
        // guarantee from the entity guide.
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            var result = ExportSize.ParseWidth("2.5in");
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
        var px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(1181, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_in_uses_correct_conversion()
    {
        // 4 in at 300 DPI = 1200 px
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        var px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(1200, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_px_preserves_requested_width()
    {
        var req = new WidthRequest(ExportUnit.Pixels, 600.0);
        var px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(600, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_height_respects_aspect_ratio()
    {
        // 720 pt wide, 360 pt tall → 2:1 ratio. Height = half of width.
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        var px = ExportSize.ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);

        Assert.Equal(1200, px.PixelWidth);
        Assert.Equal(600, px.PixelHeight);
    }

    [Fact]
    public void ToPixels_zero_scene_width_throws()
    {
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        Assert.Throws<ArgumentOutOfRangeException>(
            () => ExportSize.ToPixels(req, sceneWidthPt: 0.0, sceneHeightPt: 360.0));
    }
}