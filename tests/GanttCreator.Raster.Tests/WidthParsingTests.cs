namespace GanttCreator.Raster.Tests;

/// <summary>
/// Tests for the raster export width/unit parsing logic.
/// Uses only pure functions; no SkiaSharp or file I/O.
/// </summary>
public class WidthParsingTests
{
    // Represent the unit enum as it will exist in the Raster project.
    private enum ExportUnit { Centimetres, Inches, Pixels }

    // Represents a parsed width request.
    private readonly record struct WidthRequest(ExportUnit Unit, double Value);

    // Represents the calculated pixel dimensions.
    private readonly record struct PixelDimensions(int PixelWidth, int PixelHeight);

    // Constants from the architecture spec.
    private const double Dpi = 300.0;
    private const double CmPerInch = 2.54;

    // Parses a string like "10cm", "4in", "800px" (case-insensitive).
    private static WidthRequest ParseWidth(string input)
    {
        input = input.Trim();
        if (input.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            return new WidthRequest(ExportUnit.Centimetres, double.Parse(input[..^2].Trim()));
        if (input.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            return new WidthRequest(ExportUnit.Inches, double.Parse(input[..^2].Trim()));
        if (input.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            return new WidthRequest(ExportUnit.Pixels, double.Parse(input[..^2].Trim()));
        throw new FormatException($"Unknown unit in '{input}'.");
    }

    // Converts a width request and scene aspect ratio to pixel dimensions.
    // aspectRatio = sceneHeightPoints / sceneWidthPoints.
    private static PixelDimensions ToPixels(WidthRequest req, double sceneWidthPt, double sceneHeightPt)
    {
        double pixelWidth = req.Unit switch
        {
            ExportUnit.Centimetres => req.Value / CmPerInch * Dpi,
            ExportUnit.Inches => req.Value * Dpi,
            ExportUnit.Pixels => req.Value,
            _ => throw new InvalidOperationException()
        };

        int pxWidth = (int)Math.Round(pixelWidth);
        int pxHeight = (int)Math.Round(pixelWidth * sceneHeightPt / sceneWidthPt);
        return new PixelDimensions(pxWidth, pxHeight);
    }

    [Theory]
    [InlineData("10cm", ExportUnit.Centimetres, 10.0)]
    [InlineData("4in", ExportUnit.Inches, 4.0)]
    [InlineData("800px", ExportUnit.Pixels, 800.0)]
    [InlineData("  15 cm  ", ExportUnit.Centimetres, 15.0)]
    [InlineData("2.5IN", ExportUnit.Inches, 2.5)]
    public void ParseWidth_accepts_valid_formats(string input, ExportUnit expectedUnit, double expectedValue)
    {
        var result = ParseWidth(input);
        Assert.Equal(expectedUnit, result.Unit);
        Assert.Equal(expectedValue, result.Value);
    }

    [Theory]
    [InlineData("10")]
    [InlineData("cm")]
    [InlineData("")]
    [InlineData("metres")]
    public void ParseWidth_rejects_invalid_format(string input)
    {
        Assert.Throws<FormatException>(() => ParseWidth(input));
    }

    [Fact]
    public void ToPixels_cm_uses_correct_conversion()
    {
        // 10 cm at 300 DPI = 10/2.54 * 300 ≈ 1181 px
        var req = new WidthRequest(ExportUnit.Centimetres, 10.0);
        var px = ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(1181, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_in_uses_correct_conversion()
    {
        // 4 in at 300 DPI = 1200 px
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        var px = ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(1200, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_px_preserves_requested_width()
    {
        var req = new WidthRequest(ExportUnit.Pixels, 600.0);
        var px = ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);
        Assert.Equal(600, px.PixelWidth);
    }

    [Fact]
    public void ToPixels_height_respects_aspect_ratio()
    {
        // 720 pt wide, 360 pt tall → 2:1 ratio
        var req = new WidthRequest(ExportUnit.Inches, 4.0);
        var px = ToPixels(req, sceneWidthPt: 720.0, sceneHeightPt: 360.0);

        Assert.Equal(1200, px.PixelWidth);
        Assert.Equal(600, px.PixelHeight); // half of width: 2:1 ratio
    }
}
