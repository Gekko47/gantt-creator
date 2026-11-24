using GanttCreator.Core;
using Xunit.Sdk;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// Acceptance pins for the R8.5 encoder: exact dimensions per unit, the
/// 11811-pixels-per-metre density on both axes with the metre unit, legal chunk
/// order, valid CRCs, and the wrong-density comparator positives.
/// </summary>
/// <remarks>
/// Every assertion reads the produced PNG bytes through the test project's own
/// parser and CRC (never the production reader), so the pins observe the file
/// exactly as a consumer would.
/// </remarks>
public class PngMetadataTests
{
    [Theory]
    [InlineData(56.69291338582677, 30.8)]   // ≈2 cm wide at the scene's aspect
    [InlineData(72, 39.12)]                  // 1 in wide
    [InlineData(192, 104.32)]                // an 800 px target at 300 DPI
    public void Dimensions_are_exact_for_each_composition_size(double widthPt, double heightPt)
    {
        PixelDimensions dimensions = ExportSize.ToPixelsFromPoints(widthPt, heightPt);

        using RasterRenderResult result = PngTestSupport.RenderRect(dimensions);
        PngExportBytes png = PngExportEncoder.Encode(result);

        Assert.Equal(dimensions, png.Dimensions);
        PngTestSupport.AssertContract(png.Bytes.ToArray(), dimensions);
    }

    [Fact]
    public void Density_is_11811_pixels_per_metre_on_both_axes_with_the_metre_unit()
    {
        using RasterRenderResult result = PngTestSupport.RenderRect(new PixelDimensions(236, 128));
        PngExportBytes png = PngExportEncoder.Encode(result);

        PngTestSupport.AssertContract(png.Bytes.ToArray(), png.Dimensions);
    }

    [Fact]
    public void Chunk_order_places_the_density_chunk_between_IHDR_and_the_first_IDAT()
    {
        byte[] bytes = EncodeSmall();
        List<TestPngChunk> chunks = PngTestSupport.Parse(bytes);

        int ihdrIndex = chunks.FindIndex(chunk => chunk.Type == "IHDR");
        int densityIndex = chunks.FindIndex(chunk => chunk.Type == "pHYs");
        int firstIdatIndex = chunks.FindIndex(chunk => chunk.Type == "IDAT");

        Assert.Equal(0, ihdrIndex);
        Assert.True(densityIndex > ihdrIndex);
        Assert.True(densityIndex < firstIdatIndex);
        Assert.Equal("IEND", chunks[^1].Type);
    }

    [Fact]
    public void Wrong_density_fixture_fails_the_comparator()
    {
        PixelDimensions dimensions = new(236, 128);
        byte[] wrong = PngTestSupport.SetDensity(EncodeSmall(), 3780, 3780, 1);

        XunitException exception = Assert.ThrowsAny<XunitException>(
            () => PngTestSupport.AssertContract(wrong, dimensions));

        Assert.Contains("pHYs", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_single_axis_wrong_density_fixture_fails_the_comparator()
    {
        PixelDimensions dimensions = new(236, 128);
        byte[] wrong = PngTestSupport.SetDensity(
            EncodeSmall(),
            PngTestSupport.PixelsPerMetre,
            3780,
            1);

        XunitException exception = Assert.ThrowsAny<XunitException>(
            () => PngTestSupport.AssertContract(wrong, dimensions));

        Assert.Contains("pHYs", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_wrong_unit_fixture_fails_the_comparator()
    {
        PixelDimensions dimensions = new(236, 128);
        byte[] wrong = PngTestSupport.SetDensity(
            EncodeSmall(),
            PngTestSupport.PixelsPerMetre,
            PngTestSupport.PixelsPerMetre,
            0);

        XunitException exception = Assert.ThrowsAny<XunitException>(
            () => PngTestSupport.AssertContract(wrong, dimensions));

        Assert.Contains("pHYs", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Encode_rejects_a_null_result()
    {
        Assert.Throws<ArgumentNullException>(() => PngExportEncoder.Encode(null!));
    }

    [Fact]
    public void ApplyDensity_rejects_bytes_that_are_not_a_png()
    {
        Assert.Throws<InvalidDataException>(
            () => PngExportEncoder.ApplyDensity([1, 2, 3], new PixelDimensions(1, 1)));
    }

    [Fact]
    public void ApplyDensity_rejects_IHDR_dimensions_that_disagree_with_the_render()
    {
        byte[] encoderBytes = EncodeSmall();

        Assert.Throws<InvalidDataException>(
            () => PngExportEncoder.ApplyDensity(encoderBytes, new PixelDimensions(999, 128)));
    }

    [Fact]
    public void ApplyDensity_replaces_an_existing_density_chunk_in_place()
    {
        PixelDimensions dimensions = new(236, 128);
        byte[] withExistingDensity = PngTestSupport.InsertDensity(EncodeSmall(), 3780, 3780, 0);

        byte[] patched = PngExportEncoder.ApplyDensity(withExistingDensity, dimensions);

        // AssertContract proves exactly one pHYs chunk remains and that it
        // carries 11811/11811 with the metre unit, so the fake was replaced
        // rather than supplemented.
        PngTestSupport.AssertContract(patched, dimensions);
    }

    private static byte[] EncodeSmall()
    {
        using RasterRenderResult result = PngTestSupport.RenderRect(new PixelDimensions(236, 128));
        return PngExportEncoder.Encode(result).Bytes.ToArray();
    }
}
