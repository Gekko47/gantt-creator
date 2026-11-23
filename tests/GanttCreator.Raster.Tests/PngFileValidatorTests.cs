using GanttCreator.Core;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// Acceptance pins for the R8.6 post-write validator: the valid file passes and
/// each corrupted fact is named, so a corrupt file can never be reported as a
/// successful write.
/// </summary>
public class PngFileValidatorTests
{
    private static readonly PixelDimensions _dimensions = new(60, 40);

    [Fact]
    public void A_valid_written_file_passes()
    {
        using TestTempDirectory directory = new();
        string path = directory.File("valid.png");
        File.WriteAllBytes(path, PngTestSupport.Encode(_dimensions).Bytes.ToArray());

        PngValidationResult result = PngFileValidator.Validate(path, _dimensions);

        Assert.True(result.IsValid);
        Assert.Equal(string.Empty, result.Failure);
    }

    [Fact]
    public void A_missing_signature_is_reported()
    {
        byte[] corrupt = PngTestSupport.Encode(_dimensions).Bytes.ToArray();
        corrupt[0] ^= 0xFF;

        AssertFailure(corrupt, "signature");
    }

    [Fact]
    public void A_truncated_structure_is_reported()
    {
        byte[] valid = PngTestSupport.Encode(_dimensions).Bytes.ToArray();
        byte[] corrupt = valid[..^5];

        AssertFailure(corrupt, "structure");
    }

    [Fact]
    public void Wrong_dimensions_are_reported()
    {
        byte[] corrupt = PngTestSupport.SetIhdrDimensions(
            PngTestSupport.Encode(_dimensions).Bytes.ToArray(),
            61,
            40);

        AssertFailure(corrupt, "dimensions");
    }

    [Fact]
    public void Wrong_density_is_reported()
    {
        byte[] corrupt = PngTestSupport.SetDensity(
            PngTestSupport.Encode(_dimensions).Bytes.ToArray(),
            3780,
            3780,
            1);

        AssertFailure(corrupt, "density");
    }

    [Fact]
    public void A_wrong_unit_is_reported()
    {
        byte[] corrupt = PngTestSupport.SetDensity(
            PngTestSupport.Encode(_dimensions).Bytes.ToArray(),
            PngTestSupport.PixelsPerMetre,
            PngTestSupport.PixelsPerMetre,
            0);

        AssertFailure(corrupt, "unit");
    }

    [Fact]
    public void A_missing_density_chunk_is_reported()
    {
        byte[] corrupt = PngTestSupport.RemoveChunk(
            PngTestSupport.Encode(_dimensions).Bytes.ToArray(),
            "pHYs");

        AssertFailure(corrupt, "pHYs");
    }

    [Fact]
    public void Empty_image_data_is_reported()
    {
        byte[] corrupt = PngTestSupport.EmptyFirstIdat(
            PngTestSupport.Encode(_dimensions).Bytes.ToArray());

        AssertFailure(corrupt, "image data");
    }

    [Fact]
    public void A_corrupted_chunk_crc_is_reported()
    {
        byte[] corrupt = PngTestSupport.CorruptIdatCrc(
            PngTestSupport.Encode(_dimensions).Bytes.ToArray());

        AssertFailure(corrupt, "CRC");
    }

    [Fact]
    public void Null_and_empty_paths_are_rejected()
    {
        Assert.Throws<ArgumentNullException>(() => PngFileValidator.Validate(null!, _dimensions));
        Assert.Throws<ArgumentException>(() => PngFileValidator.Validate("   ", _dimensions));
    }

    private static void AssertFailure(byte[] bytes, string expectedFact)
    {
        using TestTempDirectory directory = new();
        string path = directory.File("corrupt.png");
        File.WriteAllBytes(path, bytes);

        PngValidationResult result = PngFileValidator.Validate(path, _dimensions);

        Assert.False(result.IsValid);
        Assert.Contains(expectedFact, result.Failure, StringComparison.Ordinal);
    }
}
