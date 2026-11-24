using System.Text;

namespace GanttCreator.Raster.Tests;

/// <summary>Pins the production CRC-32 authority against the standard check vector.</summary>
public class PngCrc32Tests
{
    [Fact]
    public void The_standard_check_value_is_pinned()
    {
        uint crc = PngCrc32.Compute(Encoding.ASCII.GetBytes("123456789"));

        Assert.Equal(0xCBF43926u, crc);
        Assert.Equal(PngCrc32.StandardCheckValue, crc);
    }

    [Fact]
    public void An_empty_input_hashes_to_zero()
    {
        Assert.Equal(0u, PngCrc32.Compute(ReadOnlySpan<byte>.Empty));
    }

    [Fact]
    public void Segmented_computation_matches_the_contiguous_one()
    {
        byte[] data = Encoding.ASCII.GetBytes("chunk-type-and-data");

        Assert.Equal(
            PngCrc32.Compute(data),
            PngCrc32.Compute(data.AsSpan(0, 5), data.AsSpan(5)));
    }
}
