using System.Text;
using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Raster.Tests;

/// <summary>One chunk as the independent test parser locates it.</summary>
internal sealed record TestPngChunk(
    string Type,
    int ChunkOffset,
    int ChunkLength,
    int DataOffset,
    int DataLength,
    uint StoredCrc);

/// <summary>
/// Byte-level support for the PNG rows. The parser and CRC are written here,
/// independently of the production reader, so a defect cannot satisfy both the
/// implementation and its proof (R8.5 D1, R8.8 D2).
/// </summary>
internal static class PngTestSupport
{
    public const int PixelsPerMetre = 11811;

    private static readonly byte[] _signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    private static readonly uint[] _crcTable = BuildCrcTable();

    /// <summary>Parses the chunk structure of <paramref name="png"/>, asserting the fixed facts.</summary>
    public static List<TestPngChunk> Parse(byte[] png)
    {
        Assert.True(png.Length >= _signature.Length + 12, "The PNG is too short.");
        Assert.True(png.AsSpan(0, _signature.Length).SequenceEqual(_signature), "The PNG signature is wrong.");

        List<TestPngChunk> chunks = [];
        int offset = _signature.Length;
        while (offset < png.Length)
        {
            Assert.True(png.Length - offset >= 12, "A PNG chunk header is truncated.");
            int length = ReadInt32(png, offset);
            Assert.True(length >= 0 && length <= png.Length - offset - 12, "A PNG chunk length overruns the buffer.");
            string type = Encoding.ASCII.GetString(png, offset + 4, 4);
            chunks.Add(new TestPngChunk(
                type,
                offset,
                12 + length,
                offset + 8,
                length,
                ReadUInt32(png, offset + 8 + length)));
            offset += 12 + length;
        }

        return chunks;
    }

    /// <summary>Asserts the full R8.5 metadata contract on <paramref name="png"/>.</summary>
    public static void AssertContract(byte[] png, PixelDimensions expected)
    {
        List<TestPngChunk> chunks = Parse(png);
        Assert.True(chunks.Count >= 3, "The PNG has too few chunks.");
        Assert.Equal("IHDR", chunks[0].Type);
        Assert.Equal(13, chunks[0].DataLength);
        Assert.Equal(expected.PixelWidth, ReadInt32(png, chunks[0].DataOffset));
        Assert.Equal(expected.PixelHeight, ReadInt32(png, chunks[0].DataOffset + 4));

        int densityIndex = -1;
        int firstIdatIndex = -1;
        for (int index = 0; index < chunks.Count; index++)
        {
            if (chunks[index].Type == "pHYs")
            {
                Assert.True(densityIndex == -1, "The PNG contains more than one pHYs chunk.");
                densityIndex = index;
            }

            if (chunks[index].Type == "IDAT" && firstIdatIndex < 0)
            {
                firstIdatIndex = index;
            }
        }

        Assert.True(densityIndex > 0, "The pHYs chunk is missing or precedes IHDR.");
        Assert.True(firstIdatIndex > densityIndex, "The pHYs chunk must precede the first IDAT.");

        TestPngChunk density = chunks[densityIndex];
        Assert.Equal(9, density.DataLength);
        int x = ReadInt32(png, density.DataOffset);
        int y = ReadInt32(png, density.DataOffset + 4);
        Assert.True(x == PixelsPerMetre, $"The pHYs x density is {x}, expected {PixelsPerMetre}.");
        Assert.True(y == PixelsPerMetre, $"The pHYs y density is {y}, expected {PixelsPerMetre}.");
        Assert.True(png[density.DataOffset + 8] == 1, "The pHYs unit is not metres.");
        Assert.Equal("IEND", chunks[^1].Type);

        foreach (TestPngChunk chunk in chunks)
        {
            uint computed = Crc32(png.AsSpan(chunk.ChunkOffset + 4, 4), png.AsSpan(chunk.DataOffset, chunk.DataLength));
            Assert.True(computed == chunk.StoredCrc, $"The {chunk.Type} chunk CRC is wrong.");
        }
    }

    /// <summary>Rewrites the existing density chunk and fixes its CRC.</summary>
    public static byte[] SetDensity(byte[] png, int x, int y, byte unit)
    {
        List<TestPngChunk> chunks = Parse(png);
        TestPngChunk density = chunks.Find(chunk => chunk.Type == "pHYs")
            ?? throw new InvalidOperationException("The fixture has no pHYs chunk.");
        byte[] copy = (byte[])png.Clone();
        WriteUInt32(copy, density.DataOffset, (uint)x);
        WriteUInt32(copy, density.DataOffset + 4, (uint)y);
        copy[density.DataOffset + 8] = unit;
        WriteUInt32(
            copy,
            density.DataOffset + 9,
            Crc32(copy.AsSpan(density.ChunkOffset + 4, 4), copy.AsSpan(density.DataOffset, 9)));
        return copy;
    }

    /// <summary>Inserts a density chunk before the first IDAT (the replace-in-place fixture).</summary>
    public static byte[] InsertDensity(byte[] png, int x, int y, byte unit)
    {
        List<TestPngChunk> chunks = Parse(png);
        TestPngChunk firstIdat = chunks.Find(chunk => chunk.Type == "IDAT")
            ?? throw new InvalidOperationException("The fixture has no IDAT chunk.");
        byte[] densityChunk = BuildDensityChunk(x, y, unit);
        byte[] result = new byte[png.Length + densityChunk.Length];
        png.AsSpan(0, firstIdat.ChunkOffset).CopyTo(result);
        densityChunk.CopyTo(result, firstIdat.ChunkOffset);
        png.AsSpan(firstIdat.ChunkOffset).CopyTo(result.AsSpan(firstIdat.ChunkOffset + densityChunk.Length));
        return result;
    }

    /// <summary>Rewrites the IHDR dimensions and fixes the IHDR CRC.</summary>
    public static byte[] SetIhdrDimensions(byte[] png, int width, int height)
    {
        List<TestPngChunk> chunks = Parse(png);
        TestPngChunk ihdr = chunks[0];
        byte[] copy = (byte[])png.Clone();
        WriteUInt32(copy, ihdr.DataOffset, (uint)width);
        WriteUInt32(copy, ihdr.DataOffset + 4, (uint)height);
        WriteUInt32(
            copy,
            ihdr.DataOffset + 9,
            Crc32(copy.AsSpan(ihdr.ChunkOffset + 4, 4), copy.AsSpan(ihdr.DataOffset, 13)));
        return copy;
    }

    /// <summary>Empties the first IDAT's data and fixes its CRC.</summary>
    public static byte[] EmptyFirstIdat(byte[] png)
    {
        List<TestPngChunk> chunks = Parse(png);
        TestPngChunk idat = chunks.Find(chunk => chunk.Type == "IDAT")
            ?? throw new InvalidOperationException("The fixture has no IDAT chunk.");
        byte[] result = new byte[png.Length - idat.DataLength];
        png.AsSpan(0, idat.ChunkOffset).CopyTo(result);

        Span<byte> header = result.AsSpan(idat.ChunkOffset, 8);
        WriteUInt32(result, idat.ChunkOffset, 0);
        uint crc = Crc32(header[4..], ReadOnlySpan<byte>.Empty);
        WriteUInt32(result, idat.ChunkOffset + 8, crc);
        png.AsSpan(idat.ChunkOffset + idat.ChunkLength).CopyTo(result.AsSpan(idat.ChunkOffset + 12));
        return result;
    }

    /// <summary>Flips one bit inside the first IDAT's data without fixing the CRC.</summary>
    public static byte[] CorruptIdatCrc(byte[] png)
    {
        List<TestPngChunk> chunks = Parse(png);
        TestPngChunk idat = chunks.Find(chunk => chunk.Type == "IDAT" && chunk.DataLength > 0)
            ?? throw new InvalidOperationException("The fixture has no non-empty IDAT chunk.");
        byte[] copy = (byte[])png.Clone();
        copy[idat.DataOffset] ^= 0x01;
        return copy;
    }

    /// <summary>Encodes a small valid PNG for file-level tests.</summary>
    public static PngExportBytes Encode(PixelDimensions dimensions)
    {
        using RasterRenderResult result = RenderRect(dimensions);
        return PngExportEncoder.Encode(result);
    }

    /// <summary>Removes one chunk and returns the remaining bytes.</summary>
    public static byte[] RemoveChunk(byte[] png, string type)
    {
        List<TestPngChunk> chunks = Parse(png);
        TestPngChunk target = chunks.Find(chunk => chunk.Type == type)
            ?? throw new InvalidOperationException($"The fixture has no {type} chunk.");
        byte[] result = new byte[png.Length - target.ChunkLength];
        png.AsSpan(0, target.ChunkOffset).CopyTo(result);
        png.AsSpan(target.ChunkOffset + target.ChunkLength).CopyTo(result.AsSpan(target.ChunkOffset));
        return result;
    }

    /// <summary>Renders a small text-free scene at <paramref name="dimensions"/>.</summary>
    public static RasterRenderResult RenderRect(PixelDimensions dimensions)
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new SceneRect(
                "rect",
                SceneOwnerId.Chart,
                ZLayer.ActivityBody,
                new RectD(10, 10, 40, 30),
                new SceneStyle("Fixture", fillColour: ColourHex.Parse("#3366CC"))));
        return SceneRasterRenderer.Render(scene, dimensions, RasterBackground.Transparent);
    }

    /// <summary>Reads a big-endian 32-bit integer from the test buffer.</summary>
    public static int ReadInt32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    /// <summary>Writes a big-endian unsigned 32-bit integer into the test buffer.</summary>
    public static void WriteUInt32(byte[] bytes, int offset, uint value)
    {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    /// <summary>The test project's independent CRC-32.</summary>
    public static uint Crc32(ReadOnlySpan<byte> first, ReadOnlySpan<byte> second)
    {
        uint crc = 0xFFFFFFFFu;
        foreach (byte value in first)
        {
            crc = _crcTable[(crc ^ value) & 0xFF] ^ (crc >> 8);
        }

        foreach (byte value in second)
        {
            crc = _crcTable[(crc ^ value) & 0xFF] ^ (crc >> 8);
        }

        return crc ^ 0xFFFFFFFFu;
    }

    private static byte[] BuildDensityChunk(int x, int y, byte unit)
    {
        byte[] chunk = new byte[21];
        WriteUInt32(chunk, 0, 9);
        chunk[4] = (byte)'p';
        chunk[5] = (byte)'H';
        chunk[6] = (byte)'Y';
        chunk[7] = (byte)'s';
        WriteUInt32(chunk, 8, (uint)x);
        WriteUInt32(chunk, 12, (uint)y);
        chunk[16] = unit;
        WriteUInt32(chunk, 17, Crc32(chunk.AsSpan(4, 13), ReadOnlySpan<byte>.Empty));
        return chunk;
    }

    private static uint ReadUInt32(byte[] bytes, int offset) => (uint)ReadInt32(bytes, offset);

    private static uint[] BuildCrcTable()
    {
        uint[] table = new uint[256];
        for (int index = 0; index < table.Length; index++)
        {
            uint value = (uint)index;
            for (int bit = 0; bit < 8; bit++)
            {
                value = (value & 1) != 0 ? 0xEDB88320u ^ (value >> 1) : value >> 1;
            }

            table[index] = value;
        }

        return table;
    }

    private static GanttScene CreateScene(RectD bounds, params ScenePrimitive[] primitives)
    {
        SceneCreationOutcome outcome = GanttScene.TryCreate(bounds, bounds, primitives, []);
        Assert.True(outcome.Succeeded, "The fixture scene was refused.");
        return outcome.Scene!;
    }
}
