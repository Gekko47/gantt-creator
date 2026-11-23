using System.Buffers.Binary;
using System.IO;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// A deliberately separate, minimal PNG reader for R8.8 D2's independent verification:
/// it walks the chunks, checks each CRC with its own implementation, and reports the
/// <c>IHDR</c> geometry and the <c>pHYs</c> density without touching
/// <see cref="GanttCreator.Raster.PngFileValidator"/>, <see cref="GanttCreator.Raster.PngCrc32"/>
/// or any other writer-side helper. If the writer and this parser agree, the file is
/// what it claims to be; if only the writer's own check passes, nothing is proven.
/// </summary>
internal static class PngFileInspection
{
    private static readonly byte[] Signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    /// <summary>Reads and verifies one PNG file independently of the production writer.</summary>
    /// <param name="path">The file to read.</param>
    /// <returns>The inspected geometry, density and chunk summary.</returns>
    /// <exception cref="InvalidDataException">The signature, a CRC, or a required chunk is wrong.</exception>
    internal static PngInspection Read(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        if (bytes.Length < Signature.Length
            || !bytes.AsSpan(0, Signature.Length).SequenceEqual(Signature))
        {
            throw new InvalidDataException("The file does not start with the PNG signature.");
        }

        int offset = Signature.Length;
        int width = 0;
        int height = 0;
        byte bitDepth = 0;
        byte colourType = 0;
        int physX = 0;
        int physY = 0;
        byte physUnit = 0;
        int dataLength = 0;
        bool sawHeader = false;
        bool sawDensity = false;
        var chunkTypes = new List<string>();

        while (offset + 12 <= bytes.Length)
        {
            int length = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(offset, 4));
            if (length < 0 || offset + 12 + length > bytes.Length)
            {
                throw new InvalidDataException($"Chunk at {offset} has an impossible length {length}.");
            }

            string type = System.Text.Encoding.ASCII.GetString(bytes, offset + 4, 4);
            ReadOnlySpan<byte> data = bytes.AsSpan(offset + 8, length);
            uint stored = BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset + 8 + length, 4));
            if (stored != Crc(bytes.AsSpan(offset + 4, 4), data))
            {
                throw new InvalidDataException($"Chunk {type} carries a wrong CRC.");
            }

            chunkTypes.Add(type);
            switch (type)
            {
                case "IHDR":
                    width = BinaryPrimitives.ReadInt32BigEndian(data[..4]);
                    height = BinaryPrimitives.ReadInt32BigEndian(data.Slice(4, 4));
                    bitDepth = data[8];
                    colourType = data[9];
                    sawHeader = true;
                    break;
                case "pHYs":
                    physX = BinaryPrimitives.ReadInt32BigEndian(data[..4]);
                    physY = BinaryPrimitives.ReadInt32BigEndian(data.Slice(4, 4));
                    physUnit = data[8];
                    sawDensity = true;
                    break;
                case "IDAT":
                    dataLength += length;
                    break;
            }

            offset += 12 + length;
            if (type == "IEND")
            {
                break;
            }
        }

        if (!sawHeader)
        {
            throw new InvalidDataException("The file carries no IHDR chunk.");
        }

        return new PngInspection(
            width,
            height,
            bitDepth,
            colourType,
            sawDensity ? physX : 0,
            sawDensity ? physY : 0,
            physUnit,
            dataLength,
            chunkTypes);
    }

    /// <summary>Computes a PNG chunk CRC with an independent table-less implementation.</summary>
    /// <param name="type">The four chunk-type bytes.</param>
    /// <param name="data">The chunk payload.</param>
    /// <returns>The CRC32 value PNG stores after the payload.</returns>
    private static uint Crc(ReadOnlySpan<byte> type, ReadOnlySpan<byte> data)
    {
        uint crc = 0xFFFFFFFFu;
        foreach (byte value in type)
        {
            crc = Update(crc, value);
        }

        foreach (byte value in data)
        {
            crc = Update(crc, value);
        }

        return ~crc;
    }

    private static uint Update(uint crc, byte value)
    {
        crc ^= value;
        for (int bit = 0; bit < 8; bit++)
        {
            crc = (crc & 1) != 0 ? 0xEDB88320u ^ (crc >> 1) : crc >> 1;
        }

        return crc;
    }
}

/// <summary>What the independent PNG reader found.</summary>
/// <param name="Width">The IHDR width in pixels.</param>
/// <param name="Height">The IHDR height in pixels.</param>
/// <param name="BitDepth">The IHDR bit depth.</param>
/// <param name="ColourType">The IHDR colour type.</param>
/// <param name="PixelsPerMetreX">The pHYs horizontal density, or zero when absent.</param>
/// <param name="PixelsPerMetreY">The pHYs vertical density, or zero when absent.</param>
/// <param name="PhysUnit">The pHYs unit: 1 means metre.</param>
/// <param name="DataLength">The summed IDAT payload length.</param>
/// <param name="ChunkTypes">Every chunk type in file order.</param>
internal sealed record PngInspection(
    int Width,
    int Height,
    byte BitDepth,
    byte ColourType,
    int PixelsPerMetreX,
    int PixelsPerMetreY,
    byte PhysUnit,
    int DataLength,
    IReadOnlyList<string> ChunkTypes);
