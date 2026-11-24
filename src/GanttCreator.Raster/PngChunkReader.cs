using System.Text;

namespace GanttCreator.Raster;

/// <summary>One PNG chunk located within a byte buffer.</summary>
/// <param name="Type">The four-character chunk type.</param>
/// <param name="ChunkOffset">The offset of the chunk's four-byte length field.</param>
/// <param name="ChunkLength">The chunk's total length: twelve bytes plus its data.</param>
/// <param name="DataOffset">The offset of the chunk's data bytes.</param>
/// <param name="DataLength">The chunk's data length in bytes.</param>
/// <param name="StoredCrc">The CRC stored in the chunk's final four bytes.</param>
internal readonly record struct PngChunk(
    string Type,
    int ChunkOffset,
    int ChunkLength,
    int DataOffset,
    int DataLength,
    uint StoredCrc);

/// <summary>
/// Walks the chunk structure of a PNG byte buffer. Deliberately minimal: it
/// locates chunks, never interprets image data, and throws on any structural
/// fault so callers cannot trust a truncated stream.
/// </summary>
internal static class PngChunkReader
{
    /// <summary>The eight-byte PNG signature.</summary>
    internal const int SignatureLength = 8;

    private static ReadOnlySpan<byte> Signature => [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    /// <summary>Returns whether the buffer starts with the PNG signature.</summary>
    internal static bool HasSignature(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= SignatureLength && bytes[..SignatureLength].SequenceEqual(Signature);

    /// <summary>Reads every chunk in the buffer.</summary>
    /// <param name="bytes">The PNG bytes, signature included.</param>
    /// <returns>The chunks in file order.</returns>
    /// <exception cref="InvalidDataException">Thrown when the signature, a chunk header, or a chunk length is malformed.</exception>
    internal static List<PngChunk> ReadChunks(ReadOnlySpan<byte> bytes)
    {
        if (!HasSignature(bytes))
        {
            throw new InvalidDataException("The buffer does not start with the PNG signature.");
        }

        List<PngChunk> chunks = [];
        int offset = SignatureLength;
        while (offset < bytes.Length)
        {
            if (bytes.Length - offset < 12)
            {
                throw new InvalidDataException("A PNG chunk header is truncated.");
            }

            int dataLength = ReadInt32BigEndian(bytes, offset);
            if (dataLength < 0 || dataLength > bytes.Length - offset - 12)
            {
                throw new InvalidDataException("A PNG chunk length exceeds the buffer.");
            }

            string type = Encoding.ASCII.GetString(bytes.Slice(offset + 4, 4));
            int dataOffset = offset + 8;
            uint storedCrc = ReadUInt32BigEndian(bytes, offset + 8 + dataLength);
            chunks.Add(new PngChunk(type, offset, 12 + dataLength, dataOffset, dataLength, storedCrc));
            offset += 12 + dataLength;
        }

        return chunks;
    }

    /// <summary>Reads a big-endian 32-bit integer.</summary>
    internal static int ReadInt32BigEndian(ReadOnlySpan<byte> bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    /// <summary>Reads a big-endian unsigned 32-bit integer.</summary>
    internal static uint ReadUInt32BigEndian(ReadOnlySpan<byte> bytes, int offset) =>
        (uint)ReadInt32BigEndian(bytes, offset);

    /// <summary>Writes a big-endian unsigned 32-bit integer.</summary>
    internal static void WriteUInt32BigEndian(Span<byte> bytes, int offset, uint value)
    {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }
}
