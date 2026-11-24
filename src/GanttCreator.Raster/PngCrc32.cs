namespace GanttCreator.Raster;

/// <summary>
/// Computes the CRC-32 checksum (ISO-HDLC, reflected polynomial
/// <c>0xEDB88320</c>) that every PNG chunk carries over its type and data.
/// </summary>
/// <remarks>
/// <para>
/// This is the single CRC authority used by the encoder's density patch and the
/// post-write validator. The tests compute CRC-32 through their own independent
/// implementation, so a wrong table cannot satisfy both directions.
/// </para>
/// </remarks>
public static class PngCrc32
{
    /// <summary>The CRC-32 of the ASCII string <c>123456789</c> (the standard check value).</summary>
    public const uint StandardCheckValue = 0xCBF43926u;

    private static readonly uint[] _table = CreateTable();

    /// <summary>Computes the CRC-32 of <paramref name="data"/>.</summary>
    /// <param name="data">The bytes to checksum.</param>
    /// <returns>The CRC-32 value.</returns>
    public static uint Compute(ReadOnlySpan<byte> data)
    {
        uint crc = 0xFFFFFFFFu;
        foreach (byte value in data)
        {
            crc = _table[(crc ^ value) & 0xFF] ^ (crc >> 8);
        }

        return crc ^ 0xFFFFFFFFu;
    }

    /// <summary>Computes the CRC-32 of two contiguous segments without copying them.</summary>
    /// <param name="first">The first segment.</param>
    /// <param name="second">The second segment.</param>
    /// <returns>The CRC-32 value of the concatenated segments.</returns>
    public static uint Compute(ReadOnlySpan<byte> first, ReadOnlySpan<byte> second)
    {
        uint crc = 0xFFFFFFFFu;
        foreach (byte value in first)
        {
            crc = _table[(crc ^ value) & 0xFF] ^ (crc >> 8);
        }

        foreach (byte value in second)
        {
            crc = _table[(crc ^ value) & 0xFF] ^ (crc >> 8);
        }

        return crc ^ 0xFFFFFFFFu;
    }

    private static uint[] CreateTable()
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
}
