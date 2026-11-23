namespace GanttCreator.Raster;

/// <summary>
/// Re-reads a written PNG file and verifies the pinned export contract:
/// signature, exact <c>IHDR</c> dimensions, 300-DPI <c>pHYs</c> density and
/// unit, non-empty image data, and every chunk's CRC.
/// </summary>
/// <remarks>
/// The checks are a deliberate superset of the architecture's four required
/// facts: the row's corrupt-file failure arm cannot be caught by signature,
/// dimensions, density, and non-empty content alone, so chunk CRCs are
/// recomputed too. The validator reads from disk, never from the encoder's
/// in-memory buffer, so success is grounded in bytes that exist in the file.
/// </remarks>
public static class PngFileValidator
{
    /// <summary>The expected density in pixels per metre (the encoder's contract).</summary>
    public const int ExpectedPixelsPerMetre = PngExportEncoder.PixelsPerMetre;

    /// <summary>Re-reads <paramref name="path"/> from disk and validates it.</summary>
    /// <param name="path">The file to validate.</param>
    /// <param name="expected">The exact dimensions the encoder declared.</param>
    /// <returns>A passing result, or the first failing fact.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is empty or whitespace.</exception>
    public static PngValidationResult Validate(string path, PixelDimensions expected)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return ValidateBytes(File.ReadAllBytes(path), expected);
    }

    /// <summary>Validates PNG bytes already read from a file.</summary>
    /// <param name="bytes">The file bytes.</param>
    /// <param name="expected">The exact dimensions the encoder declared.</param>
    /// <returns>A passing result, or the first failing fact.</returns>
    internal static PngValidationResult ValidateBytes(ReadOnlySpan<byte> bytes, PixelDimensions expected)
    {
        if (!PngChunkReader.HasSignature(bytes))
        {
            return new PngValidationResult(false, "The PNG signature is missing.");
        }

        List<PngChunk> chunks;
        try
        {
            chunks = PngChunkReader.ReadChunks(bytes);
        }
        catch (InvalidDataException)
        {
            return new PngValidationResult(false, "The PNG structure is truncated or malformed.");
        }

        if (chunks.Count == 0 || chunks[0].Type != "IHDR" || chunks[0].DataLength != 13)
        {
            return new PngValidationResult(false, "The IHDR chunk is missing or malformed.");
        }

        PngChunk ihdr = chunks[0];
        int width = PngChunkReader.ReadInt32BigEndian(bytes, ihdr.DataOffset);
        int height = PngChunkReader.ReadInt32BigEndian(bytes, ihdr.DataOffset + 4);
        if (width != expected.PixelWidth || height != expected.PixelHeight)
        {
            return new PngValidationResult(false, "The IHDR dimensions do not match the requested export.");
        }

        PngChunk? density = null;
        long imageDataLength = 0;
        foreach (PngChunk chunk in chunks)
        {
            if (chunk.Type == "pHYs" && density is null)
            {
                density = chunk;
            }

            if (chunk.Type == "IDAT")
            {
                imageDataLength += chunk.DataLength;
            }
        }

        if (density is not { DataLength: 9 })
        {
            return new PngValidationResult(false, "The pHYs density chunk is missing or malformed.");
        }

        PngChunk densityChunk = density.Value;
        int x = PngChunkReader.ReadInt32BigEndian(bytes, densityChunk.DataOffset);
        int y = PngChunkReader.ReadInt32BigEndian(bytes, densityChunk.DataOffset + 4);
        if (x != ExpectedPixelsPerMetre || y != ExpectedPixelsPerMetre)
        {
            return new PngValidationResult(false, "The pHYs density is not 11811 pixels per metre on both axes.");
        }

        if (bytes[densityChunk.DataOffset + 8] != 1)
        {
            return new PngValidationResult(false, "The pHYs unit is not metres.");
        }

        if (imageDataLength == 0)
        {
            return new PngValidationResult(false, "The image data is empty.");
        }

        foreach (PngChunk chunk in chunks)
        {
            uint computed = PngCrc32.Compute(
                bytes.Slice(chunk.ChunkOffset + 4, 4),
                bytes.Slice(chunk.DataOffset, chunk.DataLength));
            if (computed != chunk.StoredCrc)
            {
                return new PngValidationResult(false, $"A {chunk.Type} chunk CRC is invalid.");
            }
        }

        return new PngValidationResult(true, string.Empty);
    }
}
