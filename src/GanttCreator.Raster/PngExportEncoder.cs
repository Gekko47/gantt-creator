using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// Encodes a rendered raster result as a PNG whose <c>pHYs</c> chunk declares
/// the pinned 300-DPI density.
/// </summary>
/// <remarks>
/// <para>
/// The pinned SkiaSharp encoder (4.157.0, ADR-0071) writes no <c>pHYs</c> chunk
/// — the R8.5 Step-0 probe recorded the exact layout
/// <c>IHDR → sBIT → IDAT → IEND</c> — so the density is added by a documented
/// post-encode chunk patch: the correct <c>pHYs</c> is inserted directly after
/// <c>IHDR</c> (always before the first <c>IDAT</c>), and any density chunk the
/// encoder might emit in a future version is replaced rather than duplicated,
/// because a second <c>pHYs</c> would make the file invalid.
/// </para>
/// <para>
/// The encoder never writes a file; <c>AtomicPngFileWriter</c> (R8.6) owns the
/// file stage.
/// </para>
/// </remarks>
public static class PngExportEncoder
{
    /// <summary>
    /// The density in pixels per metre for 300 DPI. The conversion is
    /// <c>300 × 100 / 2.54 = 11811.02…</c>, rounded once to the integer the PNG
    /// <c>pHYs</c> chunk stores.
    /// </summary>
    public const int PixelsPerMetre = 11811;

    private const int _pngQuality = 100;
    private const int _densityChunkLength = 21;

    /// <summary>Encodes <paramref name="result"/> as a PNG with the 300-DPI <c>pHYs</c> chunk.</summary>
    /// <param name="result">The owned render result to encode.</param>
    /// <returns>The encoded bytes and the dimensions they declare.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="result"/> is null.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the Skia encoder produces no bytes.</exception>
    /// <exception cref="InvalidDataException">Thrown when the encoder output is not a well-formed PNG or its dimensions disagree with the render.</exception>
    public static PngExportBytes Encode(RasterRenderResult result)
    {
        ArgumentNullException.ThrowIfNull(result);

        using SKData? encoded = result.Bitmap.Encode(SKEncodedImageFormat.Png, _pngQuality);
        if (encoded is null || encoded.Size == 0)
        {
            throw new InvalidOperationException("The Skia encoder produced no PNG bytes.");
        }

        PixelDimensions dimensions = new(result.PixelWidth, result.PixelHeight);
        byte[] patched = ApplyDensity(encoded.ToArray(), dimensions);
        return new PngExportBytes(patched, dimensions);
    }

    /// <summary>
    /// Applies the density contract to an encoded PNG: exactly one <c>pHYs</c>
    /// chunk immediately after <c>IHDR</c>, every other chunk preserved.
    /// </summary>
    /// <param name="encoderBytes">The encoder's own output.</param>
    /// <param name="expected">The dimensions the render requested.</param>
    /// <returns>The patched PNG bytes.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="encoderBytes"/> is null.</exception>
    /// <exception cref="InvalidDataException">Thrown when the input is not a well-formed PNG or its <c>IHDR</c> disagrees with <paramref name="expected"/>.</exception>
    internal static byte[] ApplyDensity(byte[] encoderBytes, PixelDimensions expected)
    {
        ArgumentNullException.ThrowIfNull(encoderBytes);

        ReadOnlySpan<byte> source = encoderBytes;
        List<PngChunk> chunks = PngChunkReader.ReadChunks(source);
        ValidateStructure(source, chunks, expected);

        int existingDensityBytes = 0;
        foreach (PngChunk chunk in chunks)
        {
            if (chunk.Type == "pHYs")
            {
                existingDensityBytes += chunk.ChunkLength;
            }
        }

        byte[] densityChunk = CreateDensityChunk();
        byte[] output = new byte[encoderBytes.Length - existingDensityBytes + densityChunk.Length];
        Span<byte> destination = output;

        source[..PngChunkReader.SignatureLength].CopyTo(destination);
        int writeOffset = PngChunkReader.SignatureLength;

        // IHDR is validated as the first chunk, so it is copied verbatim and the
        // density chunk is placed directly behind it, before every IDAT.
        PngChunk ihdr = chunks[0];
        source.Slice(ihdr.ChunkOffset, ihdr.ChunkLength).CopyTo(destination[writeOffset..]);
        writeOffset += ihdr.ChunkLength;
        densityChunk.CopyTo(destination[writeOffset..]);
        writeOffset += densityChunk.Length;

        for (int index = 1; index < chunks.Count; index++)
        {
            PngChunk chunk = chunks[index];
            if (chunk.Type == "pHYs")
            {
                continue;
            }

            source.Slice(chunk.ChunkOffset, chunk.ChunkLength).CopyTo(destination[writeOffset..]);
            writeOffset += chunk.ChunkLength;
        }

        return output;
    }

    private static void ValidateStructure(ReadOnlySpan<byte> source, List<PngChunk> chunks, PixelDimensions expected)
    {
        if (chunks.Count < 3 || chunks[0].Type != "IHDR" || chunks[0].DataLength != 13)
        {
            throw new InvalidDataException("The encoder output does not start with a well-formed IHDR chunk.");
        }

        PngChunk ihdr = chunks[0];
        int width = PngChunkReader.ReadInt32BigEndian(source, ihdr.DataOffset);
        int height = PngChunkReader.ReadInt32BigEndian(source, ihdr.DataOffset + 4);
        if (width != expected.PixelWidth || height != expected.PixelHeight)
        {
            throw new InvalidDataException(
                $"The encoder output's IHDR dimensions ({width}x{height}) do not match the render ({expected.PixelWidth}x{expected.PixelHeight}).");
        }

        bool hasIdat = false;
        foreach (PngChunk chunk in chunks)
        {
            hasIdat |= chunk.Type == "IDAT";
        }

        if (!hasIdat)
        {
            throw new InvalidDataException("The encoder output contains no IDAT chunk.");
        }

        if (chunks[^1].Type != "IEND")
        {
            throw new InvalidDataException("The encoder output does not end with an IEND chunk.");
        }
    }

    private static byte[] CreateDensityChunk()
    {
        byte[] chunk = new byte[_densityChunkLength];

        // Chunk layout: four-byte length, four-byte type, nine data bytes, four-byte CRC.
        PngChunkReader.WriteUInt32BigEndian(chunk, 0, 9);
        chunk[4] = (byte)'p';
        chunk[5] = (byte)'H';
        chunk[6] = (byte)'Y';
        chunk[7] = (byte)'s';
        PngChunkReader.WriteUInt32BigEndian(chunk, 8, PixelsPerMetre);
        PngChunkReader.WriteUInt32BigEndian(chunk, 12, PixelsPerMetre);
        chunk[16] = 1;
        PngChunkReader.WriteUInt32BigEndian(chunk, 17, PngCrc32.Compute(chunk.AsSpan(4, 13)));

        return chunk;
    }
}
