namespace GanttCreator.Raster;

/// <summary>The complete PNG file bytes and the pixel dimensions their <c>IHDR</c> declares.</summary>
/// <param name="Bytes">The PNG file bytes, including the signature and the <c>pHYs</c> chunk.</param>
/// <param name="Dimensions">The exact pixel dimensions the encoder wrote into <c>IHDR</c>.</param>
/// <remarks>
/// The density is part of the encoder's contract (11811 pixels per metre, metre
/// unit) and is re-checked by the validator; it is deliberately not duplicated
/// as per-instance DPI fields, so a second density authority cannot drift.
/// </remarks>
public sealed record PngExportBytes(ReadOnlyMemory<byte> Bytes, PixelDimensions Dimensions);
