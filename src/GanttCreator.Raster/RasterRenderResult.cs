using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// The pixel output of one raster render, owned by the caller.
/// </summary>
/// <remarks>
/// <para>
/// The result wraps the render's pixel buffer without copying it: a legal
/// 100M-pixel frame must not be duplicated in managed memory (R8.2 I5). The
/// pixel layout is BGRA, top-down, matching the pinned
/// <see cref="SKColorType.Bgra8888"/> surface; consumers encode or compare
/// these bytes and must not assume any other order.
/// </para>
/// <para>
/// The bitmap is created by this type and disposed by <see cref="Dispose()"/>,
/// so the renderer's exception path has exactly one owner to release and the
/// success path transfers that owner to the caller.
/// </para>
/// </remarks>
public sealed class RasterRenderResult : IDisposable
{
    private bool _disposed;

    internal RasterRenderResult(PixelDimensions dimensions, RasterDeterminism determinism)
    {
        Bitmap = new SKBitmap(
            new SKImageInfo(
                dimensions.PixelWidth,
                dimensions.PixelHeight,
                SKColorType.Bgra8888,
                SKAlphaType.Premul));
        PixelWidth = dimensions.PixelWidth;
        PixelHeight = dimensions.PixelHeight;
        Determinism = determinism;
    }

    /// <summary>Gets the canvas width in pixels.</summary>
    public int PixelWidth { get; }

    /// <summary>Gets the canvas height in pixels.</summary>
    public int PixelHeight { get; }

    /// <summary>Gets the pinned settings that make this render reproducible.</summary>
    public RasterDeterminism Determinism { get; }

    /// <summary>Gets the raw BGRA pixel bytes, top-down, without copying.</summary>
    public ReadOnlySpan<byte> Pixels => Bitmap.GetPixelSpan();

    /// <summary>Gets the underlying bitmap so the renderer can draw into it.</summary>
    internal SKBitmap Bitmap { get; }

    /// <inheritdoc />
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    private void Dispose(bool disposing)
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        if (disposing)
        {
            Bitmap.Dispose();
        }
    }
}
