using GanttCreator.Core;

namespace GanttCreator.Raster;

/// <summary>
/// The single points-to-pixels mapping for one render: one origin and one
/// per-axis scale factor, with every coordinate rounded exactly once.
/// </summary>
/// <remarks>
/// <para>
/// <c>docs/02-ARCHITECTURE.md</c> ("Coordinate and rounding policy") fixes
/// <c>pixels = round(points / 72 * dpi * scale)</c> and requires the edge
/// rounding to be chosen once. R8.2's decision I1 reads that formula against
/// the passed <see cref="PixelDimensions"/>: the X and Y factors are derived
/// from the dimensions <see cref="ExportSize.ToPixels"/> already rounded, so
/// the canvas is filled exactly instead of a shared scale being re-derived and
/// landing one pixel outside it.
/// </para>
/// <para>
/// The rounding mode is <see cref="MidpointRounding.ToEven"/>, the default
/// mode <see cref="Math.Round(double)"/> uses and the same one
/// <see cref="ExportSize"/> uses for its dimension rounding, so the two
/// boundaries agree.
/// </para>
/// </remarks>
internal sealed class RasterTransform
{
    private readonly double _originX;
    private readonly double _originY;
    private readonly double _scaleX;
    private readonly double _scaleY;

    private RasterTransform(double originX, double originY, double scaleX, double scaleY)
    {
        _originX = originX;
        _originY = originY;
        _scaleX = scaleX;
        _scaleY = scaleY;
    }

    /// <summary>Creates the mapping for a scene bounds and validated pixel dimensions.</summary>
    /// <param name="sceneBounds">The scene bounds the canvas crops to, in points.</param>
    /// <param name="dimensions">The validated pixel dimensions of the canvas.</param>
    /// <returns>The mapping.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when the scene bounds are not finite and positive or a pixel
    /// dimension is not positive.
    /// </exception>
    public static RasterTransform Create(RectD sceneBounds, PixelDimensions dimensions)
    {
        // RectD's constructor already guarantees finite coordinates and extents,
        // so only a non-positive extent is reachable here.
        if (sceneBounds.Width <= 0 || sceneBounds.Height <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sceneBounds),
                sceneBounds,
                "Scene bounds must have a positive width and height.");
        }

        ArgumentOutOfRangeException.ThrowIfLessThan(dimensions.PixelWidth, 1, nameof(dimensions));
        ArgumentOutOfRangeException.ThrowIfLessThan(dimensions.PixelHeight, 1, nameof(dimensions));

        return new RasterTransform(
            sceneBounds.X,
            sceneBounds.Y,
            dimensions.PixelWidth / sceneBounds.Width,
            dimensions.PixelHeight / sceneBounds.Height);
    }

    /// <summary>Maps a scene X coordinate to a pixel coordinate.</summary>
    /// <param name="pointX">The scene X coordinate in points.</param>
    /// <returns>The pixel X coordinate.</returns>
    public float X(double pointX) => (float)Math.Round((pointX - _originX) * _scaleX);

    /// <summary>Maps a scene Y coordinate to a pixel coordinate.</summary>
    /// <param name="pointY">The scene Y coordinate in points.</param>
    /// <returns>The pixel Y coordinate.</returns>
    public float Y(double pointY) => (float)Math.Round((pointY - _originY) * _scaleY);

    /// <summary>Maps a scene length in points to a pixel length.</summary>
    /// <param name="lengthPt">The length in points.</param>
    /// <returns>The pixel length, never smaller than one pixel.</returns>
    public float Length(double lengthPt) => Math.Max(1f, (float)Math.Round(lengthPt * _scaleX));
}
