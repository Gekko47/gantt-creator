namespace GanttCreator.Raster;

/// <summary>
/// The calculated pixel dimensions of an export, derived from the export
/// composition's bounds in points at 300 DPI.
/// </summary>
/// <param name="PixelWidth">The rounded pixel width.</param>
/// <param name="PixelHeight">The rounded pixel height.</param>
public readonly record struct PixelDimensions(int PixelWidth, int PixelHeight);

/// <summary>
/// Converts the export composition's bounds in points to pixel dimensions at
/// 300 DPI.
/// </summary>
/// <remarks>
/// <para>
/// The PNG is the chart the user sees (owner ruling 2026-11-23), so there is no
/// width or unit input: the composition's own bounds are the size. The former
/// <c>cm</c>/<c>in</c>/<c>px</c> width grammar was withdrawn with the R8.7 control
/// surface it fed; this type is now the single points-to-pixels boundary.
/// </para>
/// <para>
/// Each dimension is rounded exactly once at the final boundary, through
/// <see cref="Math.Round(double)"/>'s default (<see cref="MidpointRounding.ToEven"/>),
/// which is the single-rounding-boundary rule
/// <c>docs/02-ARCHITECTURE.md</c> ("Coordinate and rounding policy") states.
/// </para>
/// </remarks>
public static class ExportSize
{
    /// <summary>Rendering resolution: 300 dots per inch.</summary>
    public const double Dpi = 300.0;

    /// <summary>Points per inch, the denominator of the points-to-pixels conversion.</summary>
    public const double PointsPerInch = 72.0;

    /// <summary>
    /// Practical upper bound for any single exported pixel dimension. This
    /// sits far below int.MaxValue but above every rasterizer's bitmap
    /// limit, so an absurd composition fails here with an actionable message
    /// instead of surfacing as a GDI+/Skia allocation failure at encode time.
    /// </summary>
    public const double MaxPixelDimension = 65_535;

    /// <summary>
    /// Maximum total pixel count (width × height) for an export. Sized as a
    /// tested safe allocation budget: the raster pipeline's 32-bit pixel
    /// format costs 4 bytes per pixel, so a full-frame buffer is 4 × pixel
    /// count. The 100,000,000-pixel budget keeps a full frame
    /// around 400 MB with headroom for the renderer's working buffers, while
    /// the per-axis <see cref="MaxPixelDimension"/> still governs each side.
    /// Square boundary: √100,000,000 = 10,000.
    /// </summary>
    public const long MaxTotalPixels = 100_000_000L;

    /// <summary>
    /// Converts the export composition's width and height in points to pixel
    /// dimensions at 300 DPI, rounding each dimension exactly once.
    /// </summary>
    /// <param name="sceneWidthPt">The composition width in points. Must be finite and greater than zero.</param>
    /// <param name="sceneHeightPt">The composition height in points. Must be finite and greater than zero.</param>
    /// <returns>The rounded pixel dimensions.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A dimension is not finite or not greater than zero, the computed pixel
    /// dimensions exceed the per-axis limit of <see cref="MaxPixelDimension"/> or
    /// the total pixel budget of <see cref="MaxTotalPixels"/>, or a sub-pixel
    /// extent rounds to a zero dimension.
    /// </exception>
    public static PixelDimensions ToPixelsFromPoints(double sceneWidthPt, double sceneHeightPt)
    {
        if (!double.IsFinite(sceneWidthPt) || sceneWidthPt <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sceneWidthPt), sceneWidthPt, "Scene width must be finite and greater than zero.");
        }

        if (!double.IsFinite(sceneHeightPt) || sceneHeightPt <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sceneHeightPt), sceneHeightPt, "Scene height must be finite and greater than zero.");
        }

        double pixelWidth = sceneWidthPt / PointsPerInch * Dpi;
        double pixelHeight = sceneHeightPt / PointsPerInch * Dpi;

        if (!double.IsFinite(pixelWidth) || pixelWidth > MaxPixelDimension)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sceneWidthPt), pixelWidth,
                $"Computed pixel width exceeds the practical export limit of {MaxPixelDimension} px.");
        }

        if (!double.IsFinite(pixelHeight) || pixelHeight <= 0 || pixelHeight > MaxPixelDimension)
        {
            // Zero/negative heights must be rejected here too: a zero height
            // would otherwise round to PixelDimensions with a zero height.
            throw new ArgumentOutOfRangeException(
                nameof(sceneHeightPt), pixelHeight,
                $"Computed pixel height must be finite, greater than zero, and within the practical export limit of {MaxPixelDimension} px.");
        }

        int pxWidth = (int)Math.Round(pixelWidth);
        int pxHeight = (int)Math.Round(pixelHeight);

        // Reject rounded dimensions below 1: a sub-pixel extent is positive and
        // finite pre-rounding but rounds to 0, which must not produce a
        // zero-pixel export.
        if (pxWidth < 1 || pxHeight < 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sceneWidthPt),
                $"Rounded pixel dimensions must each be at least 1 px (got {pxWidth}x{pxHeight}).");
        }

        // Combined allocation budget: reject dimension pairs whose total pixel
        // count exceeds the budget before rasterization, even when each axis
        // is within the per-axis MaxPixelDimension limit.
        long totalPixels = (long)pxWidth * pxHeight;
#pragma warning disable IDE0046
        if (totalPixels > MaxTotalPixels)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sceneHeightPt), totalPixels,
                $"Combined pixel dimensions ({pxWidth}x{pxHeight} = {totalPixels} px) exceed the maximum total of {MaxTotalPixels} pixels.");
        }
#pragma warning restore IDE0046

        return new PixelDimensions(pxWidth, pxHeight);
    }
}
