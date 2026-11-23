using GanttCreator.Core;

namespace GanttCreator.Raster;

/// <summary>
/// The explicit canvas background for a raster render: either fully
/// transparent or one resolved colour.
/// </summary>
/// <remarks>
/// <c>docs/02-ARCHITECTURE.md</c> ("PNG renderer") requires the background
/// transparency or colour to be explicit, so this type offers no implicit
/// default beyond the named <see cref="Transparent"/> value. The renderer
/// never derives a background from the scene or from a host setting.
/// </remarks>
public sealed record RasterBackground
{
    private RasterBackground(ColourHex? fill)
    {
        Fill = fill;
    }

    /// <summary>
    /// Gets the background fill colour, or <see langword="null"/> for a
    /// transparent canvas.
    /// </summary>
    public ColourHex? Fill { get; }

    /// <summary>Gets a value indicating whether the canvas is left transparent.</summary>
    public bool IsTransparent => Fill is null;

    /// <summary>Gets the transparent background.</summary>
    public static RasterBackground Transparent { get; } = new(fill: null);

    /// <summary>Creates a background filled with the supplied colour.</summary>
    /// <param name="colour">The opaque or translucent fill colour.</param>
    /// <returns>The coloured background.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="colour"/> is null.</exception>
    public static RasterBackground FromColour(ColourHex colour)
    {
        ArgumentNullException.ThrowIfNull(colour);
        return new(colour);
    }
}
