using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// The fixed inputs that make a raster render reproducible: the surface kind,
/// the antialias setting, and the renderer library version.
/// </summary>
/// <param name="SurfaceKind">
/// The raster surface kind. Always <c>cpu-raster</c>; no GPU surface is used.
/// </param>
/// <param name="Antialias">Whether antialiasing is enabled. Always <see langword="true"/>.</param>
/// <param name="SkiaVersion">The SkiaSharp assembly version that produced the pixels.</param>
/// <remarks>
/// R8.2 D4 names the CPU raster surface and the pinned antialias setting as
/// the two inputs that could otherwise vary across machines. The library
/// version is recorded beside them so a future package bump is visible in the
/// determinism metadata rather than inferred from a failing golden comparison.
/// </remarks>
public sealed record RasterDeterminism(string SurfaceKind, bool Antialias, string SkiaVersion)
{
    /// <summary>
    /// Gets the descriptor for the pinned settings used by
    /// <see cref="SceneRasterRenderer"/>.
    /// </summary>
    public static RasterDeterminism Pinned { get; } = new(
        "cpu-raster",
        true,
        typeof(SKSurface).Assembly.GetName().Version?.ToString() ?? "unknown");
}
