using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// The fixed inputs that make a raster render reproducible: the surface kind,
/// the antialias setting, the renderer library version, and the resolved font
/// identity.
/// </summary>
/// <param name="SurfaceKind">
/// The raster surface kind. Always <c>cpu-raster</c>; no GPU surface is used.
/// </param>
/// <param name="Antialias">Whether antialiasing is enabled. Always <see langword="true"/>.</param>
/// <param name="SkiaVersion">The SkiaSharp assembly version that produced the pixels.</param>
/// <param name="FontIdentity">
/// The resolved pinned-face identities this render drew text with, or
/// <c>none</c> for a text-free scene. The value names each used face's
/// PostScript name, <c>name</c>-table version, and design-unit denominator, so
/// a golden failure names its own font cause (ADR-0070 D4).
/// </param>
/// <remarks>
/// R8.2 D4 names the CPU raster surface and the pinned antialias setting as
/// the two inputs that could otherwise vary across machines; R8.3 extends the
/// descriptor with the font identity, which is the third input a text-bearing
/// render depends on. The library version is recorded beside them so a future
/// package bump is visible in the determinism metadata rather than inferred
/// from a failing golden comparison.
/// </remarks>
public sealed record RasterDeterminism(
    string SurfaceKind,
    bool Antialias,
    string SkiaVersion,
    string FontIdentity = "none")
{
    /// <summary>
    /// Gets the descriptor for the pinned settings used by
    /// <see cref="SceneRasterRenderer"/> for a text-free render.
    /// </summary>
    public static RasterDeterminism Pinned { get; } = new(
        "cpu-raster",
        true,
        typeof(SKSurface).Assembly.GetName().Version?.ToString() ?? "unknown");
}
