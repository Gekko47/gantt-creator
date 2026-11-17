using System.Globalization;

namespace GanttCreator.AddIn;

/// <summary>
/// Turns a control id or a state slug into the picture the host displays.
/// </summary>
/// <remarks>
/// <para>
/// <b>The file name is the whole contract.</b> A glyph asset is
/// <c>slug-palette-size.png</c>, so the palette and the pixel size are visible in the
/// repository, in the generator and in a failure message, with no second mapping to keep
/// in step.
/// </para>
/// <para>
/// <b>Control identity comes from the generated catalogue.</b>
/// <see cref="RibbonGlyphCatalog"/> is emitted by the same script that draws the PNGs, so
/// a control cannot ask for a glyph the assets do not contain; an id that is not mapped
/// returns null, which is the state an image-less control is already in.
/// </para>
/// <para>
/// <b>Nothing here throws.</b> <see cref="RibbonImages"/> already returns null for a
/// missing or unreadable asset, and a ribbon callback must never surface an exception.
/// </para>
/// </remarks>
internal static class RibbonGlyphAssets
{
    private const string _lightPaletteName = "light";
    private const string _darkPaletteName = "dark";

    /// <summary>
    /// Builds the embedded file name for one glyph.
    /// </summary>
    /// <param name="slug">The glyph slug, e.g. <c>refresh-chart</c>.</param>
    /// <param name="palette">The palette to serve.</param>
    /// <param name="size">The pixel size: 16 or 32.</param>
    /// <returns>The file name, e.g. <c>refresh-chart-dark-32.png</c>.</returns>
    /// <exception cref="ArgumentException"><paramref name="slug"/> is null or whitespace.</exception>
    internal static string Suffix(string slug, RibbonGlyphPalette palette, int size)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(slug);

        string paletteName = palette == RibbonGlyphPalette.Dark
            ? _darkPaletteName
            : _lightPaletteName;

        return string.Create(
            CultureInfo.InvariantCulture,
            $"{slug}-{paletteName}-{size}.png");
    }

    /// <summary>
    /// Returns the picture for one control.
    /// </summary>
    /// <param name="controlId">The RibbonX control id, or null.</param>
    /// <returns>The picture, or null when the id is unmapped or the asset is unavailable.</returns>
    internal static object? ForControl(string? controlId) =>
        RibbonGlyphCatalog.TryGet(controlId, out RibbonGlyphCatalog.Entry entry)
            ? RibbonImages.For(Suffix(entry.Glyph, RibbonGlyphEnvironment.Current, entry.Size))
            : null;

    /// <summary>
    /// Returns the picture for one state glyph, whose size is fixed by its consumer.
    /// </summary>
    /// <param name="slug">The state glyph slug, e.g. <c>cell-lock-open</c>.</param>
    /// <param name="size">The pixel size: 16 or 32.</param>
    /// <returns>The picture, or null when the asset is unavailable.</returns>
    internal static object? ForState(string slug, int size) =>
        RibbonImages.For(Suffix(slug, RibbonGlyphEnvironment.Current, size));

    /// <summary>
    /// Reads one committed glyph, for the tests that prove the assets exist.
    /// </summary>
    /// <param name="slug">The glyph slug.</param>
    /// <param name="palette">The palette to read.</param>
    /// <param name="size">The pixel size.</param>
    /// <returns>The PNG bytes, or null when the asset is missing.</returns>
    internal static byte[]? ReadPng(string slug, RibbonGlyphPalette palette, int size) =>
        RibbonImages.ReadPng(Suffix(slug, palette, size));
}
