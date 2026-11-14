namespace GanttCreator.AddIn;

/// <summary>
/// The two item images the column dropdown shows (R5.18h): a tick when the column is
/// shown, a dash when it is hidden.
/// </summary>
/// <remarks>
/// <para>
/// <b>The glyphs are committed PNGs, rasterised from Lucide's geometry by
/// <c>scripts/gen-column-visibility-icons.ps1</c>.</b> RibbonX wants an
/// <c>IPictureDisp</c>, and the add-in deliberately carries no SVG renderer and no drawing
/// dependency of its own, so nothing renders a path at run time: <see cref="RibbonImages"/>
/// decodes an embedded resource and hands the result to the host.
/// </para>
/// <para>
/// <b>Two sizes are embedded, the 32 px one is served.</b> Office scales the picture it is
/// given, and a 32 px source stays crisp at both ribbon sizes and at 150 %/200 % display
/// scaling; the 16 px pair is committed beside it because the owner's ruling names both.
/// </para>
/// </remarks>
internal static class ColumnVisibilityImages
{
    /// <summary>The embedded suffix of the "column is shown" glyph.</summary>
    internal const string VisibleResourceSuffix = "column-visible-32.png";

    /// <summary>The embedded suffix of the "column is hidden" glyph.</summary>
    internal const string HiddenResourceSuffix = "column-hidden-32.png";

    /// <summary>
    /// Returns the picture for a visibility state.
    /// </summary>
    /// <param name="visible">Whether the column is currently shown.</param>
    /// <returns>The <c>IPictureDisp</c> wrapper, or null when the glyph is unavailable.</returns>
    internal static object? For(bool visible) =>
        RibbonImages.For(visible ? VisibleResourceSuffix : HiddenResourceSuffix);

    /// <summary>
    /// Reads one embedded PNG by its file-name suffix, for the tests that prove the assets are
    /// committed and are real PNGs.
    /// </summary>
    /// <param name="suffix">The file name, e.g. <c>column-visible-32.png</c>.</param>
    /// <returns>The PNG bytes, or null when no embedded resource ends with the suffix.</returns>
    internal static byte[]? ReadPng(string suffix) => RibbonImages.ReadPng(suffix);
}
