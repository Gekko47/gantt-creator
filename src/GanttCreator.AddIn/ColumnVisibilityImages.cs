using System.Runtime.InteropServices;

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
/// dependency of its own, so nothing renders a path at run time: this type decodes an
/// embedded resource and hands the result to the host.
/// </para>
/// <para>
/// <b>Two sizes are embedded, the 32 px one is served.</b> Office scales the picture it is
/// given, and a 32 px source stays crisp at both ribbon sizes and at 150 %/200 % display
/// scaling; the 16 px pair is committed beside it because the owner's ruling names both,
/// and because a future per-DPI path must not have to re-derive them.
/// </para>
/// <para>
/// <b>One picture per state, cached for the process.</b> Excel asks for an item image on
/// every dropdown query, and building a fresh OLE picture per call would allocate (and
/// leak) a COM wrapper for nothing. The instance is immutable from the host's side.
/// </para>
/// <para>
/// <b>Failures return null, never throw.</b> A missing resource or a host that refuses the
/// conversion leaves the item without an image, which is the state every other item in the
/// ribbon is already in; a thrown exception inside a callback takes the whole ribbon down.
/// </para>
/// </remarks>
internal static class ColumnVisibilityImages
{
    /// <summary>The embedded suffix of the "column is shown" glyph.</summary>
    internal const string VisibleResourceSuffix = "column-visible-32.png";

    /// <summary>The embedded suffix of the "column is hidden" glyph.</summary>
    internal const string HiddenResourceSuffix = "column-hidden-32.png";

    private static readonly Lazy<object?> _visible = new(() => Build(VisibleResourceSuffix));

    private static readonly Lazy<object?> _hidden = new(() => Build(HiddenResourceSuffix));

    /// <summary>
    /// Returns the picture for a visibility state.
    /// </summary>
    /// <param name="visible">Whether the column is currently shown.</param>
    /// <returns>The <c>IPictureDisp</c> wrapper, or null when the glyph is unavailable.</returns>
    internal static object? For(bool visible) => (visible ? _visible : _hidden).Value;

    /// <summary>
    /// Reads one embedded PNG by its file-name suffix.
    /// </summary>
    /// <param name="suffix">The file name, e.g. <c>column-visible-32.png</c>.</param>
    /// <returns>The PNG bytes, or null when no embedded resource ends with the suffix.</returns>
    /// <remarks>
    /// Matched by suffix rather than by an exact logical name: the manifest name depends on
    /// the resource's folder and on the project's naming rules, so an exact name would be a
    /// second place to keep in step with the .csproj. The suffix is the file's own name.
    /// </remarks>
    internal static byte[]? ReadPng(string suffix)
    {
        ArgumentNullException.ThrowIfNull(suffix);

        foreach (string name in typeof(ColumnVisibilityImages).Assembly.GetManifestResourceNames())
        {
            if (!name.EndsWith(suffix, StringComparison.Ordinal))
            {
                continue;
            }

            using Stream? stream = typeof(ColumnVisibilityImages).Assembly.GetManifestResourceStream(name);
            if (stream is null)
            {
                continue;
            }

            using MemoryStream buffer = new();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }

        return null;
    }

    /// <summary>Decodes one PNG and wraps it as the host's picture type.</summary>
    /// <param name="suffix">The embedded file name.</param>
    /// <returns>The picture, or null when the bytes or the conversion are unavailable.</returns>
    private static object? Build(string suffix)
    {
        byte[]? png = ReadPng(suffix);
        if (png is null)
        {
            return null;
        }

        try
        {
            using MemoryStream stream = new(png, writable: false);
            using Bitmap bitmap = new(stream);
            return PictureHost.GetPicture(bitmap);
        }
        catch (ArgumentException)
        {
            // A corrupt or non-image resource: the item simply carries no glyph.
            return null;
        }
        catch (ExternalException)
        {
            // The host refused the picture conversion.
            return null;
        }
    }

    /// <summary>
    /// Reaches <c>AxHost.GetIPictureDispFromPicture</c>, which the framework declares
    /// <c>protected</c>.
    /// </summary>
    /// <remarks>
    /// The subclass is the supported way to obtain an <c>IPictureDisp</c> from a managed
    /// image: Excel-DNA's own ribbon helpers call the same member the same way, and the
    /// alternative — hand-marshalling an OLE picture — is a COM structure to get wrong.
    /// The instance is never created; only the static accessor is used.
    /// </remarks>
    private sealed class PictureHost() : AxHost(string.Empty)
    {
        /// <summary>Wraps a managed image as the host's picture type.</summary>
        /// <param name="image">The image to convert.</param>
        /// <returns>The <c>IPictureDisp</c> wrapper.</returns>
        internal static object GetPicture(Image image) =>
            GetIPictureDispFromPicture(image) ?? throw new InvalidOperationException(
                "The framework returned no picture for the image.");
    }
}
