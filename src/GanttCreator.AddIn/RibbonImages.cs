using System.Collections.Concurrent;
using System.Runtime.InteropServices;

namespace GanttCreator.AddIn;

/// <summary>
/// The one place a committed PNG becomes the picture type RibbonX can display.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a picture at all.</b> A callback that hands the host an image returns an
/// <c>IPictureDisp</c>. Office cannot be given a file path from a ribbon callback, and the
/// add-in deliberately carries no SVG renderer, so every glyph is rasterised once by
/// <c>scripts/gen-column-visibility-icons.ps1</c>, committed under <c>Resources/</c>, and only
/// decoded at run time.
/// </para>
/// <para>
/// <b>One picture per suffix, cached for the process.</b> Excel asks for an image on every
/// query it makes — every dropdown open, every ribbon invalidation — and building a fresh OLE
/// wrapper per call would allocate and leak a COM object for nothing. The cache is keyed by the
/// resource suffix, so adding a second glyph family costs no second cache.
/// </para>
/// <para>
/// <b>Failures return null, never throw.</b> A missing resource, a corrupt file, or a host that
/// refuses the conversion leaves the control without an image, which is the state every
/// <c>imageMso</c>-less control is already in. An exception raised inside a ribbon callback
/// takes the whole ribbon down, so no failure is allowed through.
/// </para>
/// </remarks>
internal static class RibbonImages
{
    private static readonly ConcurrentDictionary<string, Lazy<object?>> _cache =
        new(StringComparer.Ordinal);

    /// <summary>
    /// Returns the picture for one embedded glyph.
    /// </summary>
    /// <param name="suffix">The file name, e.g. <c>lock-locked-32.png</c>.</param>
    /// <returns>The <c>IPictureDisp</c> wrapper, or null when the glyph is unavailable.</returns>
    internal static object? For(string suffix)
    {
        ArgumentNullException.ThrowIfNull(suffix);

        return _cache.GetOrAdd(suffix, key => new Lazy<object?>(() => Build(key))).Value;
    }

    /// <summary>
    /// Reads one embedded PNG by its file-name suffix.
    /// </summary>
    /// <param name="suffix">The file name, e.g. <c>lock-locked-32.png</c>.</param>
    /// <returns>The PNG bytes, or null when no embedded resource ends with the suffix.</returns>
    /// <remarks>
    /// Matched by suffix rather than by an exact logical name: the manifest name depends on the
    /// resource's folder and on the project's naming rules, so an exact name would be a second
    /// place to keep in step with the .csproj. The suffix is the file's own name.
    /// </remarks>
    internal static byte[]? ReadPng(string suffix)
    {
        ArgumentNullException.ThrowIfNull(suffix);

        foreach (string name in typeof(RibbonImages).Assembly.GetManifestResourceNames())
        {
            if (!name.EndsWith(suffix, StringComparison.Ordinal))
            {
                continue;
            }

            using Stream? stream = typeof(RibbonImages).Assembly.GetManifestResourceStream(name);
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
            // A corrupt or non-image resource: the control simply carries no glyph.
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
    private sealed class PictureHost() : AxHost(string.Empty)
    {
        /// <summary>Converts a bitmap to the OLE picture the host displays.</summary>
        /// <param name="bitmap">The decoded glyph.</param>
        /// <returns>The <c>IPictureDisp</c> wrapper.</returns>
        internal static object? GetPicture(Bitmap bitmap) => GetIPictureDispFromPicture(bitmap);
    }
}
