using System.Runtime.InteropServices;
using GanttCreator.Office;
using GanttCreator.Raster;

namespace GanttCreator.AddIn;

/// <summary>The typed outcome of one clipboard placement (R8.8 D8).</summary>
/// <param name="Succeeded">Whether the image was placed and verified present.</param>
/// <param name="Failure">The one actionable line when <paramref name="Succeeded"/> is false.</param>
internal sealed record PngClipboardOutcome(bool Succeeded, string? Failure)
{
    /// <summary>Builds a successful outcome.</summary>
    /// <returns>The success outcome.</returns>
    internal static PngClipboardOutcome Ok() => new(true, null);

    /// <summary>Builds a failed outcome.</summary>
    /// <param name="failure">The one actionable line.</param>
    /// <returns>The failure outcome.</returns>
    internal static PngClipboardOutcome Fail(string failure) => new(false, failure);
}

/// <summary>Places a rendered PNG on the clipboard (R8.8 D8, injectable for tests).</summary>
/// <param name="png">The encoded PNG bytes with their dimensions.</param>
/// <returns>The verified placement outcome.</returns>
internal delegate PngClipboardOutcome PngClipboardPlace(PngExportBytes png);

/// <summary>
/// The production clipboard port: a WinForms <see cref="Clipboard"/> placement on the calling
/// STA thread, verified through the R6.6 format reader (R8.8 D8).
/// </summary>
/// <remarks>
/// <para>
/// <b>STA, on the caller's thread.</b> <see cref="Clipboard.SetImage(Image)"/>
/// talks to the OLE clipboard through the calling thread's apartment. The ribbon command runs on
/// Excel's main STA thread, so no thread is created and nothing is marshalled (the
/// architecture's COM rule, and no <c>Task.Run</c>).
/// </para>
/// <para>
/// <b>Verified, never assumed.</b> A non-throwing placement is not success (the PPT transfer
/// rule applied to the clipboard): the formats are read back immediately and the placement
/// fails when no image format is present, exactly as Copy Editable verifies its own copy.
/// </para>
/// </remarks>
internal sealed class WindowsPngClipboardPort
{
    private readonly Func<IClipboardFormatReader> _readerFactory;

    /// <summary>Creates the port.</summary>
    /// <param name="readerFactory">Resolves the format reader used for the verification read.</param>
    internal WindowsPngClipboardPort(Func<IClipboardFormatReader>? readerFactory = null)
    {
        _readerFactory = readerFactory ?? (static () => new WindowsClipboardFormatReader());
    }

    /// <summary>Places the PNG on the clipboard and verifies the formats read back.</summary>
    /// <param name="png">The encoded PNG bytes with their dimensions.</param>
    /// <returns>The verified placement outcome.</returns>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Microsoft.Globalization",
        "CA1303:Do not pass literals as localized parameters",
        Justification = "User-facing export messages are not localized in this release.")]
    public PngClipboardOutcome Place(PngExportBytes png)
    {
        ArgumentNullException.ThrowIfNull(png);

        try
        {
            byte[] bytes = png.Bytes.ToArray();
            using MemoryStream stream = new(bytes, writable: false);
            using Bitmap bitmap = new(stream);
            Clipboard.SetImage(bitmap);
        }
        catch (ExternalException)
        {
            return PngClipboardOutcome.Fail(
                "The clipboard is busy, so the PNG was not copied. Close the other copying application and export again.");
        }
        catch (ArgumentException)
        {
            return PngClipboardOutcome.Fail(
                "The PNG could not be placed on the clipboard. Export to a file instead.");
        }

        ClipboardFormatRead read = _readerFactory().Read();
        if (!read.Succeeded || read.Snapshot is null)
        {
            return PngClipboardOutcome.Fail(
                "The PNG could not be verified on the clipboard. Export to a file instead.");
        }

        bool carriesImage = read.Snapshot.FormatNames.Any(name =>
            name.Contains("Bitmap", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "PNG", StringComparison.OrdinalIgnoreCase)
            || name.Contains("DIB", StringComparison.OrdinalIgnoreCase));
        return carriesImage
            ? PngClipboardOutcome.Ok()
            : PngClipboardOutcome.Fail(
                "The clipboard did not keep the PNG image. Export to a file instead.");
    }
}
