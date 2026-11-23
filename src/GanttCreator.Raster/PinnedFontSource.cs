using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// The typed refusals the pinned-font resolution can report (ADR-0070 D3).
/// </summary>
/// <remarks>
/// A refusal aborts the render before any pixel is drawn. The closed set is
/// the ADR's list; a new environment or identity fault is added here with a
/// positive test in the same commit.
/// </remarks>
public enum PinnedFontRefusal
{
    /// <summary>The pinned face file is absent from the cloud-font cache.</summary>
    FontFileMissing,

    /// <summary>The face file could not be loaded as a typeface.</summary>
    LoadFailed,

    /// <summary>
    /// The requested family is not one of the pinned families, or the loaded
    /// family differs from the pin.
    /// </summary>
    FamilyMismatch,

    /// <summary>The loaded PostScript name differs from the pin.</summary>
    PostScriptMismatch,

    /// <summary>The loaded <c>name</c>-table version, <c>head</c> revision, or SHA-256 differs from the pin.</summary>
    VersionMismatch,

    /// <summary>The loaded face's design-unit denominator differs from the advance table's.</summary>
    UnitsPerEmMismatch,

    /// <summary>The <c>name</c> or <c>head</c> table could not be read, so the face cannot be verified.</summary>
    ParseFailure,
}

/// <summary>
/// The identity observed from a loaded typeface: the members ADR-0070 D2 pins,
/// read from the face itself rather than trusted from a file name.
/// </summary>
/// <param name="Family">The typeface's family name.</param>
/// <param name="PostScriptName">The typeface's PostScript name.</param>
/// <param name="NameTableVersion">The <c>name</c> table's version string, or empty when unreadable.</param>
/// <param name="HeadRevision">The <c>head</c> revision, or <see cref="double.NaN"/> when unreadable.</param>
/// <param name="UnitsPerEm">The face's design-unit denominator.</param>
/// <param name="FilePath">The file the face was loaded from, or null for a system lookup.</param>
/// <param name="Sha256">The file's SHA-256, or empty when there is no file.</param>
public sealed record PinnedFontObservation(
    string Family,
    string PostScriptName,
    string NameTableVersion,
    double HeadRevision,
    int UnitsPerEm,
    string? FilePath,
    string Sha256)
{
    /// <summary>
    /// Returns the single-line evidence form the probe and the determinism
    /// metadata report (ADR-0070 D4).
    /// </summary>
    /// <returns>A culture-invariant evidence line naming every identity member.</returns>
    public string ToEvidenceLine() =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"{Family}/{PostScriptName} name='{NameTableVersion}' head={HeadRevision:0.0000} upem={UnitsPerEm} sha256={Sha256} file={FilePath ?? "<system>"}");
}

/// <summary>
/// The outcome of resolving one pinned face: the loaded typeface and its
/// observed identity on success, or the typed refusal and the offending
/// observation on failure.
/// </summary>
public sealed class PinnedFontResolution
{
    private PinnedFontResolution(
        bool succeeded,
        SKTypeface? typeface,
        PinnedFontRefusal? refusal,
        PinnedFontObservation? observed)
    {
        Succeeded = succeeded;
        Typeface = typeface;
        Refusal = refusal;
        Observed = observed;
    }

    /// <summary>Gets whether the face was resolved and verified.</summary>
    public bool Succeeded { get; }

    /// <summary>
    /// Gets the verified typeface on success; owned by the resolving
    /// <see cref="PinnedFontSource"/> and never disposed by a renderer.
    /// </summary>
    public SKTypeface? Typeface { get; }

    /// <summary>Gets the typed refusal on failure, or null on success.</summary>
    public PinnedFontRefusal? Refusal { get; }

    /// <summary>Gets the observed identity, when one could be read.</summary>
    public PinnedFontObservation? Observed { get; }

    /// <summary>Creates a success outcome.</summary>
    /// <param name="typeface">The verified typeface.</param>
    /// <param name="observed">The observed identity that matched the pin.</param>
    /// <returns>The success outcome.</returns>
    public static PinnedFontResolution Resolved(SKTypeface typeface, PinnedFontObservation observed) =>
        new(succeeded: true, typeface, refusal: null, observed);

    /// <summary>Creates a refusal outcome.</summary>
    /// <param name="refusal">The typed refusal.</param>
    /// <param name="observed">The offending observation, when one could be read.</param>
    /// <returns>The refusal outcome.</returns>
    public static PinnedFontResolution Refused(PinnedFontRefusal refusal, PinnedFontObservation? observed) =>
        new(succeeded: false, typeface: null, refusal, observed);
}

/// <summary>
/// The exception the raster renderer raises when a pinned face cannot be
/// resolved. It carries the typed refusal so a caller (or a test) can react to
/// the cause without parsing a message.
/// </summary>
public sealed class PinnedFontException : Exception
{
    /// <summary>Initialises an exception with no message.</summary>
    public PinnedFontException()
    {
    }

    /// <summary>Initialises an exception with a message.</summary>
    /// <param name="message">The message.</param>
    public PinnedFontException(string message)
        : base(message)
    {
    }

    /// <summary>Initialises an exception with a message and an inner exception.</summary>
    /// <param name="message">The message.</param>
    /// <param name="innerException">The cause.</param>
    public PinnedFontException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    /// <summary>Initialises an exception from a typed refusal.</summary>
    /// <param name="refusal">The typed refusal.</param>
    /// <param name="fontFamily">The scene style's family name.</param>
    /// <param name="observed">The offending observation, when one could be read.</param>
    public PinnedFontException(PinnedFontRefusal refusal, string fontFamily, PinnedFontObservation? observed)
        : base(BuildMessage(refusal, fontFamily, observed))
    {
        Refusal = refusal;
        FontFamily = fontFamily;
        Observed = observed;
    }

    /// <summary>Gets the typed refusal that aborted the render.</summary>
    public PinnedFontRefusal Refusal { get; }

    /// <summary>Gets the scene style's family name that could not be resolved.</summary>
    public string FontFamily { get; } = string.Empty;

    /// <summary>Gets the offending observation, when one could be read.</summary>
    public PinnedFontObservation? Observed { get; }

    private static string BuildMessage(PinnedFontRefusal refusal, string fontFamily, PinnedFontObservation? observed)
    {
        string detail = observed is null ? string.Empty : $" Observed: {observed.ToEvidenceLine()}.";
        return $"The pinned font face for '{fontFamily}' could not be resolved ({refusal}). The render was aborted rather than substituting another font.{detail}";
    }
}

/// <summary>
/// The single pinned-font resolver (ADR-0070 D3/D5): it loads each face from
/// the Office cloud-font cache through the explicit <see cref="SKTypeface.FromFile"/>
/// route, verifies its identity against <see cref="PinnedFontSet"/>, and caches
/// successful resolutions for the process.
/// </summary>
/// <remarks>
/// <para>
/// The system font manager is deliberately never consulted in this path: the
/// R8.2a probe proved <c>MatchFamily("Aptos")</c> returns null and
/// <c>FromFamilyName("Aptos")</c> silently substitutes Segoe UI, the
/// substitution ADR-0070 forbids.
/// </para>
/// <para>
/// The source owns the typefaces it caches. A renderer resolves through it and
/// draws with the returned typeface; it never disposes a cached face. The
/// process-wide <see cref="Default"/> instance lives for the process, which is
/// what makes "resolved once per process" true without a per-render cost.
/// </para>
/// </remarks>
/// <param name="cloudFontsRoot">
/// The cloud-font cache root; null or blank selects
/// <c>%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts</c>.
/// </param>
public sealed class PinnedFontSource(string? cloudFontsRoot = null) : IDisposable
{
    /// <summary>
    /// The Office cloud-font cache path relative to the local application-data
    /// folder: <c>%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts</c>.
    /// </summary>
    public const string CloudFontsRelativePath = @"Microsoft\FontCache\4\CloudFonts";

    private readonly Lock _gate = new();
    private readonly Dictionary<(string Family, bool Bold), PinnedFontResolution> _cache = [];
    private bool _disposed;

    /// <summary>Gets the process-wide default source over the Office cloud-font cache.</summary>
    public static PinnedFontSource Default { get; } = new();

    /// <summary>Gets the cloud-font cache root this source resolves from.</summary>
    public string CloudFontsRoot { get; } = string.IsNullOrWhiteSpace(cloudFontsRoot)
        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), CloudFontsRelativePath)
        : cloudFontsRoot;

    /// <summary>
    /// Resolves and verifies one pinned face for a scene style's family and weight.
    /// </summary>
    /// <param name="fontFamily">The resolved family name carried by the scene style.</param>
    /// <param name="bold">Whether the resolved style is bold.</param>
    /// <returns>The resolution outcome; a refusal never throws.</returns>
    /// <exception cref="ArgumentException">Thrown when <paramref name="fontFamily"/> is blank.</exception>
    /// <exception cref="ObjectDisposedException">Thrown when this source has been disposed.</exception>
    public PinnedFontResolution Resolve(string fontFamily, bool bold)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fontFamily);

        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            (string Family, bool Bold) key = (fontFamily, bold);
            if (_cache.TryGetValue(key, out PinnedFontResolution? cached))
            {
                return cached;
            }

            PinnedFontFace? pinned = PinnedFontSet.FindFace(fontFamily, bold);
            if (pinned is null)
            {
                return PinnedFontResolution.Refused(
                    PinnedFontRefusal.FamilyMismatch,
                    new PinnedFontObservation(
                        Family: fontFamily,
                        PostScriptName: string.Empty,
                        NameTableVersion: string.Empty,
                        HeadRevision: double.NaN,
                        UnitsPerEm: 0,
                        FilePath: null,
                        Sha256: string.Empty));
            }

            string path = Path.Combine(CloudFontsRoot, pinned.Family, pinned.FileName);
            if (!File.Exists(path))
            {
                return PinnedFontResolution.Refused(PinnedFontRefusal.FontFileMissing, observed: null);
            }

            SKTypeface? typeface = null;
            try
            {
                typeface = SKTypeface.FromFile(path);
                if (typeface is null)
                {
                    return PinnedFontResolution.Refused(PinnedFontRefusal.LoadFailed, observed: null);
                }

                PinnedFontObservation observed = Observe(typeface, path);
                PinnedFontRefusal? refusal = pinned.Match(observed);
                if (refusal is not null)
                {
                    return PinnedFontResolution.Refused(refusal.Value, observed);
                }

                PinnedFontResolution resolution = PinnedFontResolution.Resolved(typeface, observed);
                _cache[key] = resolution;

                // Ownership moves to the cache; the finally must not dispose it.
                typeface = null;
                return resolution;
            }
            finally
            {
                typeface?.Dispose();
            }
        }
    }

    /// <summary>
    /// Reads a loaded typeface's identity for comparison with a pin.
    /// </summary>
    /// <param name="typeface">The loaded typeface to observe.</param>
    /// <param name="filePath">
    /// The file the typeface was loaded from, when known; null for a system
    /// lookup, which then carries no SHA-256.
    /// </param>
    /// <returns>The observed identity; unreadable tables are recorded, not thrown.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="typeface"/> is null.</exception>
    public static PinnedFontObservation Observe(SKTypeface typeface, string? filePath = null)
    {
        ArgumentNullException.ThrowIfNull(typeface);

        return new PinnedFontObservation(
            Family: typeface.FamilyName ?? string.Empty,
            PostScriptName: typeface.PostScriptName ?? string.Empty,
            NameTableVersion: ReadNameTableVersion(typeface) ?? string.Empty,
            HeadRevision: ReadHeadRevision(typeface),
            UnitsPerEm: typeface.UnitsPerEm,
            FilePath: filePath,
            Sha256: filePath is not null && File.Exists(filePath) ? Sha256Of(filePath) : string.Empty);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            foreach (PinnedFontResolution resolution in _cache.Values)
            {
                resolution.Typeface?.Dispose();
            }

            _cache.Clear();
        }

        GC.SuppressFinalize(this);
    }

    private static string? ReadNameTableVersion(SKTypeface typeface)
    {
        if (!typeface.TryGetTableData(Tag("name"), out byte[]? name) || name is null || name.Length < 6)
        {
            return null;
        }

        ushort count = BinaryPrimitives.ReadUInt16BigEndian(name.AsSpan(2));
        ushort stringOffset = BinaryPrimitives.ReadUInt16BigEndian(name.AsSpan(4));

        for (int index = 0; index < count; index++)
        {
            int record = 6 + (index * 12);
            if (record + 12 > name.Length)
            {
                return null;
            }

            ushort platform = BinaryPrimitives.ReadUInt16BigEndian(name.AsSpan(record));
            ushort nameId = BinaryPrimitives.ReadUInt16BigEndian(name.AsSpan(record + 6));
            ushort length = BinaryPrimitives.ReadUInt16BigEndian(name.AsSpan(record + 8));
            ushort offset = BinaryPrimitives.ReadUInt16BigEndian(name.AsSpan(record + 10));

            if (nameId != 5)
            {
                continue;
            }

            int start = stringOffset + offset;
            if (start + length > name.Length)
            {
                return null;
            }

            ReadOnlySpan<byte> raw = name.AsSpan(start, length);
            string text = platform is 3 or 0
                ? Encoding.BigEndianUnicode.GetString(raw)
                : Encoding.Latin1.GetString(raw);

            if (text.Length > 0)
            {
                return text;
            }
        }

        return null;
    }

    private static double ReadHeadRevision(SKTypeface typeface)
    {
        if (!typeface.TryGetTableData(Tag("head"), out byte[]? head) || head is null || head.Length < 8)
        {
            return double.NaN;
        }

        int fixedRevision = BinaryPrimitives.ReadInt32BigEndian(head.AsSpan(4));
        return fixedRevision / 65536.0;
    }

    private static uint Tag(string value) =>
        ((uint)value[0] << 24) | ((uint)value[1] << 16) | ((uint)value[2] << 8) | value[3];

    private static string Sha256Of(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
}
