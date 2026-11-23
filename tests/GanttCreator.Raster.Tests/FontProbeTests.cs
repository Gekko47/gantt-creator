using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using SkiaSharp;
using Xunit.Abstractions;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// R8.2a's font-presence probe: it reports the exact faces the raster renderer
/// resolves on the reference machine, and it proves the fail-closed identity
/// contract that R8.3 implements.
/// </summary>
/// <remarks>
/// <para>
/// The probe is evidence, not a routine gate. The SkiaSharp default font
/// manager cannot see Office cloud fonts (the observed negative control is
/// <c>MatchFamily("Aptos") -> null</c> and
/// <c>FromFamilyName("Aptos") -> Segoe UI</c>), so the only faithful route is
/// the explicit <see cref="SKTypeface.FromFile(string)"/> loader plus the
/// identity check below. Every test carries
/// <c>[Trait("Category", "FontProbe")]</c>; the routine entry points exclude
/// that category, because a GitHub-hosted Windows runner has no Office cloud
/// fonts. The reference-machine run is the gate, and its record is quoted in
/// ADR-0070.
/// </para>
/// <para>
/// The pinned identities are the four files
/// <c>scripts/gen-font-advances.py</c> read to build
/// <c>FontAdvanceTable</c>: Aptos Regular/Bold and Aptos Narrow Regular/Bold.
/// The shared <c>UnitsPerEm</c> of 2048 is the cross-check against the
/// measurement authority. Each face's version is recorded separately because
/// the family's Regular and Bold faces differ (2.01 vs 2.00).
/// </para>
/// </remarks>
[Trait("Category", "FontProbe")]
public class FontProbeTests
{
    private const string CloudFontsRelativePath = @"Microsoft\FontCache\4\CloudFonts";
    private const int ExpectedUnitsPerEm = 2048;

    private static readonly PinnedFace[] _pinnedFaces =
    [
        new(
            Family: "Aptos",
            PostScriptName: "Aptos",
            FileName: "30153066857.ttf",
            NameTableVersion: "Version 2.01;O365",
            HeadRevision: 2.0100,
            Sha256: "95980114FCFD42F2F9C446DAE429B70582BF2F03097D68433EA9E7D85A49DA0B"),
        new(
            Family: "Aptos",
            PostScriptName: "Aptos-Bold",
            FileName: "31531277363.ttf",
            NameTableVersion: "Version 2.00;O365",
            HeadRevision: 2.0000,
            Sha256: "B4124A6C17E73189574D9E51ADDA221F236D1DF901F1BBCD57E9A09AE5B18599"),
        new(
            Family: "Aptos Narrow",
            PostScriptName: "Aptos-Narrow",
            FileName: "27970306811.ttf",
            NameTableVersion: "Version 2.01;O365",
            HeadRevision: 2.0100,
            Sha256: "DED7515AA064578485A09C11A798A05D0303BDACDD7F2DE92D0521F3BC7793BA"),
        new(
            Family: "Aptos Narrow",
            PostScriptName: "AptosNarrow-2bc0000-560000",
            FileName: "24328775610.ttf",
            NameTableVersion: "Version 2.00;O365",
            HeadRevision: 2.0000,
            Sha256: "D1E33B7FFA6C481063DF87516D0C2976DCE00C717663C7943C2742E3E4BCAA99"),
    ];

    private readonly ITestOutputHelper _output;

    /// <summary>Initialises the probe with the xUnit test output sink.</summary>
    /// <param name="output">The output sink the evidence record is written to.</param>
    public FontProbeTests(ITestOutputHelper output) => _output = output;

    /// <summary>
    /// Resolves each of the four pinned faces from the Office cloud-font cache
    /// and asserts the exact identity the renderer must consume.
    /// </summary>
    [Fact]
    public void The_reference_machine_resolves_the_four_pinned_faces()
    {
        string cacheRoot = CloudFontsRoot();
        Assert.True(
            Directory.Exists(cacheRoot),
            $"The Office cloud-font cache was not found at '{cacheRoot}'. The R8.2a probe runs on the reference machine that has Office-delivered Aptos.");

        foreach (PinnedFace pinned in _pinnedFaces)
        {
            string path = Path.Combine(cacheRoot, pinned.Family, pinned.FileName);
            Assert.True(
                File.Exists(path),
                $"The pinned face '{pinned.Family}/{pinned.FileName}' is missing. The renderer must fail closed rather than substitute another font.");

            ProbeResult result = Resolve(path, pinned);
            Assert.True(
                result.Succeeded,
                $"{pinned.Family}/{pinned.FileName} was refused as {result.Refusal}. Observed: {result.Observed}");

            FontIdentity identity = result.Identity!;
            Assert.Equal(pinned.Family, identity.Family);
            Assert.Equal(pinned.PostScriptName, identity.PostScriptName);
            Assert.Equal(pinned.NameTableVersion, identity.NameTableVersion);
            Assert.Equal(pinned.HeadRevision, identity.HeadRevision, precision: 4);
            Assert.Equal(ExpectedUnitsPerEm, identity.UnitsPerEm);
            Assert.Equal(pinned.Sha256, identity.Sha256);

            _output.WriteLine(identity.ToEvidenceLine(pinned));
        }
    }

    /// <summary>
    /// Proves the D1 rule: a face obtained without the explicit file route is
    /// never accepted as the pin, even when its family name matches.
    /// </summary>
    [Fact]
    public void A_system_font_lookup_is_never_accepted_as_the_pinned_face()
    {
        string[] families = SKFontManager.Default.GetFontFamilies();
        _output.WriteLine($"system font families={families.Length}; contains Aptos={families.Contains("Aptos")}");

        using SKTypeface? systemTypeface = SKTypeface.FromFamilyName("Aptos");
        Assert.NotNull(systemTypeface);

        FontIdentity observed = Observe(systemTypeface!, filePath: "<system>");
        PinnedFace pinned = _pinnedFaces[0];

        ProbeRefusal? refusal = Mismatch(pinned, observed);
        Assert.NotNull(refusal);
        Assert.False(
            Matches(pinned, observed),
            "A typeface obtained from the system font manager must never be accepted as the pinned face; the explicit file route and its hash are part of the identity.");
        _output.WriteLine($"negative control: FromFamilyName(\"Aptos\") -> {observed.Family}/{observed.PostScriptName}; refused as {refusal}");
    }

    /// <summary>
    /// A missing pinned face is a typed refusal, never a fallback.
    /// </summary>
    [Fact]
    public void A_missing_pinned_face_is_refused_with_a_typed_reason()
    {
        string missing = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "Aptos", "30153066857.ttf");
        Assert.False(File.Exists(missing));

        ProbeResult result = Resolve(missing, _pinnedFaces[0]);

        Assert.False(result.Succeeded);
        Assert.Equal(ProbeRefusal.FontFileMissing, result.Refusal);
        Assert.Null(result.Identity);
    }

    /// <summary>
    /// A face whose version identity differs from the pin is refused; this is
    /// the positive test for the version arm of the validator.
    /// </summary>
    [Fact]
    public void A_face_with_a_different_version_is_refused()
    {
        string cacheRoot = CloudFontsRoot();
        string path = Path.Combine(cacheRoot, _pinnedFaces[0].Family, _pinnedFaces[0].FileName);
        Assert.True(File.Exists(path));

        using SKTypeface? typeface = SKTypeface.FromFile(path);
        Assert.NotNull(typeface);

        FontIdentity actual = Observe(typeface!, path);
        FontIdentity changed = actual with { NameTableVersion = "Version 9.99;O365" };

        Assert.Equal<ProbeRefusal?>(ProbeRefusal.VersionMismatch, Mismatch(_pinnedFaces[0], changed));
        Assert.False(Matches(_pinnedFaces[0], changed));
    }

    /// <summary>
    /// The pinned faces share the advance table's design-unit denominator, so
    /// the measured label boxes and the drawn glyphs are the same unit system.
    /// </summary>
    [Fact]
    public void The_pinned_faces_share_the_advance_table_units_per_em()
    {
        string cacheRoot = CloudFontsRoot();
        foreach (PinnedFace pinned in _pinnedFaces)
        {
            string path = Path.Combine(cacheRoot, pinned.Family, pinned.FileName);
            Assert.True(File.Exists(path));

            using SKTypeface? typeface = SKTypeface.FromFile(path);
            Assert.NotNull(typeface);
            Assert.Equal(ExpectedUnitsPerEm, typeface.UnitsPerEm);
        }
    }

    private static string CloudFontsRoot() =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            CloudFontsRelativePath);

    private static ProbeResult Resolve(string path, PinnedFace pinned)
    {
        if (!File.Exists(path))
        {
            return ProbeResult.Failure(ProbeRefusal.FontFileMissing, observed: null);
        }

        using SKTypeface? typeface = SKTypeface.FromFile(path);
        if (typeface is null)
        {
            return ProbeResult.Failure(ProbeRefusal.LoadFailed, observed: null);
        }

        FontIdentity observed = Observe(typeface, path);
        ProbeRefusal? refusal = Mismatch(pinned, observed);
        return refusal is null
            ? ProbeResult.Success(observed)
            : ProbeResult.Failure(refusal.Value, observed);
    }

    private static bool Matches(PinnedFace pinned, FontIdentity observed) => Mismatch(pinned, observed) is null;

    private static ProbeRefusal? Mismatch(PinnedFace pinned, FontIdentity observed)
    {
        if (!string.Equals(pinned.Family, observed.Family, StringComparison.Ordinal))
        {
            return ProbeRefusal.FamilyMismatch;
        }

        if (!string.Equals(pinned.PostScriptName, observed.PostScriptName, StringComparison.Ordinal))
        {
            return ProbeRefusal.PostScriptMismatch;
        }

        if (!string.Equals(pinned.NameTableVersion, observed.NameTableVersion, StringComparison.Ordinal))
        {
            return ProbeRefusal.VersionMismatch;
        }

        if (Math.Abs(pinned.HeadRevision - observed.HeadRevision) > 0.00005)
        {
            return ProbeRefusal.VersionMismatch;
        }

        if (pinned.UnitsPerEm != observed.UnitsPerEm)
        {
            return ProbeRefusal.UnitsPerEmMismatch;
        }

        if (!string.Equals(pinned.Sha256, observed.Sha256, StringComparison.Ordinal))
        {
            return ProbeRefusal.VersionMismatch;
        }

        return null;
    }

    private static FontIdentity Observe(SKTypeface typeface, string filePath)
    {
        string? nameVersion = ReadNameTableVersion(typeface);
        double headRevision = ReadHeadRevision(typeface);

        return new FontIdentity(
            Family: typeface.FamilyName ?? string.Empty,
            PostScriptName: typeface.PostScriptName ?? string.Empty,
            NameTableVersion: nameVersion ?? "<parse-failure>",
            HeadRevision: headRevision,
            UnitsPerEm: typeface.UnitsPerEm,
            FilePath: filePath,
            Sha256: File.Exists(filePath) ? Sha256Of(filePath) : string.Empty);
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

    private sealed record PinnedFace(
        string Family,
        string PostScriptName,
        string FileName,
        string NameTableVersion,
        double HeadRevision,
        int UnitsPerEm = ExpectedUnitsPerEm,
        string Sha256 = "");

    private sealed record FontIdentity(
        string Family,
        string PostScriptName,
        string NameTableVersion,
        double HeadRevision,
        int UnitsPerEm,
        string FilePath,
        string Sha256)
    {
        public string ToEvidenceLine(PinnedFace pinned) =>
            $"{pinned.Family}/{pinned.FileName} family='{Family}' ps='{PostScriptName}' name='{NameTableVersion}' head={HeadRevision.ToString("0.0000", System.Globalization.CultureInfo.InvariantCulture)} upem={UnitsPerEm} sha256={Sha256}";
    }

    private enum ProbeRefusal
    {
        FontFileMissing,
        LoadFailed,
        FamilyMismatch,
        PostScriptMismatch,
        VersionMismatch,
        UnitsPerEmMismatch,
    }

    private readonly record struct ProbeResult(
        bool Succeeded,
        FontIdentity? Identity,
        ProbeRefusal? Refusal,
        FontIdentity? Observed)
    {
        public static ProbeResult Success(FontIdentity identity) =>
            new(Succeeded: true, identity, Refusal: null, Observed: identity);

        public static ProbeResult Failure(ProbeRefusal refusal, FontIdentity? observed) =>
            new(Succeeded: false, Identity: null, refusal, observed);
    }
}
