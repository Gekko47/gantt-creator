using System.Text;
using SkiaSharp;
using Xunit.Abstractions;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// R8.2a's font-presence probe, consuming R8.3's production pin: it reports the
/// exact faces the raster renderer resolves on the reference machine and proves
/// the fail-closed identity contract the renderer depends on.
/// </summary>
/// <remarks>
/// <para>
/// The probe is evidence, not a routine gate. The SkiaSharp default font
/// manager cannot see Office cloud fonts (the observed negative control is
/// <c>MatchFamily("Aptos") -> null</c> and
/// <c>FromFamilyName("Aptos") -> Segoe UI</c>), so the only faithful route is
/// the explicit <see cref="SKTypeface.FromFile(string)"/> loader plus the
/// identity check. Every test carries
/// <c>[Trait("Category", "FontProbe")]</c>; the routine entry points exclude
/// that category, because a GitHub-hosted Windows runner has no Office cloud
/// fonts. The reference-machine run is the gate, and its record is quoted in
/// ADR-0070.
/// </para>
/// <para>
/// Since R8.3 the probe and the renderer share
/// <see cref="PinnedFontSet"/> and <see cref="PinnedFontSource"/>, so a probe
/// that disagrees with the renderer is structurally impossible (ADR-0070 D5).
/// The shared <c>UnitsPerEm</c> of 2048 is the cross-check against the
/// measurement authority; each face's version is recorded separately because
/// the family's Regular and Bold faces differ (2.01 vs 2.00).
/// </para>
/// </remarks>
[Trait("Category", "FontProbe")]
public class FontProbeTests
{
    private const int ExpectedUnitsPerEm = 2048;

    private readonly ITestOutputHelper _output;

    /// <summary>Initialises the probe with the xUnit test output sink.</summary>
    /// <param name="output">The output sink the evidence record is written to.</param>
    public FontProbeTests(ITestOutputHelper output) => _output = output;

    /// <summary>
    /// Resolves each of the four pinned faces through the production resolver
    /// and asserts the exact identity the renderer consumes.
    /// </summary>
    [Fact]
    public void The_reference_machine_resolves_the_four_pinned_faces()
    {
        using PinnedFontSource source = new();
        Assert.True(
            Directory.Exists(source.CloudFontsRoot),
            $"The Office cloud-font cache was not found at '{source.CloudFontsRoot}'. The R8.2a probe runs on the reference machine that has Office-delivered Aptos.");

        foreach (PinnedFontFace pinned in PinnedFontSet.Faces)
        {
            PinnedFontResolution result = source.Resolve(pinned.Family, pinned.IsBold);
            Assert.True(
                result.Succeeded,
                $"{pinned.Family} bold={pinned.IsBold} was refused as {result.Refusal}. Observed: {result.Observed?.ToEvidenceLine()}");

            PinnedFontObservation observed = result.Observed!;
            Assert.Equal(pinned.Family, observed.Family);
            Assert.Equal(pinned.PostScriptName, observed.PostScriptName);
            Assert.Equal(pinned.NameTableVersion, observed.NameTableVersion);
            Assert.Equal(pinned.HeadRevision, observed.HeadRevision, precision: 4);
            Assert.Equal(ExpectedUnitsPerEm, observed.UnitsPerEm);
            Assert.Equal(pinned.Sha256, observed.Sha256);
            Assert.Null(pinned.Match(observed));

            _output.WriteLine(observed.ToEvidenceLine());
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

        PinnedFontObservation observed = PinnedFontSource.Observe(systemTypeface);
        PinnedFontFace pinned = PinnedFontSet.FindFace("Aptos", bold: false)!;

        PinnedFontRefusal? refusal = pinned.Match(observed);
        Assert.NotNull(refusal);
        Assert.False(
            refusal is null,
            "A typeface obtained from the system font manager must never be accepted as the pinned face; the explicit file route and its hash are part of the identity.");
        _output.WriteLine($"negative control: FromFamilyName(\"Aptos\") -> {observed.Family}/{observed.PostScriptName}; refused as {refusal}");
    }

    /// <summary>
    /// A missing pinned face is a typed refusal, never a fallback.
    /// </summary>
    [Fact]
    public void A_missing_pinned_face_is_refused_with_a_typed_reason()
    {
        string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Assert.False(Directory.Exists(root));

        using PinnedFontSource source = new(root);
        PinnedFontResolution result = source.Resolve("Aptos", bold: false);

        Assert.False(result.Succeeded);
        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.FontFileMissing, result.Refusal);
        Assert.Null(result.Typeface);
    }

    /// <summary>
    /// A face whose version identity differs from the pin is refused; this is
    /// the positive test for the version arm of the validator.
    /// </summary>
    [Fact]
    public void A_face_with_a_different_version_is_refused()
    {
        using PinnedFontSource source = new();
        PinnedFontFace pinned = PinnedFontSet.FindFace("Aptos", bold: false)!;
        PinnedFontResolution real = source.Resolve("Aptos", bold: false);
        Assert.True(real.Succeeded, $"The reference face was refused: {real.Refusal}.");

        PinnedFontObservation changed = real.Observed! with { NameTableVersion = "Version 9.99;O365" };

        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.VersionMismatch, pinned.Match(changed));
    }

    /// <summary>
    /// The pinned faces share the advance table's design-unit denominator, so
    /// the measured label boxes and the drawn glyphs are the same unit system.
    /// </summary>
    [Fact]
    public void The_pinned_faces_share_the_advance_table_units_per_em()
    {
        using PinnedFontSource source = new();
        foreach (PinnedFontFace pinned in PinnedFontSet.Faces)
        {
            PinnedFontResolution result = source.Resolve(pinned.Family, pinned.IsBold);
            Assert.True(result.Succeeded, $"{pinned.Family} bold={pinned.IsBold} was refused: {result.Refusal}.");
            Assert.Equal(ExpectedUnitsPerEm, result.Typeface!.UnitsPerEm);
        }
    }

    /// <summary>
    /// The resolver caches a verified face: a repeated resolution returns the
    /// same typeface instance, which is what "resolved once per process" means
    /// (ADR-0070 D2/D5).
    /// </summary>
    [Fact]
    public void The_resolver_returns_the_same_typeface_for_a_repeated_resolution()
    {
        using PinnedFontSource source = new();
        PinnedFontResolution first = source.Resolve("Aptos", bold: false);
        PinnedFontResolution second = source.Resolve("Aptos", bold: false);

        Assert.True(first.Succeeded);
        Assert.True(second.Succeeded);
        Assert.Same(first.Typeface, second.Typeface);
    }

    /// <summary>
    /// The glyph-coverage stop condition (work item Step 0): every pinned face
    /// contains every character the renderer's own text emitters produce —
    /// printable ASCII plus the U+2026 ellipsis the label truncation writes.
    /// </summary>
    [Fact]
    public void The_pinned_faces_contain_every_glyph_the_renderer_emits()
    {
        string required = RequiredGlyphSample();
        using PinnedFontSource source = new();

        foreach (PinnedFontFace pinned in PinnedFontSet.Faces)
        {
            PinnedFontResolution result = source.Resolve(pinned.Family, pinned.IsBold);
            Assert.True(result.Succeeded, $"{pinned.Family} bold={pinned.IsBold} was refused: {result.Refusal}.");
            Assert.True(
                ContainsEveryGlyph(result.Typeface!, required),
                $"{pinned.Family} bold={pinned.IsBold} lacks at least one glyph the renderer emits.");
        }
    }

    private static string RequiredGlyphSample()
    {
        StringBuilder builder = new();
        for (int codePoint = 0x20; codePoint <= 0x7E; codePoint++)
        {
            _ = builder.Append((char)codePoint);
        }

        // LabelText.Ellipsize writes this single-character ellipsis.
        _ = builder.Append('\u2026');
        return builder.ToString();
    }

    private static bool ContainsEveryGlyph(SKTypeface typeface, string sample)
    {
        SKFont? font = null;
        try
        {
            font = new SKFont(typeface, 12f, 1f, 0f);
            return font.ContainsGlyphs(sample);
        }
        finally
        {
            font?.Dispose();
        }
    }
}
