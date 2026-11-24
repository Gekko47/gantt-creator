using System.Text;
using SkiaSharp;
using Xunit;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// Host-free pins for R8.3's font authority: the pin data, the family/weight
/// mapping, the fail-closed identity comparison (one positive test per refusal
/// arm), and the resolver's environment refusals.
/// </summary>
/// <remarks>
/// Everything here runs on a machine without Office cloud fonts. The real-face
/// identity validation is <see cref="FontProbeTests"/> (reference machine,
/// <c>Category=FontProbe</c>); these tests pin the logic that decides it.
/// </remarks>
public class PinnedFontSourceTests
{
    private static readonly string[] _expectedFamilies = ["Aptos", "Aptos Narrow"];

    private static PinnedFontFace PinnedAptos => PinnedFontSet.FindFace("Aptos", bold: false)!;

    private static PinnedFontObservation MatchingObservation() => new(
        Family: PinnedAptos.Family,
        PostScriptName: PinnedAptos.PostScriptName,
        NameTableVersion: PinnedAptos.NameTableVersion,
        HeadRevision: PinnedAptos.HeadRevision,
        UnitsPerEm: PinnedAptos.UnitsPerEm,
        FilePath: "<pin>",
        Sha256: PinnedAptos.Sha256);

    [Fact]
    public void The_pin_holds_the_four_adr_0070_faces()
    {
        Assert.Equal(4, PinnedFontSet.Faces.Count);
        Assert.All(PinnedFontSet.Faces, face => Assert.Equal(2048, face.UnitsPerEm));
        Assert.All(PinnedFontSet.Faces, face => Assert.Matches("^[0-9A-F]{64}$", face.Sha256));
        Assert.Equal(
            4,
            PinnedFontSet.Faces.Select(face => face.PostScriptName).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(
            _expectedFamilies,
            PinnedFontSet.Faces.Select(face => face.Family).Distinct(StringComparer.Ordinal).OrderBy(f => f, StringComparer.Ordinal).ToArray());
    }

    [Theory]
    [InlineData("Aptos", false, "Aptos")]
    [InlineData("Aptos", true, "Aptos-Bold")]
    [InlineData("Aptos Narrow", false, "Aptos-Narrow")]
    [InlineData("Aptos Narrow", true, "AptosNarrow-2bc0000-560000")]
    public void FindFace_maps_the_family_and_weight(string family, bool bold, string expectedPostScriptName)
    {
        PinnedFontFace? face = PinnedFontSet.FindFace(family, bold);

        Assert.NotNull(face);
        Assert.Equal(expectedPostScriptName, face.PostScriptName);
    }

    [Fact]
    public void FindFace_returns_null_for_a_family_that_is_not_pinned()
    {
        Assert.Null(PinnedFontSet.FindFace("Segoe UI", bold: false));

        // The lookup is ordinal: a case-shifted family is not the pin.
        Assert.Null(PinnedFontSet.FindFace("aptos", bold: false));
    }

    [Fact]
    public void Match_accepts_the_exact_pinned_identity()
    {
        Assert.Null(PinnedAptos.Match(MatchingObservation()));
    }

    [Fact]
    public void Match_refuses_a_family_mismatch()
    {
        PinnedFontObservation observed = MatchingObservation() with { Family = "Segoe UI" };

        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.FamilyMismatch, PinnedAptos.Match(observed));
    }

    [Fact]
    public void Match_refuses_a_postscript_mismatch()
    {
        PinnedFontObservation observed = MatchingObservation() with { PostScriptName = "SegoeUI" };

        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.PostScriptMismatch, PinnedAptos.Match(observed));
    }

    [Fact]
    public void Match_refuses_a_name_table_version_mismatch()
    {
        PinnedFontObservation observed = MatchingObservation() with { NameTableVersion = "Version 9.99;O365" };

        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.VersionMismatch, PinnedAptos.Match(observed));
    }

    [Fact]
    public void Match_refuses_a_head_revision_mismatch()
    {
        PinnedFontObservation observed = MatchingObservation() with { HeadRevision = 9.9900 };

        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.VersionMismatch, PinnedAptos.Match(observed));
    }

    [Fact]
    public void Match_refuses_a_units_per_em_mismatch()
    {
        PinnedFontObservation observed = MatchingObservation() with { UnitsPerEm = 1000 };

        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.UnitsPerEmMismatch, PinnedAptos.Match(observed));
    }

    [Fact]
    public void Match_refuses_a_hash_mismatch()
    {
        PinnedFontObservation observed = MatchingObservation() with
        {
            Sha256 = "0000000000000000000000000000000000000000000000000000000000000000",
        };

        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.VersionMismatch, PinnedAptos.Match(observed));
    }

    [Fact]
    public void Match_refuses_an_observation_whose_tables_could_not_be_read()
    {
        PinnedFontObservation observed = MatchingObservation() with
        {
            NameTableVersion = string.Empty,
            HeadRevision = double.NaN,
        };

        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.ParseFailure, PinnedAptos.Match(observed));
    }

    [Fact]
    public void Match_rejects_a_null_observation()
    {
        _ = Assert.Throws<ArgumentNullException>(() => PinnedAptos.Match(null!));
    }

    [Fact]
    public void Resolve_refuses_a_missing_root()
    {
        string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        using PinnedFontSource source = new(root);

        PinnedFontResolution result = source.Resolve("Aptos", bold: false);

        Assert.False(result.Succeeded);
        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.FontFileMissing, result.Refusal);
        Assert.Null(result.Typeface);
    }

    [Fact]
    public void Resolve_refuses_a_family_that_is_not_pinned()
    {
        using PinnedFontSource source = new(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")));

        PinnedFontResolution result = source.Resolve("Segoe UI", bold: false);

        Assert.False(result.Succeeded);
        Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.FamilyMismatch, result.Refusal);
        Assert.Equal("Segoe UI", result.Observed!.Family);
    }

    [Fact]
    public void Resolve_maps_a_corrupt_face_file_to_a_load_failure()
    {
        string root = Directory.CreateTempSubdirectory("gantt-font-").FullName;
        try
        {
            string family = Directory.CreateDirectory(Path.Combine(root, "Aptos")).FullName;
            File.WriteAllBytes(
                Path.Combine(family, "30153066857.ttf"),
                Encoding.ASCII.GetBytes("this is not a font file"));

            using PinnedFontSource source = new(root);
            PinnedFontResolution result = source.Resolve("Aptos", bold: false);

            Assert.False(result.Succeeded);
            Assert.Equal<PinnedFontRefusal?>(PinnedFontRefusal.LoadFailed, result.Refusal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Resolve_rejects_a_blank_family()
    {
        using PinnedFontSource source = new(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")));

        _ = Assert.Throws<ArgumentException>(() => source.Resolve("   ", bold: false));
    }

    [Fact]
    public void Observe_rejects_a_null_typeface()
    {
        _ = Assert.Throws<ArgumentNullException>(() => PinnedFontSource.Observe(null!));
    }

    [Fact]
    public void Resolve_after_dispose_throws()
    {
        PinnedFontSource source = new(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")));
        source.Dispose();

        _ = Assert.Throws<ObjectDisposedException>(() => source.Resolve("Aptos", bold: false));
    }

    [Fact]
    public void Dispose_is_idempotent()
    {
        PinnedFontSource source = new(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")));

        source.Dispose();
        source.Dispose();
    }
}
