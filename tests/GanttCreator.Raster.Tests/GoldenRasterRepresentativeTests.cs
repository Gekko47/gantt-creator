using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using SkiaSharp;
using Xunit;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// The R8.4 representative golden baselines: the committed reference scene rendered
/// at the five owner-ratified widths (2026-11-23), each with a metadata-pinned
/// SHA-256 digest of the exact BGRA buffer, a human-review PNG, and
/// scene-derived hard-geometry interior probes.
/// </summary>
/// <remarks>
/// <para>
/// <b>Fixture route.</b> The scene is the committed
/// <c>tests/golden/scene/reference-scene.json</c> loaded through the public
/// <see cref="SceneSnapshot.Deserialize"/>; the internal Core test fixture is not
/// referenced, so one byte-pinned input serves the Core snapshot guard and this
/// suite (Phase-8 review Revision A).
/// </para>
/// <para>
/// <b>Storage ruling (2026-11-23).</b> The baseline artefact is a
/// <c>representative-<key>.meta</c> sidecar plus the human-review
/// <c>representative-<key>.png</c>. The sidecar's SHA-256 digest is the exact
/// whole-buffer gate; the five raw BGRA buffers would total ≈142 MB and every
/// regeneration would re-commit that size, so the buffers are not committed.
/// </para>
/// <para>
/// <b>Exact where geometry is exact.</b> The digest gates every pixel; the region
/// probes below then prove, per primitive, that its own mapped rect was painted
/// with its own resolved fill. A probe scans the primitive's mapped rect for the
/// exact BGRA value of the fill, because the fixture's inside labels cover their
/// own bar boxes and a fixed sample point would land on label ink; the scan is
/// deterministic and the comparison is still exact pixel equality. Thin strokes
/// are antialiased and are covered by the digest rather than by a region probe.
/// The documented text-tolerance fallback is a measured switch: it is implemented
/// only if exact text pixels prove unstable at representative scale, and the
/// sidecar's <c>mode</c> field records which contract the committed bytes answer to.
/// </para>
/// <para>
/// <b>Regeneration is explicit.</b> Setting <c>GANTTCREATOR_RASTER_GOLDEN_OUT</c> to
/// a directory rewrites the per-width artefact set; the comparison path never
/// writes, and CI never sees this suite because the scene carries text, so it
/// carries <c>Category=FontProbe</c> and runs only on the reference machine
/// (ADR-0070 D4). The reference-machine command is
/// <c>dotnet test tests/GanttCreator.Raster.Tests -c Release --filter "Category=FontProbe"</c>.
/// </para>
/// </remarks>
[Trait("Category", "FontProbe")]
public sealed class GoldenRasterRepresentativeTests
{
    private const string _goldenSubdirectory = "raster";
    private const string _sceneSubdirectory = "scene";
    private const string _sceneFile = "reference-scene.json";
    private const string _exactMode = "exact";

    private static readonly WidthArm _narrowArm = new("narrow-800px", "800px", 800, 435);

    /// <summary>
    /// Gets the five owner-ratified arms as serializable theory rows: key, width
    /// request, and exact pixel size. The preset arms derive their width from
    /// <see cref="SizePresets"/> at full precision (never a hand-typed literal); the
    /// two pixel arms are the ratified narrow and wide widths.
    /// </summary>
    public static TheoryData<string, string, int, int> Arms { get; } = new()
    {
        { "a4-landscape", CmWidth(SizePresets.A4Landscape.WidthPt), 3508, 1906 },
        { "presentation-16x9", CmWidth(SizePresets.Presentation16x9.WidthPt), 5333, 2898 },
        { "presentation-4x3", CmWidth(SizePresets.Presentation4x3.WidthPt), 4267, 2318 },
        { _narrowArm.Key, _narrowArm.WidthText, _narrowArm.PixelWidth, _narrowArm.PixelHeight },
        { "wide-2400px", "2400px", 2400, 1304 },
    };

    /// <summary>One owner-ratified representative render width.</summary>
    private sealed record WidthArm(string Key, string WidthText, int PixelWidth, int PixelHeight);

    [Theory]
    [MemberData(nameof(Arms))]
    public void The_representative_render_at_each_approved_width_matches_its_approved_baseline(
        string key, string widthText, int pixelWidth, int pixelHeight)
    {
        ArgumentNullException.ThrowIfNull(key);
        ArgumentNullException.ThrowIfNull(widthText);

        WidthArm arm = new(key, widthText, pixelWidth, pixelHeight);
        GanttScene scene = LoadReferenceScene();
        PixelDimensions dimensions = ExportSize.ToPixels(
            ExportSize.ParseWidth(arm.WidthText),
            scene.ChartBounds.Width,
            scene.ChartBounds.Height);

        Assert.Equal(arm.PixelWidth, dimensions.PixelWidth);
        Assert.Equal(arm.PixelHeight, dimensions.PixelHeight);

        using RasterRenderResult result = SceneRasterRenderer.Render(scene, dimensions, RasterBackground.Transparent);

        if (ResolveRegenerationDirectory() is { Length: > 0 } directory)
        {
            WriteBaseline(directory, arm, result);
            return;
        }

        Assert.Empty(CompareAgainstBaselines(GoldenDirectory, [arm]));
        AssertHardGeometryInteriors(scene, result);
    }

    [Fact]
    public void A_repeated_render_at_one_width_produces_an_identical_digest()
    {
        GanttScene scene = LoadReferenceScene();
        PixelDimensions dimensions = ExportSize.ToPixels(
            ExportSize.ParseWidth(_narrowArm.WidthText),
            scene.ChartBounds.Width,
            scene.ChartBounds.Height);

        using RasterRenderResult first = SceneRasterRenderer.Render(scene, dimensions, RasterBackground.Transparent);
        using RasterRenderResult second = SceneRasterRenderer.Render(scene, dimensions, RasterBackground.Transparent);

        Assert.Equal(ComputeDigest(first), ComputeDigest(second));
    }

    [Fact]
    public void A_tampered_digest_is_reported_as_a_mismatch_and_never_regenerates()
    {
        // The bad input is constructed here: a self-consistent baseline is written to a
        // temporary directory and its digest is then corrupted. The comparison must
        // report the digest field by name - a path that silently re-wrote the baseline
        // would pass a digest check against itself and defeat the whole guard.
        using TemporaryDirectory temp = new();
        using (RasterRenderResult result = RenderArm(_narrowArm))
        {
            WriteBaseline(temp.Path, _narrowArm, result);
        }

        string metaPath = MetaPath(temp.Path, _narrowArm);
        File.WriteAllText(
            metaPath,
            File.ReadAllText(metaPath).Replace("digest=sha256:", "digest=sha256:0", StringComparison.Ordinal));

        IReadOnlyList<string> mismatches = CompareAgainstBaselines(temp.Path, [_narrowArm]);

        Assert.Contains(mismatches, mismatch => mismatch.StartsWith("digest:", StringComparison.Ordinal));
    }

    [Fact]
    public void A_tampered_metadata_field_fails_with_the_field_named()
    {
        using TemporaryDirectory temp = new();
        using (RasterRenderResult result = RenderArm(_narrowArm))
        {
            WriteBaseline(temp.Path, _narrowArm, result);
        }

        string metaPath = MetaPath(temp.Path, _narrowArm);
        File.WriteAllText(
            metaPath,
            File.ReadAllText(metaPath).Replace("mode=" + _exactMode, "mode=text-tolerance", StringComparison.Ordinal));

        IReadOnlyList<string> mismatches = CompareAgainstBaselines(temp.Path, [_narrowArm]);

        Assert.Contains(mismatches, mismatch => mismatch.StartsWith("mode:", StringComparison.Ordinal));
    }

    [Fact]
    public void The_regeneration_route_writes_exactly_the_artefact_set()
    {
        using TemporaryDirectory temp = new();
        using RasterRenderResult result = RenderArm(_narrowArm);

        WriteBaseline(temp.Path, _narrowArm, result);

        string[] files = [.. Directory.GetFiles(temp.Path).Order(StringComparer.Ordinal)];
        Assert.Equal([MetaPath(temp.Path, _narrowArm), PngPath(temp.Path, _narrowArm)], files);
        Assert.Equal("sha256:" + ComputeDigest(result), ReadMeta(temp.Path, _narrowArm)["digest"]);
    }

    [Fact]
    public void The_comparison_path_never_writes()
    {
        using TemporaryDirectory temp = new();
        using (RasterRenderResult result = RenderArm(_narrowArm))
        {
            WriteBaseline(temp.Path, _narrowArm, result);
        }

        string[] before = Snapshot(temp.Path);
        Assert.Empty(CompareAgainstBaselines(temp.Path, [_narrowArm]));

        Assert.Equal(before, Snapshot(temp.Path));
    }

    /// <summary>
    /// Gets the directory the explicit regeneration route writes to, or
    /// <see langword="null"/> when the comparison path should run.
    /// </summary>
    /// <remarks>
    /// The only environment read in the suite, named so the two paths are visible
    /// where they are chosen. Never set in CI.
    /// </remarks>
    private static string? ResolveRegenerationDirectory() =>
        Environment.GetEnvironmentVariable("GANTTCREATOR_RASTER_GOLDEN_OUT");

    private static GanttScene LoadReferenceScene() =>
        SceneSnapshot.Deserialize(File.ReadAllText(ScenePath));

    private static string ScenePath =>
        Path.Combine(AppContext.BaseDirectory, "golden", _sceneSubdirectory, _sceneFile);

    private static string GoldenDirectory =>
        Path.Combine(AppContext.BaseDirectory, "golden", _goldenSubdirectory);

    private static string CmWidth(double widthPt) =>
        (widthPt / SizePresets.PointsPerInch * ExportSize.CmPerInch).ToString("R", CultureInfo.InvariantCulture) + "cm";

    private static string MetaPath(string directory, WidthArm arm) =>
        Path.Combine(directory, "representative-" + arm.Key + ".meta");

    private static string PngPath(string directory, WidthArm arm) =>
        Path.Combine(directory, "representative-" + arm.Key + ".png");

    private static RasterRenderResult RenderArm(WidthArm arm)
    {
        GanttScene scene = LoadReferenceScene();
        PixelDimensions dimensions = ExportSize.ToPixels(
            ExportSize.ParseWidth(arm.WidthText),
            scene.ChartBounds.Width,
            scene.ChartBounds.Height);
        return SceneRasterRenderer.Render(scene, dimensions, RasterBackground.Transparent);
    }

    private static string ComputeDigest(RasterRenderResult result) =>
        Convert.ToHexString(SHA256.HashData(result.Pixels));

    /// <summary>Compares the recorded metadata of the supplied arms against the current run.</summary>
    /// <returns>One message per mismatched or missing field; empty when every field matches.</returns>
    /// <remarks>Pure comparison: this path never writes a baseline.</remarks>
    private static List<string> CompareAgainstBaselines(string directory, IReadOnlyList<WidthArm> arms)
    {
        List<string> mismatches = [];
        foreach (WidthArm arm in arms)
        {
            Dictionary<string, string> recorded = ReadMeta(directory, arm);
            GanttScene scene = LoadReferenceScene();
            PixelDimensions dimensions = ExportSize.ToPixels(
                ExportSize.ParseWidth(arm.WidthText),
                scene.ChartBounds.Width,
                scene.ChartBounds.Height);
            using RasterRenderResult result = SceneRasterRenderer.Render(scene, dimensions, RasterBackground.Transparent);

            foreach ((string field, string expected) in CurrentFields(arm, result))
            {
                if (!recorded.TryGetValue(field, out string? value))
                {
                    mismatches.Add(field + ": <missing> (expected '" + expected + "')");
                }
                else if (!string.Equals(expected, value, StringComparison.Ordinal))
                {
                    mismatches.Add(field + ": expected '" + expected + "', recorded '" + value + "'");
                }
            }
        }

        return mismatches;
    }

    private static Dictionary<string, string> ReadMeta(string directory, WidthArm arm)
    {
        string path = MetaPath(directory, arm);
        Assert.True(
            File.Exists(path),
            "Missing representative baseline metadata: " + path
                + ". Regenerate through GANTTCREATOR_RASTER_GOLDEN_OUT after human review.");

        Dictionary<string, string> fields = new(StringComparer.Ordinal);
        foreach (string line in File.ReadAllText(path).Split('\n'))
        {
            string trimmed = line.TrimEnd('\r');
            if (trimmed.Length == 0)
            {
                continue;
            }

            int separator = trimmed.IndexOf('=', StringComparison.Ordinal);
            Assert.True(separator > 0, "Malformed metadata line in " + path + ": '" + trimmed + "'.");
            fields[trimmed[..separator]] = trimmed[(separator + 1)..];
        }

        Assert.True(fields.ContainsKey("digest"), "Metadata for " + arm.Key + " carries no digest field.");
        return fields;
    }

    /// <summary>The metadata the current run would record for one arm, in sidecar order.</summary>
    private static IEnumerable<(string Field, string Value)> CurrentFields(WidthArm arm, RasterRenderResult result) =>
    [
        ("width", arm.WidthText),
        ("pixels", arm.PixelWidth.ToString(CultureInfo.InvariantCulture) + "x" + arm.PixelHeight.ToString(CultureInfo.InvariantCulture)),
        ("dpi", "300"),
        ("surface", result.Determinism.SurfaceKind),
        ("antialias", result.Determinism.Antialias.ToString()),
        ("skiasharp", result.Determinism.SkiaVersion),
        ("font-identity", result.Determinism.FontIdentity),
        ("os", Environment.OSVersion.VersionString),
        ("culture", CultureInfo.CurrentCulture.Name),
        ("mode", _exactMode),
        ("digest", "sha256:" + ComputeDigest(result)),
    ];

    private static void WriteBaseline(string directory, WidthArm arm, RasterRenderResult result)
    {
        _ = System.IO.Directory.CreateDirectory(directory);
        File.WriteAllText(
            MetaPath(directory, arm),
            string.Join('\n', CurrentFields(arm, result).Select(field => field.Field + "=" + field.Value)),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        using SKImage image = SKImage.FromPixelCopy(
            new SKImageInfo(result.PixelWidth, result.PixelHeight, SKColorType.Bgra8888, SKAlphaType.Premul),
            result.Pixels);
        using SKData data = image.Encode(SKEncodedImageFormat.Png, 100);
        File.WriteAllBytes(PngPath(directory, arm), data.ToArray());
    }

    private static string[] Snapshot(string directory) =>
        [.. System.IO.Directory.GetFiles(directory)
            .Order(StringComparer.Ordinal)
            .Select(file => file + ":" + new FileInfo(file).Length.ToString(CultureInfo.InvariantCulture)
                + ":" + File.GetLastWriteTimeUtc(file).Ticks.ToString(CultureInfo.InvariantCulture))];

    /// <summary>
    /// Asserts the scene's resolved fill reaches three hard-geometry interiors at the
    /// mapped pixels: an alternate band, an activity bar, and a milestone marker.
    /// </summary>
    private static void AssertHardGeometryInteriors(GanttScene scene, RasterRenderResult result)
    {
        SceneRect band = RequireRect(
            scene,
            rect => rect.PrimitiveId.StartsWith("chart:band:", StringComparison.Ordinal),
            "alternate band");
        SceneRect bar = RequireRect(
            scene,
            rect => rect.PrimitiveId.EndsWith(":bar", StringComparison.Ordinal),
            "activity bar");
        ScenePolygon marker = scene.Primitives.OfType<ScenePolygon>()
            .FirstOrDefault(polygon => polygon.PrimitiveId.EndsWith(":marker", StringComparison.Ordinal))
            ?? throw new InvalidOperationException("The committed reference scene carries no milestone polygon.");

        AssertFillPaintedInside(scene, result, band.Bounds, band.Style.FillColour, "band");
        AssertFillPaintedInside(scene, result, bar.Bounds, bar.Style.FillColour, "bar");
        AssertFillPaintedInside(
            scene,
            result,
            BoundingRect(marker.Points),
            marker.Style.FillColour,
            "milestone");
    }

    private static SceneRect RequireRect(GanttScene scene, Func<SceneRect, bool> predicate, string label) =>
        scene.Primitives.OfType<SceneRect>().FirstOrDefault(predicate)
            ?? throw new InvalidOperationException("The committed reference scene carries no " + label + " rectangle.");

    private static RectD BoundingRect(IReadOnlyList<PointD> points)
    {
        Assert.NotEmpty(points);
        double left = points.Min(point => point.X);
        double top = points.Min(point => point.Y);
        double right = points.Max(point => point.X);
        double bottom = points.Max(point => point.Y);
        return new RectD(left, top, right - left, bottom - top);
    }

    /// <summary>
    /// Asserts a primitive's own mapped rect was painted with its own resolved fill,
    /// by scanning the rect's interior for one pixel of the exact BGRA value.
    /// </summary>
    /// <remarks>
    /// The digest is the exact whole-buffer gate; this probe exists to attribute a
    /// digest change to a named primitive. A fixed sample point cannot be used for a
    /// bar: the fixture's inside labels cover their bar boxes, so a sample would be
    /// label ink rather than the fill. The scan is deterministic (row-major from the
    /// rect's top-left interior) and compares exact bytes, so it never accepts a
    /// near colour.
    /// </remarks>
    private static void AssertFillPaintedInside(
        GanttScene scene,
        RasterRenderResult result,
        RectD rect,
        ColourHex? fill,
        string label)
    {
        ColourHex resolved = fill
            ?? throw new InvalidOperationException("The probe " + label + " resolved no fill colour, so no pixel can match it.");

        uint argb = resolved.ARGB;
        byte blue = (byte)argb;
        byte green = (byte)(argb >> 8);
        byte red = (byte)(argb >> 16);
        byte alpha = (byte)(argb >> 24);

        // One pixel in from each mapped edge: the edge row and column are antialiased.
        int left = Math.Max(0, MapX(rect.Left, scene, result) + 1);
        int top = Math.Max(0, MapY(rect.Top, scene, result) + 1);
        int right = Math.Min(result.PixelWidth - 1, MapX(rect.Right, scene, result) - 1);
        int bottom = Math.Min(result.PixelHeight - 1, MapY(rect.Bottom, scene, result) - 1);

        for (int y = top; y <= bottom; y++)
        {
            for (int x = left; x <= right; x++)
            {
                int offset = ((y * result.PixelWidth) + x) * 4;
                ReadOnlySpan<byte> pixel = result.Pixels.Slice(offset, 4);
                if (pixel[0] == blue && pixel[1] == green && pixel[2] == red && pixel[3] == alpha)
                {
                    return;
                }
            }
        }

        throw new InvalidOperationException(
            "The " + label + " mapped to [" + left.ToString(CultureInfo.InvariantCulture) + ","
            + top.ToString(CultureInfo.InvariantCulture) + " - " + right.ToString(CultureInfo.InvariantCulture) + ","
            + bottom.ToString(CultureInfo.InvariantCulture) + "] in a "
            + result.PixelWidth.ToString(CultureInfo.InvariantCulture) + "x"
            + result.PixelHeight.ToString(CultureInfo.InvariantCulture)
            + " render contains no pixel of its resolved fill " + resolved.ToString() + ".");
    }

    // The renderer's own mapping (RasterTransform is internal to the Raster project):
    // one rounded multiplication per axis, matching RasterTransform.Create/X/Y.
    private static int MapX(double pointX, GanttScene scene, RasterRenderResult result) =>
        (int)Math.Round((pointX - scene.ChartBounds.X) * result.PixelWidth / scene.ChartBounds.Width);

    private static int MapY(double pointY, GanttScene scene, RasterRenderResult result) =>
        (int)Math.Round((pointY - scene.ChartBounds.Y) * result.PixelHeight / scene.ChartBounds.Height);

    /// <summary>A unique temporary directory removed in <see cref="Dispose"/>.</summary>
    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ganttcreator-r8-4-" + Guid.NewGuid().ToString("N"));
            _ = System.IO.Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (System.IO.Directory.Exists(Path))
            {
                System.IO.Directory.Delete(Path, recursive: true);
            }
        }
    }
}
