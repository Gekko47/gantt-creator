using System.Text;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using SkiaSharp;
using Xunit;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// The R8.3 category-2 golden baseline: a small text scene drawn with the four
/// pinned Aptos faces, committed as canonical JSON plus its exact BGRA pixels
/// and the font identity that produced them, with a human-review PNG alongside.
/// </summary>
/// <remarks>
/// <para>
/// This is the only raster baseline that needs a font, so it carries
/// <c>Category=FontProbe</c>: the routine gates exclude that category because a
/// GitHub-hosted runner has no Office cloud fonts (ADR-0070 D4). The
/// reference-machine command is
/// <c>dotnet test tests/GanttCreator.Raster.Tests -c Release --filter "Category=FontProbe"</c>.
/// </para>
/// <para>
/// The baseline pins the identity, not just the bytes: the comparison asserts
/// the rendered <see cref="RasterDeterminism.FontIdentity"/> equals the committed
/// <c>.identity</c> file and names each pinned face, so a font swap cannot pass
/// as "close enough" (ADR-0070 D5: a font change is a measured change).
/// </para>
/// </remarks>
[Trait("Category", "FontProbe")]
public class GoldenRasterTextBaselineTests
{
    private const string _name = "category-2-text";

    private static readonly PixelDimensions _dimensions = new(300, 140);
    private static readonly RectD _bounds = new(0, 0, 150, 70);

    private static string GoldenDirectory => Path.Combine(AppContext.BaseDirectory, "golden", "raster");

    [Fact]
    public void The_committed_text_fixture_matches_the_scene_builder()
    {
        string path = GoldenPath("json");
        Assert.True(File.Exists(path), "Missing category-2 fixture: " + path);

        string committed = File.ReadAllText(path).TrimEnd('\r', '\n');
        Assert.Equal(committed, SceneSnapshot.Serialize(BuildCategoryTwoScene()));
    }

    [Fact]
    public void Category_two_text_matches_the_approved_golden_baseline()
    {
        string? directory = Environment.GetEnvironmentVariable("GANTTCREATOR_RASTER_GOLDEN_OUT");
        if (!string.IsNullOrEmpty(directory))
        {
            WriteBaselines(directory);
            return;
        }

        GanttScene scene = SceneSnapshot.Deserialize(File.ReadAllText(GoldenPath("json")));
        using RasterRenderResult result = SceneRasterRenderer.Render(
            scene,
            _dimensions,
            RasterBackground.Transparent);

        string expectedIdentity = File.ReadAllText(GoldenPath("identity")).TrimEnd('\r', '\n');
        Assert.Equal(expectedIdentity, result.Determinism.FontIdentity);
        foreach (PinnedFontFace face in PinnedFontSet.Faces)
        {
            Assert.Contains(face.PostScriptName, result.Determinism.FontIdentity, StringComparison.Ordinal);
        }

        byte[] expected = File.ReadAllBytes(GoldenPath("bgra"));
        Assert.Equal(_dimensions.PixelWidth, result.PixelWidth);
        Assert.Equal(_dimensions.PixelHeight, result.PixelHeight);
        Assert.Equal(expected.Length, result.Pixels.Length);
        Assert.Equal(expected, result.Pixels.ToArray());
    }

    private static void WriteBaselines(string directory)
    {
        GanttScene scene = BuildCategoryTwoScene();
        _ = Directory.CreateDirectory(directory);
        File.WriteAllText(
            Path.Combine(directory, _name + ".json"),
            SceneSnapshot.Serialize(scene),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        using RasterRenderResult result = SceneRasterRenderer.Render(
            scene,
            _dimensions,
            RasterBackground.Transparent);
        File.WriteAllBytes(Path.Combine(directory, _name + ".bgra"), result.Pixels.ToArray());
        File.WriteAllText(
            Path.Combine(directory, _name + ".identity"),
            result.Determinism.FontIdentity,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        // The review PNG is a test-only artefact for human approval; production
        // PNG encoding (and pHYs metadata) is R8.5's row.
        using SKImage image = SKImage.FromPixelCopy(
            new SKImageInfo(_dimensions.PixelWidth, _dimensions.PixelHeight, SKColorType.Bgra8888, SKAlphaType.Premul),
            result.Pixels);
        using SKData data = image.Encode(SKEncodedImageFormat.Png, 100);
        File.WriteAllBytes(Path.Combine(directory, _name + ".png"), data.ToArray());
    }

    private static string GoldenPath(string extension) =>
        Path.Combine(GoldenDirectory, _name + "." + extension);

    /// <summary>
    /// Builds the category-2 scene: eight label cells exercising all three
    /// alignments, the left inset, bold, the narrow family, the narrow bold
    /// face, and the ellipsis glyph the label truncation writes.
    /// </summary>
    private static GanttScene BuildCategoryTwoScene()
    {
        ColourHex white = ColourHex.Parse("#FFFFFF");
        ColourHex black = ColourHex.Parse("#000000");
        ColourHex blue = ColourHex.Parse("#3366CC");
        ColourHex green = ColourHex.Parse("#22AA66");
        ColourHex red = ColourHex.Parse("#CC3333");
        ColourHex light = ColourHex.Parse("#F2F2F2");

        const string aptos = "Aptos";
        const string narrow = "Aptos Narrow";

        ScenePrimitive[] primitives =
        [
            new SceneRect(
                "text:background",
                SceneOwnerId.Chart,
                ZLayer.Background,
                _bounds,
                new SceneStyle("Fixture", fillColour: white)),
            Cell("text:left", 8, 4, blue),
            Label("text:left:label", "Planned span", 8, 4, aptos, bold: false, GanttTextAlignment.Left, white),
            Cell("text:centre", 8, 20, green),
            Label("text:centre:label", "Centre label", 8, 20, aptos, bold: false, GanttTextAlignment.Centre, white),
            Cell("text:right", 8, 36, red),
            Label("text:right:label", "Right aligned", 8, 36, aptos, bold: false, GanttTextAlignment.Right, white),
            Cell("text:bold", 8, 52, light),
            Label("text:bold:label", "Bold label", 8, 52, aptos, bold: true, GanttTextAlignment.Left, black),
            Cell("text:narrow", 78, 4, light),
            Label("text:narrow:label", "Narrow label", 78, 4, narrow, bold: false, GanttTextAlignment.Left, black),
            Cell("text:inset", 78, 20, light),
            Label("text:inset:label", "Indented", 78, 20, aptos, bold: false, GanttTextAlignment.Left, black, leftInsetPt: 6),
            Cell("text:ellipsis", 78, 36, light),
            Label("text:ellipsis:label", "Long text…", 78, 36, aptos, bold: false, GanttTextAlignment.Left, black),
            Cell("text:narrow-bold", 78, 52, light),
            Label("text:narrow-bold:label", "Narrow bold", 78, 52, narrow, bold: true, GanttTextAlignment.Left, black),
        ];

        SceneCreationOutcome outcome = GanttScene.TryCreate(_bounds, _bounds, primitives, []);
        Assert.True(outcome.Succeeded, $"Category-2 fixture scene was refused: {outcome.Refusal}.");
        return outcome.Scene!;
    }

    private static SceneRect Cell(string primitiveId, double x, double y, ColourHex fill) => new(
        primitiveId,
        SceneOwnerId.Chart,
        ZLayer.ActivityBody,
        new RectD(x, y, 60, 12),
        new SceneStyle("Fixture", fillColour: fill));

    private static SceneText Label(
        string primitiveId,
        string text,
        double x,
        double y,
        string family,
        bool bold,
        GanttTextAlignment alignment,
        ColourHex colour,
        double leftInsetPt = 0) => new(
            primitiveId,
            SceneOwnerId.Chart,
            ZLayer.Label,
            text,
            new RectD(x, y, 60, 12),
            new SceneStyle(
                "Fixture",
                fontFamily: family,
                fontSizePt: 8,
                bold: bold,
                alignment: alignment,
                textColour: colour),
            alignment,
            leftInsetPt: leftInsetPt);
}
