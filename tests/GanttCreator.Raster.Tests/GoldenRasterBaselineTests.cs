using System.Text;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using SkiaSharp;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// The R8.2 category-1 golden baseline: a small primitives/z-order scene,
/// committed as canonical JSON, compared byte-for-byte against its pinned BGRA
/// pixels, with a human-review PNG alongside.
/// </summary>
/// <remarks>
/// <para>
/// Regeneration is explicit (<c>GANTTCREATOR_RASTER_GOLDEN_OUT</c>), never
/// automatic in CI, and a baseline update is a dedicated commit stating why
/// (test strategy "Golden images"). The committed baseline is the raw BGRA
/// pixel buffer; the PNG is the human-review artefact only.
/// </para>
/// <para>
/// The golden metadata's font fields are not applicable: the category-1 scene
/// carries one text primitive only to prove the R8.2 text skip, and no text is
/// drawn. Full font pinning arrives with R8.2a/R8.3.
/// </para>
/// </remarks>
public class GoldenRasterBaselineTests
{
    private const int _canvasPixelWidth = 200;
    private const int _canvasPixelHeight = 120;

    private static string GoldenDirectory => Path.Combine(AppContext.BaseDirectory, "golden", "raster");

    private static string FixturePath => Path.Combine(GoldenDirectory, "category-1-primitives.json");

    private static string BaselinePath => Path.Combine(GoldenDirectory, "category-1-primitives.bgra");

    [Fact]
    public void The_committed_fixture_matches_the_scene_builder()
    {
        Assert.True(File.Exists(FixturePath), "Missing category-1 fixture: " + FixturePath);

        string committed = File.ReadAllText(FixturePath).TrimEnd('\r', '\n');
        Assert.Equal(committed, SceneSnapshot.Serialize(BuildCategoryOneScene()));
    }

    [Fact]
    public void Category_one_primitives_match_the_approved_golden_baseline()
    {
        string? directory = Environment.GetEnvironmentVariable("GANTTCREATOR_RASTER_GOLDEN_OUT");
        if (!string.IsNullOrEmpty(directory))
        {
            WriteBaselines(directory, BuildCategoryOneScene());
            return;
        }

        GanttScene scene = SceneSnapshot.Deserialize(File.ReadAllText(FixturePath));
        using RasterRenderResult result = SceneRasterRenderer.Render(
            scene,
            new PixelDimensions(_canvasPixelWidth, _canvasPixelHeight),
            RasterBackground.Transparent);

        byte[] expected = File.ReadAllBytes(BaselinePath);
        Assert.Equal(_canvasPixelWidth, result.PixelWidth);
        Assert.Equal(_canvasPixelHeight, result.PixelHeight);
        Assert.Equal(expected.Length, result.Pixels.Length);
        Assert.Equal(expected, result.Pixels.ToArray());
    }

    private static void WriteBaselines(string directory, GanttScene scene)
    {
        _ = Directory.CreateDirectory(directory);
        File.WriteAllText(
            Path.Combine(directory, "category-1-primitives.json"),
            SceneSnapshot.Serialize(scene),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        using RasterRenderResult result = SceneRasterRenderer.Render(
            scene,
            new PixelDimensions(_canvasPixelWidth, _canvasPixelHeight),
            RasterBackground.Transparent);
        File.WriteAllBytes(Path.Combine(directory, "category-1-primitives.bgra"), result.Pixels.ToArray());

        // The review PNG is a test-only artefact for human approval; production
        // PNG encoding (and pHYs metadata) is R8.5's row.
        using SKImage image = SKImage.FromPixelCopy(
            new SKImageInfo(_canvasPixelWidth, _canvasPixelHeight, SKColorType.Bgra8888, SKAlphaType.Premul),
            result.Pixels);
        using SKData data = image.Encode(SKEncodedImageFormat.Png, 100);
        File.WriteAllBytes(Path.Combine(directory, "category-1-primitives.png"), data.ToArray());
    }

    private static GanttScene BuildCategoryOneScene()
    {
        ColourHex background = ColourHex.Parse("#F2F2F2");
        ColourHex grid = ColourHex.Parse("#999999");
        ColourHex blue = ColourHex.Parse("#3366CC");
        ColourHex red = ColourHex.Parse("#CC3333");
        ColourHex black = ColourHex.Parse("#000000");
        ColourHex green = ColourHex.Parse("#22AA66");
        ColourHex yellow = ColourHex.Parse("#FFCC00");

        ScenePrimitive[] primitives =
        [
            new SceneRect(
                "chart:background",
                SceneOwnerId.Chart,
                ZLayer.Background,
                new RectD(0, 0, 100, 60),
                new SceneStyle("Fixture", fillColour: background)),
            new SceneLine(
                "chart:grid",
                SceneOwnerId.Chart,
                ZLayer.Grid,
                new PointD(0, 30),
                new PointD(100, 30),
                new SceneStyle("Fixture", strokeColour: grid, outlineWidthPt: 0.5)),
            new SceneRect(
                "activity:body",
                SceneOwnerId.Chart,
                ZLayer.ActivityBody,
                new RectD(10, 10, 40, 12),
                new SceneStyle("Fixture", fillColour: blue)),
            new SceneRect(
                "activity:second",
                SceneOwnerId.Chart,
                ZLayer.ActivityBody,
                new RectD(55, 10, 35, 12),
                new SceneStyle("Fixture", fillColour: red)),
            new ScenePolygon(
                "milestone:marker",
                SceneOwnerId.Chart,
                ZLayer.Milestone,
                [new PointD(70, 25), new PointD(75, 30), new PointD(70, 35), new PointD(65, 30)],
                new SceneStyle("Fixture", fillColour: black)),
            new SceneRect(
                "overlay:band",
                SceneOwnerId.Chart,
                ZLayer.ActivityBody,
                new RectD(20, 45, 30, 10),
                new SceneStyle("Fixture", fillColour: green)),
            new SceneRect(
                "label:bar",
                SceneOwnerId.Chart,
                ZLayer.Label,
                new RectD(30, 47, 30, 10),
                new SceneStyle("Fixture", fillColour: yellow)),
            new SceneText(
                "chart:title:text",
                SceneOwnerId.Chart,
                ZLayer.Title,
                "Primitives",
                new RectD(5, 2, 90, 8),
                new SceneStyle("FixtureText", fontFamily: "Aptos", fontSizePt: 10),
                GanttTextAlignment.Left),
        ];

        SceneCreationOutcome outcome = GanttScene.TryCreate(
            new RectD(0, 0, 100, 60),
            new RectD(0, 0, 100, 60),
            primitives,
            []);
        Assert.True(outcome.Succeeded, $"Category-1 fixture scene was refused: {outcome.Refusal}.");
        return outcome.Scene!;
    }
}
