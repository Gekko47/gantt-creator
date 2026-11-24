using System.Text;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using SkiaSharp;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// The host-free raster golden baselines: small scenes committed as canonical
/// JSON, compared byte-for-byte against pinned BGRA pixels, with a human-review
/// PNG alongside. Category 1 is the R8.2 z-order primitive scene; category 3 is
/// R8.3's hatch and dash scene.
/// </summary>
/// <remarks>
/// <para>
/// Regeneration is explicit (<c>GANTTCREATOR_RASTER_GOLDEN_OUT</c>), never
/// automatic in CI, and a baseline update is a dedicated commit stating why
/// (test strategy "Golden images"). The committed baseline is the raw BGRA
/// pixel buffer; the PNG is the human-review artefact only.
/// </para>
/// <para>
/// Neither scene carries text, so both compare without a pinned font and stay
/// in the routine gate. The text baseline is
/// <see cref="GoldenRasterTextBaselineTests"/>, which carries
/// <c>Category=FontProbe</c> because it needs the Office cloud fonts.
/// </para>
/// </remarks>
public class GoldenRasterBaselineTests
{
    private const string _categoryOneName = "category-1-primitives";
    private const string _categoryThreeName = "category-3-patterns";

    private static readonly PixelDimensions _categoryOneDimensions = new(200, 120);
    private static readonly PixelDimensions _categoryThreeDimensions = new(200, 140);

    private static string GoldenDirectory => Path.Combine(AppContext.BaseDirectory, "golden", "raster");

    [Fact]
    public void The_committed_fixture_matches_the_scene_builder()
    {
        AssertFixtureMatches(_categoryOneName, BuildCategoryOneScene());
    }

    [Fact]
    public void The_committed_pattern_fixture_matches_the_scene_builder()
    {
        AssertFixtureMatches(_categoryThreeName, BuildCategoryThreeScene());
    }

    [Fact]
    public void Category_one_primitives_match_the_approved_golden_baseline()
    {
        AssertBaselineMatches(_categoryOneName, BuildCategoryOneScene(), _categoryOneDimensions);
    }

    [Fact]
    public void Category_three_patterns_match_the_approved_golden_baseline()
    {
        AssertBaselineMatches(_categoryThreeName, BuildCategoryThreeScene(), _categoryThreeDimensions);
    }

    private static void AssertFixtureMatches(string name, GanttScene scene)
    {
        string path = GoldenPath(name, "json");
        Assert.True(File.Exists(path), $"Missing {name} fixture: {path}");

        string committed = File.ReadAllText(path).TrimEnd('\r', '\n');
        Assert.Equal(committed, SceneSnapshot.Serialize(scene));
    }

    private static void AssertBaselineMatches(string name, GanttScene scene, PixelDimensions dimensions)
    {
        string? directory = Environment.GetEnvironmentVariable("GANTTCREATOR_RASTER_GOLDEN_OUT");
        if (!string.IsNullOrEmpty(directory))
        {
            WriteBaselines(directory, name, scene, dimensions);
            return;
        }

        GanttScene committed = SceneSnapshot.Deserialize(File.ReadAllText(GoldenPath(name, "json")));
        using RasterRenderResult result = SceneRasterRenderer.Render(
            committed,
            dimensions,
            RasterBackground.Transparent);

        byte[] expected = File.ReadAllBytes(GoldenPath(name, "bgra"));
        Assert.Equal(dimensions.PixelWidth, result.PixelWidth);
        Assert.Equal(dimensions.PixelHeight, result.PixelHeight);
        Assert.Equal(expected.Length, result.Pixels.Length);
        Assert.Equal(expected, result.Pixels.ToArray());
    }

    private static void WriteBaselines(string directory, string name, GanttScene scene, PixelDimensions dimensions)
    {
        _ = Directory.CreateDirectory(directory);
        File.WriteAllText(
            Path.Combine(directory, name + ".json"),
            SceneSnapshot.Serialize(scene),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        using RasterRenderResult result = SceneRasterRenderer.Render(
            scene,
            dimensions,
            RasterBackground.Transparent);
        File.WriteAllBytes(Path.Combine(directory, name + ".bgra"), result.Pixels.ToArray());

        // The review PNG is a test-only artefact for human approval; production
        // PNG encoding (and pHYs metadata) is R8.5's row.
        using SKImage image = SKImage.FromPixelCopy(
            new SKImageInfo(dimensions.PixelWidth, dimensions.PixelHeight, SKColorType.Bgra8888, SKAlphaType.Premul),
            result.Pixels);
        using SKData data = image.Encode(SKEncodedImageFormat.Png, 100);
        File.WriteAllBytes(Path.Combine(directory, name + ".png"), data.ToArray());
    }

    private static string GoldenPath(string name, string extension) =>
        Path.Combine(GoldenDirectory, name + "." + extension);

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
        ];

        SceneCreationOutcome outcome = GanttScene.TryCreate(
            new RectD(0, 0, 100, 60),
            new RectD(0, 0, 100, 60),
            primitives,
            []);
        Assert.True(outcome.Succeeded, $"Category-1 fixture scene was refused: {outcome.Refusal}.");
        return outcome.Scene!;
    }

    private static GanttScene BuildCategoryThreeScene()
    {
        ColourHex background = ColourHex.Parse("#F2F2F2");
        ColourHex white = ColourHex.Parse("#FFFFFF");
        ColourHex black = ColourHex.Parse("#000000");
        ColourHex blue = ColourHex.Parse("#3366CC");

        ScenePrimitive[] primitives =
        [
            new SceneRect(
                "pattern:background",
                SceneOwnerId.Chart,
                ZLayer.Background,
                new RectD(0, 0, 100, 70),
                new SceneStyle("Fixture", fillColour: background)),
            new SceneRect(
                "pattern:forward",
                SceneOwnerId.Chart,
                ZLayer.ActivityBody,
                new RectD(5, 5, 42, 22),
                new SceneStyle(
                    "Fixture",
                    fillColour: white,
                    strokeColour: black,
                    outlineWidthPt: 1,
                    hatchPattern: GanttHatchPattern.ForwardDiagonal)),
            new SceneRect(
                "pattern:cross",
                SceneOwnerId.Chart,
                ZLayer.ActivityBody,
                new RectD(53, 5, 42, 22),
                new SceneStyle(
                    "Fixture",
                    fillColour: white,
                    strokeColour: blue,
                    outlineWidthPt: 1,
                    hatchPattern: GanttHatchPattern.Cross)),
            new ScenePolygon(
                "pattern:polygon",
                SceneOwnerId.Chart,
                ZLayer.ActivityBody,
                [new PointD(5, 33), new PointD(40, 33), new PointD(33, 48), new PointD(12, 48)],
                new SceneStyle(
                    "Fixture",
                    fillColour: white,
                    strokeColour: black,
                    outlineWidthPt: 1,
                    hatchPattern: GanttHatchPattern.BackwardDiagonal)),
            DashLine("pattern:dash", 5, 53, 47, GanttStrokeDash.Dash, black),
            DashLine("pattern:dash-dot", 53, 53, 95, GanttStrokeDash.DashDot, black),
            DashLine("pattern:round-dot", 5, 60, 47, GanttStrokeDash.RoundDot, black),
            DashLine("pattern:long-dash", 53, 60, 95, GanttStrokeDash.LongDash, black),
            DashLine("pattern:square-dot", 5, 67, 47, GanttStrokeDash.SquareDot, black),
            DashLine("pattern:solid", 53, 67, 95, GanttStrokeDash.Solid, black),
        ];

        SceneCreationOutcome outcome = GanttScene.TryCreate(
            new RectD(0, 0, 100, 70),
            new RectD(0, 0, 100, 70),
            primitives,
            []);
        Assert.True(outcome.Succeeded, $"Category-3 fixture scene was refused: {outcome.Refusal}.");
        return outcome.Scene!;
    }

    private static SceneLine DashLine(
        string primitiveId,
        double fromX,
        double y,
        double toX,
        GanttStrokeDash dash,
        ColourHex colour) => new(
            primitiveId,
            SceneOwnerId.Chart,
            ZLayer.Grid,
            new PointD(fromX, y),
            new PointD(toX, y),
            new SceneStyle("Fixture", strokeColour: colour, outlineWidthPt: 1, strokeDash: dash));
}
