using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// Host-free pins for R8.3's fail-closed text contract: a text-bearing scene
/// resolves its pinned faces before any pixel is drawn, so a missing or
/// unpinned face aborts the render, and a malformed typography token is a
/// scene-contract fault. These run without Office cloud fonts.
/// </summary>
public class SceneRasterTextRefusalTests
{
    private static readonly ColourHex _black = ColourHex.Parse("#000000");
    private static readonly PixelDimensions _square = new(100, 100);

    [Fact]
    public void Render_refuses_a_text_scene_before_drawing_when_the_pinned_face_is_missing()
    {
        GanttScene scene = TextScene(Text("label", "Hello", family: "Aptos", sizePt: 10));
        string missingRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        using PinnedFontSource source = new(missingRoot);

        PinnedFontException exception = Assert.Throws<PinnedFontException>(
            () => SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent, source));

        Assert.Equal(PinnedFontRefusal.FontFileMissing, exception.Refusal);
        Assert.Equal("Aptos", exception.FontFamily);
    }

    [Fact]
    public void Render_refuses_a_font_family_that_is_not_pinned()
    {
        GanttScene scene = TextScene(Text("label", "Hello", family: "Segoe UI", sizePt: 10));
        using PinnedFontSource source = new(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")));

        PinnedFontException exception = Assert.Throws<PinnedFontException>(
            () => SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent, source));

        Assert.Equal(PinnedFontRefusal.FamilyMismatch, exception.Refusal);
        Assert.Equal("Segoe UI", exception.FontFamily);
    }

    [Fact]
    public void Render_refuses_a_text_primitive_with_no_font_family()
    {
        GanttScene scene = TextScene(Text("label", "Hello", family: null, sizePt: 10));
        using PinnedFontSource source = new(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")));

        PinnedFontException exception = Assert.Throws<PinnedFontException>(
            () => SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent, source));

        Assert.Equal(PinnedFontRefusal.FamilyMismatch, exception.Refusal);
        Assert.Equal(string.Empty, exception.FontFamily);
    }

    [Fact]
    public void Render_refuses_a_text_primitive_with_no_font_size()
    {
        GanttScene scene = TextScene(Text("label", "Hello", family: "Aptos", sizePt: null));
        using PinnedFontSource source = new(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")));

        ArgumentException exception = Assert.Throws<ArgumentException>(
            () => SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent, source));

        Assert.Contains("label", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void The_determinism_descriptor_names_no_font_for_a_text_free_scene()
    {
        GanttScene scene = Scene(new SceneRect(
            "body",
            SceneOwnerId.Chart,
            ZLayer.ActivityBody,
            new RectD(10, 10, 20, 10),
            new SceneStyle("Fixture", fillColour: _black)));

        using RasterRenderResult result = SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent);

        Assert.Equal("none", result.Determinism.FontIdentity);
        Assert.Equal(RasterDeterminism.Pinned.SurfaceKind, result.Determinism.SurfaceKind);
    }

    [Fact]
    public void RasterDeterminism_values_with_the_font_identity_compare_by_value()
    {
        RasterDeterminism plain = RasterDeterminism.Pinned;
        RasterDeterminism withFont = plain with { FontIdentity = "Aptos Version 2.01;O365 upem=2048" };

        Assert.NotEqual(plain, withFont);
        Assert.Equal(withFont, plain with { FontIdentity = "Aptos Version 2.01;O365 upem=2048" });
    }

    private static GanttScene TextScene(params SceneText[] texts) => Scene([.. texts]);

    private static GanttScene Scene(params ScenePrimitive[] primitives)
    {
        SceneCreationOutcome outcome = GanttScene.TryCreate(
            new RectD(0, 0, 100, 100),
            new RectD(0, 0, 100, 100),
            primitives,
            []);
        Assert.True(outcome.Succeeded, $"Scene creation was refused: {outcome.Refusal}.");
        return outcome.Scene!;
    }

    private static SceneText Text(string id, string value, string? family, double? sizePt) =>
        new(
            id,
            SceneOwnerId.Chart,
            ZLayer.Label,
            value,
            new RectD(10, 10, 60, 10),
            new SceneStyle("FixtureText", fontFamily: family, fontSizePt: sizePt, textColour: _black),
            GanttTextAlignment.Left);
}
