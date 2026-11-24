using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// R8.3's pinned-font text pins: text draws with the resolved Aptos faces at
/// the scene's resolved bounds, side, alignment, and inset; nothing is
/// re-measured or reselected; the render is byte-deterministic; and clipping
/// is the canvas plus the primitive's own box.
/// </summary>
/// <remarks>
/// These tests need the Office cloud fonts, so the class carries
/// <c>Category=FontProbe</c> and is excluded from the routine gates (ADR-0070
/// D4). The reference-machine run is the font gate. The host-free refusal arms
/// live in <see cref="SceneRasterTextRefusalTests"/>.
/// </remarks>
[Trait("Category", "FontProbe")]
public class SceneRasterTextTests
{
    private static readonly ColourHex _black = ColourHex.Parse("#000000");
    private static readonly PixelDimensions _square = new(100, 100);

    [Fact]
    public void Text_is_drawn_inside_its_resolved_bounds_and_nothing_outside()
    {
        // The bounds are deliberately not derivable from the text: a renderer
        // that re-measured would not land inside these mapped pixels.
        GanttScene scene = TextScene(Text("label", "WWWW", new RectD(13.7, 20.3, 30.3, 9.9), GanttTextAlignment.Left));

        using RasterRenderResult result = Render(scene);

        InkBounds ink = AssertInk(result);
        Assert.InRange(ink.MinX, 14, 15);
        Assert.InRange(ink.MinY, 20, 21);
        Assert.InRange(ink.MaxX, 14, 44);
        Assert.InRange(ink.MaxY, 20, 30);
        Assert.Equal(0, Alpha(result, 10, 25));
        Assert.Equal(0, Alpha(result, 50, 25));
    }

    [Theory]
    [InlineData(GanttTextAlignment.Left)]
    [InlineData(GanttTextAlignment.Centre)]
    [InlineData(GanttTextAlignment.Right)]
    public void Text_alignment_positions_the_ink_inside_the_bounds(GanttTextAlignment alignment)
    {
        GanttScene scene = TextScene(Text("label", "Hi", new RectD(10, 40, 80, 10), alignment));

        using RasterRenderResult result = Render(scene);

        InkBounds ink = AssertInk(result);
        switch (alignment)
        {
            case GanttTextAlignment.Left:
                Assert.InRange(ink.MinX, 10, 13);
                Assert.True(ink.MaxX < 50, $"Left-aligned ink reached x={ink.MaxX}.");
                break;
            case GanttTextAlignment.Centre:
                Assert.InRange((ink.MinX + ink.MaxX) / 2, 47, 53);
                break;
            default:
                Assert.InRange(ink.MaxX, 86, 90);
                Assert.True(ink.MinX > 50, $"Right-aligned ink started at x={ink.MinX}.");
                break;
        }
    }

    [Fact]
    public void A_left_inset_shifts_the_ink_right()
    {
        GanttScene withoutInset = TextScene(Text("label", "Hi", new RectD(10, 40, 60, 10), GanttTextAlignment.Left));
        GanttScene withInset = TextScene(
            Text("label", "Hi", new RectD(10, 40, 60, 10), GanttTextAlignment.Left, leftInsetPt: 5));

        using RasterRenderResult plain = Render(withoutInset);
        using RasterRenderResult inset = Render(withInset);

        int plainMinX = AssertInk(plain).MinX;
        int insetMinX = AssertInk(inset).MinX;
        Assert.InRange(insetMinX - plainMinX, 3, 7);
    }

    [Fact]
    public void Bold_and_narrow_faces_draw_distinct_ink()
    {
        const string sample = "MMMMMMMM";
        RectD bounds = new(5, 40, 90, 10);

        using RasterRenderResult regular = Render(TextScene(Text("label", sample, bounds, GanttTextAlignment.Left)));
        using RasterRenderResult bold = Render(TextScene(Text("label", sample, bounds, GanttTextAlignment.Left, bold: true)));
        using RasterRenderResult narrow = Render(TextScene(
            Text("label", sample, bounds, GanttTextAlignment.Left, family: "Aptos Narrow")));

        InkBounds regularInk = AssertInk(regular);
        InkBounds boldInk = AssertInk(bold);
        InkBounds narrowInk = AssertInk(narrow);

        Assert.NotEqual(regular.Pixels.ToArray(), bold.Pixels.ToArray());
        Assert.True(boldInk.MaxX > regularInk.MaxX, "The bold face did not widen the same string.");
        Assert.True(narrowInk.MaxX < regularInk.MaxX, "The narrow face did not narrow the same string.");
    }

    [Fact]
    public void A_text_primitive_outside_the_canvas_paints_nothing()
    {
        GanttScene scene = TextScene(Text("label", "Hello", new RectD(120, 120, 20, 10), GanttTextAlignment.Left));

        using RasterRenderResult result = Render(scene);

        Assert.Equal(0, InkCount(result));
    }

    [Fact]
    public void A_partially_clipped_text_still_paints_inside_the_canvas()
    {
        GanttScene scene = TextScene(Text("label", "Hello", new RectD(-15, 40, 50, 10), GanttTextAlignment.Left));

        using RasterRenderResult result = Render(scene);

        Assert.True(InkCount(result) > 0, "A partially visible text primitive painted nothing.");
    }

    [Fact]
    public void A_text_primitive_with_no_text_colour_paints_nothing()
    {
        SceneText text = new(
            "label",
            SceneOwnerId.Chart,
            ZLayer.Label,
            "Hello",
            new RectD(10, 10, 60, 10),
            new SceneStyle("FixtureText", fontFamily: "Aptos", fontSizePt: 10),
            GanttTextAlignment.Left);

        using RasterRenderResult result = Render(TextScene(text));

        Assert.Equal(0, InkCount(result));
    }

    [Fact]
    public void Latin_text_draws_ink()
    {
        GanttScene scene = TextScene(Text("label", "Ångström ç Ñ", new RectD(5, 10, 90, 10), GanttTextAlignment.Left));

        using RasterRenderResult result = Render(scene);

        Assert.True(InkCount(result) > 0, "The Latin sample painted nothing.");
    }

    [Fact]
    public void Two_renders_of_a_text_scene_are_byte_identical_and_name_the_faces()
    {
        GanttScene scene = TextScene(
            Text("regular", "Plan A", new RectD(5, 20, 60, 10), GanttTextAlignment.Left),
            Text("bold", "Plan B", new RectD(5, 50, 60, 10), GanttTextAlignment.Left, bold: true));

        using RasterRenderResult first = Render(scene);
        using RasterRenderResult second = Render(scene);

        Assert.Equal(first.Pixels.ToArray(), second.Pixels.ToArray());
        Assert.Contains("Aptos ", first.Determinism.FontIdentity, StringComparison.Ordinal);
        Assert.Contains("Aptos-Bold ", first.Determinism.FontIdentity, StringComparison.Ordinal);
        Assert.DoesNotContain("Aptos-Narrow", first.Determinism.FontIdentity, StringComparison.Ordinal);
        Assert.Contains("upem=2048", first.Determinism.FontIdentity, StringComparison.Ordinal);
    }

    private static GanttScene TextScene(params SceneText[] texts)
    {
        SceneCreationOutcome outcome = GanttScene.TryCreate(
            new RectD(0, 0, 100, 100),
            new RectD(0, 0, 100, 100),
            [.. texts],
            []);
        Assert.True(outcome.Succeeded, $"Scene creation was refused: {outcome.Refusal}.");
        return outcome.Scene!;
    }

    private static SceneText Text(
        string id,
        string value,
        RectD bounds,
        GanttTextAlignment alignment,
        bool? bold = null,
        string family = "Aptos",
        double leftInsetPt = 0) =>
        new(
            id,
            SceneOwnerId.Chart,
            ZLayer.Label,
            value,
            bounds,
            new SceneStyle("FixtureText", fontFamily: family, fontSizePt: 10, bold: bold, textColour: _black),
            alignment,
            leftInsetPt: leftInsetPt);

    private static RasterRenderResult Render(GanttScene scene) =>
        SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent);

    private static InkBounds AssertInk(RasterRenderResult result)
    {
        int minX = int.MaxValue;
        int minY = int.MaxValue;
        int maxX = int.MinValue;
        int maxY = int.MinValue;
        for (int y = 0; y < result.PixelHeight; y++)
        {
            for (int x = 0; x < result.PixelWidth; x++)
            {
                if (Alpha(result, x, y) == 0)
                {
                    continue;
                }

                minX = Math.Min(minX, x);
                minY = Math.Min(minY, y);
                maxX = Math.Max(maxX, x);
                maxY = Math.Max(maxY, y);
            }
        }

        Assert.True(minX <= maxX, "The render produced no ink.");
        return new InkBounds(minX, minY, maxX, maxY);
    }

    private static int InkCount(RasterRenderResult result)
    {
        int count = 0;
        for (int y = 0; y < result.PixelHeight; y++)
        {
            for (int x = 0; x < result.PixelWidth; x++)
            {
                if (Alpha(result, x, y) > 0)
                {
                    count++;
                }
            }
        }

        return count;
    }

    private static int Alpha(RasterRenderResult result, int x, int y) =>
        result.Pixels[(((y * result.PixelWidth) + x) * 4) + 3];

    private readonly record struct InkBounds(int MinX, int MinY, int MaxX, int MaxY);
}
