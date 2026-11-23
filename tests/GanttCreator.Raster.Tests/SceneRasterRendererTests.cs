using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// Acceptance pins for the R8.2 Skia primitive renderer: exact-pixel regions
/// for rectangles, lines, and polygons; the style-to-draw mapping; z-order;
/// clipping; determinism; and the no-recalculation positive.
/// </summary>
/// <remarks>
/// Every assertion reads the rendered BGRA bytes directly rather than a
/// re-encoded image, so the pins observe exactly what the renderer produced.
/// Interior/core regions are asserted exactly (R8.2 I7); antialiased edges are
/// covered by the determinism pin, never by a geometry tolerance.
/// </remarks>
public class SceneRasterRendererTests
{
    private static readonly ColourHex _blue = ColourHex.Parse("#3366CC");
    private static readonly ColourHex _red = ColourHex.Parse("#FF0000");
    private static readonly ColourHex _green = ColourHex.Parse("#00FF00");
    private static readonly ColourHex _black = ColourHex.Parse("#000000");
    private static readonly ColourHex _white = ColourHex.Parse("#FFFFFF");
    private static readonly ColourHex _orange = ColourHex.Parse("#FFAA00");

    private static readonly PixelDimensions _square = new(100, 100);

    [Fact]
    public void Renders_a_rectangle_at_known_coordinates_with_exact_interior_pixels()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("rect", 10, 20, 30, 40, ZLayer.ActivityBody, _blue));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 25, 40, _blue);
        AssertPixel(result, 5, 5, Transparent);
        AssertPixel(result, 9, 40, Transparent);
        AssertPixel(result, 40, 40, Transparent);
    }

    [Fact]
    public void Renders_a_line_at_known_coordinates_with_exact_core_pixels()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new SceneLine(
                "line",
                SceneOwnerId.Chart,
                ZLayer.Grid,
                new PointD(10, 50),
                new PointD(90, 50),
                new SceneStyle("Fixture", strokeColour: _black, outlineWidthPt: 2)));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 50, 49, _black);
        AssertPixel(result, 50, 50, _black);
        AssertPixel(result, 50, 47, Transparent);
        AssertPixel(result, 5, 50, Transparent);
    }

    [Fact]
    public void Renders_a_polygon_at_known_coordinates_with_exact_interior_pixels()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new ScenePolygon(
                "polygon",
                SceneOwnerId.Chart,
                ZLayer.Milestone,
                [new PointD(10, 10), new PointD(50, 10), new PointD(30, 40)],
                new SceneStyle("Fixture", fillColour: _red)));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 30, 15, _red);
        AssertPixel(result, 5, 5, Transparent);
        AssertPixel(result, 49, 15, Transparent);
    }

    [Fact]
    public void Later_scene_primitives_paint_over_earlier_ones_at_overlap_pixels()
    {
        // Inserted in reverse layer order on purpose: the renderer must consume
        // the scene's ordered list, not the construction order.
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("label", 50, 50, 10, 10, ZLayer.Label, _blue),
            FillRect("background", 40, 40, 30, 30, ZLayer.Background, _red),
            FillRect("body", 45, 45, 20, 20, ZLayer.ActivityBody, _green));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 42, 42, _red);
        AssertPixel(result, 47, 47, _green);
        AssertPixel(result, 52, 52, _blue);
    }

    [Fact]
    public void A_primitive_outside_the_scene_bounds_paints_nothing()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("outside", 120, 120, 10, 10, ZLayer.ActivityBody, _blue));

        using RasterRenderResult result = Render(scene);

        AssertAllTransparent(result);
    }

    [Fact]
    public void A_partially_clipped_primitive_paints_only_the_inside_region()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("clipped", -20, 40, 40, 20, ZLayer.ActivityBody, _blue));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 5, 50, _blue);
        AssertPixel(result, 25, 50, Transparent);
    }

    [Fact]
    public void Two_renders_of_the_same_scene_and_width_are_byte_identical()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("background", 0, 0, 100, 100, ZLayer.Background, _white),
            FillRect("body", 10, 20, 30, 40, ZLayer.ActivityBody, _blue),
            new ScenePolygon(
                "diamond",
                SceneOwnerId.Chart,
                ZLayer.Milestone,
                [new PointD(70, 25), new PointD(75, 30), new PointD(70, 35), new PointD(65, 30)],
                new SceneStyle("Fixture", fillColour: _black)));

        using RasterRenderResult first = Render(scene);
        using RasterRenderResult second = Render(scene);

        Assert.Equal(first.Pixels.ToArray(), second.Pixels.ToArray());
        Assert.Equal("cpu-raster", first.Determinism.SurfaceKind);
        Assert.True(first.Determinism.Antialias);
        Assert.False(string.IsNullOrWhiteSpace(first.Determinism.SkiaVersion));
    }

    [Fact]
    public void The_renderer_draws_the_scene_bounds_verbatim_and_never_recomputes_geometry()
    {
        // The bounds are deliberately not derivable from any date or lane
        // arithmetic: a renderer that re-derived geometry would not land on
        // these pixels. The renderer must draw the resolved bounds verbatim.
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("odd", 13.7, 20.3, 17.3, 9.9, ZLayer.ActivityBody, _blue));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 20, 25, _blue);
        AssertPixel(result, 13, 25, Transparent);
        AssertPixel(result, 31, 25, Transparent);
    }

    [Fact]
    public void A_transparent_background_leaves_alpha_zero()
    {
        GanttScene scene = CreateScene(new RectD(0, 0, 100, 100));

        using RasterRenderResult result = SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent);

        AssertPixel(result, 0, 0, Transparent);
    }

    [Fact]
    public void An_explicit_colour_fills_the_canvas()
    {
        GanttScene scene = CreateScene(new RectD(0, 0, 100, 100));

        using RasterRenderResult result = SceneRasterRenderer.Render(
            scene,
            _square,
            RasterBackground.FromColour(_white));

        AssertPixel(result, 0, 0, _white);
        AssertPixel(result, 99, 99, _white);
    }

    [Fact]
    public void A_scene_with_text_and_groups_renders_only_the_non_text_primitives()
    {
        SceneRect body = FillRect("body", 50, 50, 10, 10, ZLayer.ActivityBody, _blue);
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new SceneText(
                "label",
                SceneOwnerId.Chart,
                ZLayer.Label,
                "Hello",
                new RectD(10, 10, 30, 10),
                new SceneStyle("FixtureText", fontFamily: "Aptos", fontSizePt: 10),
                GanttTextAlignment.Left),
            new SceneGroup("group", SceneOwnerId.Chart, ZLayer.Label, ["body"]),
            body);

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 55, 55, _blue);
        AssertPixel(result, 20, 15, Transparent);
    }

    [Fact]
    public void A_stroke_only_rectangle_is_not_filled()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new SceneRect(
                "stroke",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new RectD(10, 10, 30, 30),
                new SceneStyle("Fixture", strokeColour: _red, outlineWidthPt: 2)));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 10, 25, _red);
        AssertPixel(result, 25, 25, Transparent);
    }

    [Fact]
    public void A_rectangle_with_fill_and_stroke_draws_fill_then_stroke()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new SceneRect(
                "both",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new RectD(10, 10, 30, 30),
                new SceneStyle("Fixture", fillColour: _green, strokeColour: _red, outlineWidthPt: 2)));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 25, 25, _green);
        AssertPixel(result, 10, 25, _red);
    }

    [Fact]
    public void A_primitive_with_neither_fill_nor_stroke_paints_nothing()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new SceneRect(
                "empty",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new RectD(10, 10, 30, 30),
                new SceneStyle("Fixture")));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 25, 25, Transparent);
        AssertPixel(result, 10, 25, Transparent);
    }

    [Fact]
    public void A_stroke_with_no_resolved_width_draws_the_one_pixel_hairline()
    {
        // The single rounding pass maps the half-pixel Y to the integer pixel
        // row 50, so a one-pixel stroke straddles rows 49 and 50 at partial
        // coverage -- visibly thinner than a two-pixel stroke, which would fill
        // both rows. The pin is behavioural (visible, partial, confined) rather
        // than an exact AA edge value (R8.2 I7).
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new SceneLine(
                "hairline",
                SceneOwnerId.Chart,
                ZLayer.Grid,
                new PointD(10, 50.5),
                new PointD(90, 50.5),
                new SceneStyle("Fixture", strokeColour: _black)));

        using RasterRenderResult result = Render(scene);

        Assert.InRange(Alpha(result, 50, 49), 1, 254);
        Assert.InRange(Alpha(result, 50, 50), 1, 254);
        Assert.Equal(0, Alpha(result, 50, 48));
        Assert.Equal(0, Alpha(result, 50, 51));
    }

    [Fact]
    public void A_filled_scene_rect_renders_filled_whatever_its_entity_type()
    {
        // Checklist E revision 8: the critical interval is a filled rectangle
        // and every renderer must draw it filled. The renderer has no
        // entity-type branch, so this pins the outcome rather than a branch.
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("critical", 20, 20, 20, 10, ZLayer.CriticalOverlay, _orange, GanttEntityType.CriticalInterval));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 30, 25, _orange);
    }

    [Fact]
    public void Half_pixel_coordinates_round_once_to_even()
    {
        // 10.5 and 20.5 both round to even (10 and 20), matching ExportSize's
        // dimension rounding. Away-from-zero would start the rect at 11.
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("half", 10.5, 20, 10, 10, ZLayer.ActivityBody, _blue));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 10, 25, _blue);
        AssertPixel(result, 9, 25, Transparent);
        AssertPixel(result, 20, 25, Transparent);
    }

    [Fact]
    public void The_mapped_scene_bounds_fill_the_canvas_exactly()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            FillRect("full", 0, 0, 100, 100, ZLayer.Background, _white));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 0, 0, _white);
        AssertPixel(result, 99, 99, _white);
    }

    [Fact]
    public void Render_rejects_a_null_scene()
    {
        _ = Assert.Throws<ArgumentNullException>(
            () => SceneRasterRenderer.Render(null!, _square, RasterBackground.Transparent));
    }

    [Fact]
    public void Render_rejects_a_null_background()
    {
        GanttScene scene = CreateScene(new RectD(0, 0, 100, 100));

        _ = Assert.Throws<ArgumentNullException>(
            () => SceneRasterRenderer.Render(scene, _square, null!));
    }

    [Fact]
    public void Render_rejects_a_scene_with_no_width()
    {
        GanttScene scene = CreateScene(new RectD(0, 0, 0, 100));

        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent));
    }

    [Fact]
    public void Render_rejects_non_positive_pixel_dimensions()
    {
        GanttScene scene = CreateScene(new RectD(0, 0, 100, 100));

        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => SceneRasterRenderer.Render(scene, new PixelDimensions(0, 100), RasterBackground.Transparent));
    }

    [Fact]
    public void A_polygon_with_neither_fill_nor_stroke_paints_nothing()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new ScenePolygon(
                "empty-polygon",
                SceneOwnerId.Chart,
                ZLayer.Milestone,
                [new PointD(10, 10), new PointD(50, 10), new PointD(30, 40)],
                new SceneStyle("Fixture")));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 30, 15, Transparent);
    }

    [Fact]
    public void A_line_with_no_stroke_colour_paints_nothing()
    {
        GanttScene scene = CreateScene(
            new RectD(0, 0, 100, 100),
            new SceneLine(
                "fill-only-line",
                SceneOwnerId.Chart,
                ZLayer.Grid,
                new PointD(10, 50),
                new PointD(90, 50),
                new SceneStyle("Fixture", fillColour: _blue)));

        using RasterRenderResult result = Render(scene);

        AssertPixel(result, 50, 50, Transparent);
    }

    [Fact]
    public void Background_from_a_null_colour_is_rejected() =>
        _ = Assert.Throws<ArgumentNullException>(() => RasterBackground.FromColour(null!));

    [Fact]
    public void Background_values_compare_by_value()
    {
        Assert.True(RasterBackground.Transparent == RasterBackground.Transparent);
        Assert.True(RasterBackground.Transparent != RasterBackground.FromColour(_white));
        Assert.Equal(RasterBackground.Transparent, RasterBackground.Transparent);
        Assert.NotEqual(RasterBackground.Transparent, RasterBackground.FromColour(_white));
        Assert.Equal(
            RasterBackground.FromColour(_white).GetHashCode(),
            RasterBackground.FromColour(_white).GetHashCode());
    }

    [Fact]
    public void Determinism_descriptors_compare_by_value()
    {
        Assert.True(RasterDeterminism.Pinned == RasterDeterminism.Pinned);
        Assert.Equal(RasterDeterminism.Pinned, RasterDeterminism.Pinned);
        Assert.Equal(RasterDeterminism.Pinned.GetHashCode(), RasterDeterminism.Pinned.GetHashCode());
    }

    [Fact]
    public void Dispose_is_idempotent()
    {
        GanttScene scene = CreateScene(new RectD(0, 0, 100, 100));
        RasterRenderResult result = Render(scene);

        result.Dispose();
        result.Dispose();
    }

    private static ColourHex Transparent { get; } = ColourHex.Parse("#00000000");

    private static GanttScene CreateScene(RectD bounds, params ScenePrimitive[] primitives)
    {
        SceneCreationOutcome outcome = GanttScene.TryCreate(bounds, bounds, primitives, []);
        Assert.True(outcome.Succeeded, $"Scene creation was refused: {outcome.Refusal}.");
        return outcome.Scene!;
    }

    private static SceneRect FillRect(
        string id,
        double x,
        double y,
        double width,
        double height,
        ZLayer layer,
        ColourHex fill,
        GanttEntityType? entityType = null) =>
        new(id, SceneOwnerId.Chart, layer, new RectD(x, y, width, height), new SceneStyle("Fixture", fillColour: fill), entityType);

    private static RasterRenderResult Render(GanttScene scene) =>
        SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent);

    private static void AssertPixel(RasterRenderResult result, int x, int y, ColourHex expected)
    {
        int offset = ((y * result.PixelWidth) + x) * 4;
        ReadOnlySpan<byte> pixels = result.Pixels;

        Assert.Equal((byte)(expected.ARGB & 0xFF), pixels[offset]);
        Assert.Equal((byte)((expected.ARGB >> 8) & 0xFF), pixels[offset + 1]);
        Assert.Equal((byte)((expected.ARGB >> 16) & 0xFF), pixels[offset + 2]);
        Assert.Equal((byte)((expected.ARGB >> 24) & 0xFF), pixels[offset + 3]);
    }

    private static int Alpha(RasterRenderResult result, int x, int y) =>
        result.Pixels[(((y * result.PixelWidth) + x) * 4) + 3];

    private static void AssertAllTransparent(RasterRenderResult result)
    {
        ReadOnlySpan<byte> pixels = result.Pixels;
        for (int offset = 3; offset < pixels.Length; offset += 4)
        {
            Assert.Equal(0, pixels[offset]);
        }
    }
}
