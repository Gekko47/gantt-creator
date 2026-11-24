using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// R8.3's host-free hatch and dash pins. The reference style set resolves
/// <see cref="GanttHatchPattern.None"/> for every style, so these tests
/// synthesize hatched styles and measure the rendered pitch, orientation, and
/// dash patterns directly — they cannot ride the reference fixture.
/// </summary>
public class SceneRasterPatternTests
{
    private static readonly ColourHex _white = ColourHex.Parse("#FFFFFF");
    private static readonly ColourHex _black = ColourHex.Parse("#000000");
    private static readonly PixelDimensions _square = new(100, 100);

    [Fact]
    public void A_forward_hatch_draws_lines_at_the_catalogue_pitch()
    {
        using RasterRenderResult result = RenderRectScene(GanttHatchPattern.ForwardDiagonal);

        List<int> runStarts = DarkRunStarts(result, y: 50, fromX: 3, toX: 96);
        Assert.InRange(runStarts.Count, 12, 22);

        // The x-intercept step is pitch x sqrt(2): 4 pt at 1 px/pt = 5.66 px.
        double averagePeriod = (runStarts[^1] - runStarts[0]) / (double)(runStarts.Count - 1);
        Assert.InRange(averagePeriod, 4.6, 6.7);
    }

    [Fact]
    public void Backward_and_forward_hatches_differ_and_cross_draws_more()
    {
        using RasterRenderResult forward = RenderRectScene(GanttHatchPattern.ForwardDiagonal);
        using RasterRenderResult backward = RenderRectScene(GanttHatchPattern.BackwardDiagonal);
        using RasterRenderResult cross = RenderRectScene(GanttHatchPattern.Cross);

        Assert.NotEqual(forward.Pixels.ToArray(), backward.Pixels.ToArray());
        Assert.True(DarkCount(forward) > 0);
        Assert.True(DarkCount(backward) > 0);
        Assert.True(DarkCount(cross) > DarkCount(forward), "The cross hatch drew no more ink than one diagonal.");
    }

    [Fact]
    public void None_draws_no_hatch()
    {
        using RasterRenderResult none = RenderRectScene(GanttHatchPattern.None);
        using RasterRenderResult forward = RenderRectScene(GanttHatchPattern.ForwardDiagonal);

        // Count the interior only: the reference style strokes a 2 pt border,
        // which is ink in every variant and would mask a hatch that never drew.
        Assert.Equal(0, DarkCount(none, fromX: 3, fromY: 3, toX: 97, toY: 97));
        Assert.True(DarkCount(forward, fromX: 3, fromY: 3, toX: 97, toY: 97) > 0);
    }

    [Fact]
    public void A_hatch_with_no_stroke_colour_draws_the_fill_only()
    {
        // Like for like: the only difference between the two scenes is the
        // pattern, so the comparison isolates the hatch.
        using RasterRenderResult withoutStroke = RenderRectScene(GanttHatchPattern.ForwardDiagonal, outline: false);
        using RasterRenderResult none = RenderRectScene(GanttHatchPattern.None, outline: false);

        Assert.Equal(0, DarkCount(withoutStroke));
        Assert.Equal(none.Pixels.ToArray(), withoutStroke.Pixels.ToArray());
    }

    [Fact]
    public void The_hatch_is_clipped_to_the_shape()
    {
        SceneRect rect = new(
            "procurement",
            SceneOwnerId.Chart,
            ZLayer.ActivityBody,
            new RectD(10, 10, 30, 30),
            new SceneStyle(
                "Fixture",
                fillColour: _white,
                strokeColour: _black,
                hatchPattern: GanttHatchPattern.ForwardDiagonal));

        using RasterRenderResult result = Render(Scene(rect));

        Assert.True(DarkCount(result) > 0, "The hatched rect drew no hatch.");
        Assert.Equal(0, Alpha(result, 5, 20));
        Assert.Equal(0, Alpha(result, 45, 20));
    }

    [Theory]
    [InlineData(GanttStrokeDash.SquareDot)]
    [InlineData(GanttStrokeDash.RoundDot)]
    [InlineData(GanttStrokeDash.Dash)]
    [InlineData(GanttStrokeDash.DashDot)]
    [InlineData(GanttStrokeDash.LongDash)]
    public void Each_dash_style_draws_a_broken_line(GanttStrokeDash dash)
    {
        using RasterRenderResult result = RenderDashedLine(dash);

        // Measure painted coverage, not darkness: an unpainted pixel is
        // transparent (0,0,0,0), which a colour-sum test would read as dark
        // and collapse every gap into the run beside it.
        List<int> runs = InkRunLengths(result, y: 50, fromX: 0, toX: 100);
        Assert.True(runs.Count >= 3, $"{dash} drew {runs.Count} run(s); a broken line needs at least three.");
        Assert.True(runs.Any(length => length < 100), $"{dash} drew one continuous run.");
    }

    [Fact]
    public void A_solid_stroke_draws_one_continuous_run()
    {
        using RasterRenderResult result = RenderDashedLine(GanttStrokeDash.Solid);

        List<int> runs = InkRunLengths(result, y: 50, fromX: 0, toX: 100);
        _ = Assert.Single(runs);
        Assert.Equal(100, runs[0]);
    }

    [Fact]
    public void Round_and_square_dots_render_distinctly()
    {
        using RasterRenderResult round = RenderDashedLine(GanttStrokeDash.RoundDot);
        using RasterRenderResult square = RenderDashedLine(GanttStrokeDash.SquareDot);

        Assert.NotEqual(round.Pixels.ToArray(), square.Pixels.ToArray());
    }

    [Fact]
    public void An_unmapped_dash_style_is_refused_rather_than_drawn()
    {
        SceneLine line = new(
            "line",
            SceneOwnerId.Chart,
            ZLayer.Grid,
            new PointD(0, 50),
            new PointD(100, 50),
            new SceneStyle("Fixture", strokeColour: _black, outlineWidthPt: 2, strokeDash: (GanttStrokeDash)99));

        _ = Assert.Throws<ArgumentOutOfRangeException>(
            () => SceneRasterRenderer.Render(Scene(line), _square, RasterBackground.Transparent));
    }

    [Fact]
    public void A_hatched_and_dashed_scene_renders_byte_identically_twice()
    {
        GanttScene scene = Scene(
            new SceneRect(
                "procurement",
                SceneOwnerId.Chart,
                ZLayer.ActivityBody,
                new RectD(0, 10, 100, 30),
                new SceneStyle(
                    "Fixture",
                    fillColour: _white,
                    strokeColour: _black,
                    outlineWidthPt: 2,
                    hatchPattern: GanttHatchPattern.Cross,
                    strokeDash: GanttStrokeDash.Dash)),
            new SceneLine(
                "grid",
                SceneOwnerId.Chart,
                ZLayer.Grid,
                new PointD(0, 70),
                new PointD(100, 70),
                new SceneStyle("Fixture", strokeColour: _black, outlineWidthPt: 2, strokeDash: GanttStrokeDash.DashDot)));

        using RasterRenderResult first = Render(scene);
        using RasterRenderResult second = Render(scene);

        Assert.Equal(first.Pixels.ToArray(), second.Pixels.ToArray());
    }

    private static RasterRenderResult RenderRectScene(GanttHatchPattern pattern, bool outline = true) =>
        Render(Scene(new SceneRect(
            "procurement",
            SceneOwnerId.Chart,
            ZLayer.ActivityBody,
            new RectD(0, 0, 100, 100),
            new SceneStyle(
                "Fixture",
                fillColour: _white,
                strokeColour: outline ? _black : null,
                outlineWidthPt: 2,
                hatchPattern: pattern))));

    private static RasterRenderResult RenderDashedLine(GanttStrokeDash dash) =>
        Render(Scene(new SceneLine(
            "line",
            SceneOwnerId.Chart,
            ZLayer.Grid,
            new PointD(0, 50),
            new PointD(100, 50),
            new SceneStyle("Fixture", strokeColour: _black, outlineWidthPt: 2, strokeDash: dash))));

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

    private static RasterRenderResult Render(GanttScene scene) =>
        SceneRasterRenderer.Render(scene, _square, RasterBackground.Transparent);

    private static bool IsDark(RasterRenderResult result, int x, int y)
    {
        int offset = ((y * result.PixelWidth) + x) * 4;
        ReadOnlySpan<byte> pixels = result.Pixels;
        return pixels[offset] + pixels[offset + 1] + pixels[offset + 2] < 600;
    }

    private static bool IsInk(RasterRenderResult result, int x, int y) =>
        Alpha(result, x, y) >= 128;

    private static int Alpha(RasterRenderResult result, int x, int y) =>
        result.Pixels[(((y * result.PixelWidth) + x) * 4) + 3];

    private static int DarkCount(RasterRenderResult result) =>
        DarkCount(result, fromX: 0, fromY: 0, toX: result.PixelWidth, toY: result.PixelHeight);

    private static int DarkCount(RasterRenderResult result, int fromX, int fromY, int toX, int toY)
    {
        int count = 0;
        for (int y = fromY; y < toY; y++)
        {
            for (int x = fromX; x < toX; x++)
            {
                if (IsDark(result, x, y))
                {
                    count++;
                }
            }
        }

        return count;
    }

    private static List<int> DarkRunStarts(RasterRenderResult result, int y, int fromX, int toX)
    {
        List<int> starts = [];
        bool previousDark = false;
        for (int x = fromX; x < toX; x++)
        {
            bool dark = IsDark(result, x, y);
            if (dark && !previousDark)
            {
                starts.Add(x);
            }

            previousDark = dark;
        }

        return starts;
    }

    private static List<int> InkRunLengths(RasterRenderResult result, int y, int fromX, int toX)
    {
        List<int> runs = [];
        int length = 0;
        for (int x = fromX; x < toX; x++)
        {
            if (IsInk(result, x, y))
            {
                length++;
            }
            else if (length > 0)
            {
                runs.Add(length);
                length = 0;
            }
        }

        if (length > 0)
        {
            runs.Add(length);
        }

        return runs;
    }
}
