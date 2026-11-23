using GanttCreator.Core;
using GanttCreator.Core.Scene;
using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// Draws the immutable Gantt scene's rectangles, lines, and polygons onto a
/// CPU raster surface in the scene's own z-order.
/// </summary>
/// <remarks>
/// <para>
/// This is the R8.2 renderer: a pure consumer of
/// <see cref="GanttScene.Primitives"/>. It never re-derives dates, lanes,
/// colours, widths, or order (D1/D2); the single points-to-pixels mapping in
/// <see cref="RasterTransform"/> is the only arithmetic it performs, and the
/// canvas clip is the mapped scene bounds (D3).
/// </para>
/// <para>
/// Text and logical groups are skipped in this row: text rendering is R8.3
/// (after the R8.2a font decision) and a group carries identity but no
/// geometry. A scene containing either renders every other primitive and
/// leaves the skipped regions untouched.
/// </para>
/// <para>
/// Hatch patterns, stroke dashes, and typography are R8.3's mapping
/// obligations (I3); this renderer maps fill colour, stroke colour, and
/// stroke width only. A stroked primitive whose scene resolved no width draws
/// at the one-pixel raster hairline, the minimum visible stroke; a resolved
/// width is never replaced.
/// </para>
/// </remarks>
public static class SceneRasterRenderer
{
    /// <summary>Renders a scene onto a new CPU raster surface.</summary>
    /// <param name="scene">The immutable, ordered scene to draw.</param>
    /// <param name="dimensions">The validated pixel dimensions of the canvas.</param>
    /// <param name="background">The explicit canvas background.</param>
    /// <returns>The owned pixel result; dispose it when done.</returns>
    /// <exception cref="ArgumentNullException">Thrown when an argument is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when the scene bounds are not finite and positive or a pixel
    /// dimension is not positive.
    /// </exception>
    public static RasterRenderResult Render(
        GanttScene scene,
        PixelDimensions dimensions,
        RasterBackground background)
    {
        ArgumentNullException.ThrowIfNull(scene);
        ArgumentNullException.ThrowIfNull(background);

        RasterTransform transform = RasterTransform.Create(scene.ChartBounds, dimensions);

        RasterRenderResult result = new(dimensions, RasterDeterminism.Pinned);
        try
        {
            using (SKCanvas canvas = new(result.Bitmap))
            {
                SKColor clearColour = background.Fill is { } fill
                    ? ToSkColor(fill)
                    : SKColors.Transparent;
                canvas.Clear(clearColour);

                // The canvas is exactly the mapped scene bounds, so the clip is
                // stated once here rather than per primitive (D3).
                canvas.ClipRect(new SKRect(0, 0, dimensions.PixelWidth, dimensions.PixelHeight));

                foreach (ScenePrimitive primitive in scene.Primitives)
                {
                    switch (primitive)
                    {
                        case SceneRect rect:
                            DrawRect(canvas, transform, rect);
                            break;
                        case SceneLine line:
                            DrawLine(canvas, transform, line);
                            break;
                        case ScenePolygon polygon:
                            DrawPolygon(canvas, transform, polygon);
                            break;
                        case SceneText:
                        case SceneGroup:
                            // R8.3 draws text; a group carries no geometry.
                            break;
                        default:
                            // The scene's primitive set is closed today; an
                            // unknown kind paints nothing rather than being
                            // guessed at.
                            break;
                    }
                }
            }

            return result;
        }
        catch
        {
            result.Dispose();
            throw;
        }
    }

    private static void DrawRect(SKCanvas canvas, RasterTransform transform, SceneRect rect)
    {
        SKRect bounds = ToSkRect(transform, rect.Bounds);
        if (rect.Style.FillColour is { } fill)
        {
            using SKPaint paint = CreateFillPaint(fill);
            canvas.DrawRect(bounds, paint);
        }

        if (rect.Style.StrokeColour is { } stroke)
        {
            using SKPaint paint = CreateStrokePaint(stroke, rect.Style.OutlineWidthPt, transform);
            canvas.DrawRect(bounds, paint);
        }
    }

    private static void DrawLine(SKCanvas canvas, RasterTransform transform, SceneLine line)
    {
        if (line.Style.StrokeColour is not { } stroke)
        {
            return;
        }

        using SKPaint paint = CreateStrokePaint(stroke, line.Style.OutlineWidthPt, transform);
        canvas.DrawLine(
            transform.X(line.From.X),
            transform.Y(line.From.Y),
            transform.X(line.To.X),
            transform.Y(line.To.Y),
            paint);
    }

    private static void DrawPolygon(SKCanvas canvas, RasterTransform transform, ScenePolygon polygon)
    {
        bool hasFill = polygon.Style.FillColour is not null;
        bool hasStroke = polygon.Style.StrokeColour is not null;
        if (!hasFill && !hasStroke)
        {
            return;
        }

        SKPoint[] points = new SKPoint[polygon.Points.Count];
        for (int index = 0; index < points.Length; index++)
        {
            points[index] = new SKPoint(
                transform.X(polygon.Points[index].X),
                transform.Y(polygon.Points[index].Y));
        }

        using SKPathBuilder builder = new();
        builder.AddPoly(points, close: true);
        using SKPath path = builder.Detach();

        if (polygon.Style.FillColour is { } fill)
        {
            using SKPaint paint = CreateFillPaint(fill);
            canvas.DrawPath(path, paint);
        }

        if (polygon.Style.StrokeColour is { } stroke)
        {
            using SKPaint paint = CreateStrokePaint(stroke, polygon.Style.OutlineWidthPt, transform);
            canvas.DrawPath(path, paint);
        }
    }

    private static SKRect ToSkRect(RasterTransform transform, RectD bounds) => new(
        transform.X(bounds.Left),
        transform.Y(bounds.Top),
        transform.X(bounds.Right),
        transform.Y(bounds.Bottom));

    private static SKPaint CreateFillPaint(ColourHex fill) => new()
    {
        Style = SKPaintStyle.Fill,
        IsAntialias = true,
        Color = ToSkColor(fill),
    };

    private static SKPaint CreateStrokePaint(ColourHex stroke, double? widthPt, RasterTransform transform) => new()
    {
        Style = SKPaintStyle.Stroke,
        IsAntialias = true,
        Color = ToSkColor(stroke),
        StrokeWidth = widthPt is { } width && width > 0 ? transform.Length(width) : 1f,
        StrokeCap = SKStrokeCap.Butt,
        StrokeJoin = SKStrokeJoin.Miter,
    };

    private static SKColor ToSkColor(ColourHex colour) => new(
        (byte)(colour.ARGB >> 16),
        (byte)(colour.ARGB >> 8),
        (byte)colour.ARGB,
        (byte)(colour.ARGB >> 24));
}
