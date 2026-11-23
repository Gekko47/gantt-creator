using System.Globalization;
using System.Text;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// Draws the immutable Gantt scene's rectangles, lines, polygons, and text
/// onto a CPU raster surface in the scene's own z-order.
/// </summary>
/// <remarks>
/// <para>
/// This is the R8.2/R8.3 renderer: a pure consumer of
/// <see cref="GanttScene.Primitives"/>. It never re-derives dates, lanes,
/// colours, widths, order, text bounds, or a label side (D1/D2); the single
/// points-to-pixels mapping in <see cref="RasterTransform"/> is the only
/// arithmetic it performs, and the canvas clip is the mapped scene bounds (D3).
/// </para>
/// <para>
/// Text draws with the pinned Aptos faces resolved through
/// <see cref="PinnedFontSource"/> (ADR-0070 D3). A text-bearing scene resolves
/// every face it uses <em>before any pixel is drawn</em>: a missing or
/// mismatched face aborts with a <see cref="PinnedFontException"/> and no
/// partial result escapes. A text-free scene never touches the font, which
/// keeps every primitive, hatch, and dash render host-free.
/// </para>
/// <para>
/// A text primitive whose style resolves no <see cref="SceneStyle.TextColour"/>
/// paints nothing: the renderer never chooses a colour, the scene-first rule's
/// analogue of the null fill/stroke mapping. A group carries identity but no
/// geometry and is skipped.
/// </para>
/// <para>
/// Hatch patterns and stroke dashes are R8.3's remaining mapping obligations;
/// this renderer maps fill colour, stroke colour, stroke width, typography,
/// and text so far. A stroked primitive whose scene resolved no width draws at
/// the one-pixel raster hairline, the minimum visible stroke; a resolved width
/// is never replaced.
/// </para>
/// </remarks>
public static class SceneRasterRenderer
{
    /// <summary>Renders a scene onto a new CPU raster surface.</summary>
    /// <param name="scene">The immutable, ordered scene to draw.</param>
    /// <param name="dimensions">The validated pixel dimensions of the canvas.</param>
    /// <param name="background">The explicit canvas background.</param>
    /// <param name="fontSource">
    /// The pinned-font resolver; null selects
    /// <see cref="PinnedFontSource.Default"/>. The parameter exists so a test
    /// can inject a failure without a host, and is otherwise not needed.
    /// </param>
    /// <returns>The owned pixel result; dispose it when done.</returns>
    /// <exception cref="ArgumentNullException">Thrown when an argument is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when the scene bounds are not finite and positive or a pixel
    /// dimension is not positive.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// Thrown when a text primitive carries no resolvable font family or no
    /// positive font size — a scene-contract fault, not an environment fault.
    /// </exception>
    /// <exception cref="PinnedFontException">
    /// Thrown before any pixel is drawn when a pinned face a text primitive
    /// needs is missing, unloadable, or fails its identity check.
    /// </exception>
    public static RasterRenderResult Render(
        GanttScene scene,
        PixelDimensions dimensions,
        RasterBackground background,
        PinnedFontSource? fontSource = null)
    {
        ArgumentNullException.ThrowIfNull(scene);
        ArgumentNullException.ThrowIfNull(background);

        RasterTransform transform = RasterTransform.Create(scene.ChartBounds, dimensions);

        // R8.3 D3: collect the text work first and resolve every face it needs
        // before the result (and therefore any pixel) exists.
        List<SceneText> textPrimitives = [];
        foreach (ScenePrimitive primitive in scene.Primitives)
        {
            if (primitive is SceneText text)
            {
                textPrimitives.Add(text);
            }
        }

        Dictionary<(string Family, bool Bold), PinnedFontResolution> fonts = [];
        RasterDeterminism determinism = RasterDeterminism.Pinned;
        if (textPrimitives.Count > 0)
        {
            PinnedFontSource source = fontSource ?? PinnedFontSource.Default;
            fonts = ResolveFonts(textPrimitives, source);
            determinism = RasterDeterminism.Pinned with { FontIdentity = BuildFontIdentity(fonts) };
        }

        RasterRenderResult result = new(dimensions, determinism);
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
                        case SceneText text:
                            DrawText(canvas, transform, text, fonts);
                            break;
                        case SceneGroup:
                            // A group carries identity but no geometry.
                            break;
                        default:
                            // The scene's primitive set is closed; an unknown
                            // kind paints nothing rather than being guessed at.
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

    private static Dictionary<(string Family, bool Bold), PinnedFontResolution> ResolveFonts(
        List<SceneText> textPrimitives,
        PinnedFontSource source)
    {
        Dictionary<(string Family, bool Bold), PinnedFontResolution> fonts = [];
        foreach (SceneText text in textPrimitives)
        {
            string? family = text.Style.FontFamily;
            if (string.IsNullOrWhiteSpace(family))
            {
                throw new PinnedFontException(
                    PinnedFontRefusal.FamilyMismatch,
                    fontFamily: string.Empty,
                    observed: null);
            }

            if (text.Style.FontSizePt is not { } size || size <= 0)
            {
                throw new ArgumentException(
                    $"The text primitive '{text.PrimitiveId}' carries no positive font size; the raster renderer cannot draw it.",
                    nameof(textPrimitives));
            }

            (string Family, bool Bold) key = (family, text.Style.Bold == true);
            if (!fonts.ContainsKey(key))
            {
                PinnedFontResolution resolution = source.Resolve(family, key.Bold);
                if (!resolution.Succeeded)
                {
                    throw new PinnedFontException(resolution.Refusal!.Value, family, resolution.Observed);
                }

                fonts[key] = resolution;
            }
        }

        return fonts;
    }

    private static string BuildFontIdentity(
        Dictionary<(string Family, bool Bold), PinnedFontResolution> fonts)
    {
        StringBuilder builder = new();
        foreach (PinnedFontFace face in PinnedFontSet.Faces)
        {
            if (!fonts.TryGetValue((face.Family, face.IsBold), out PinnedFontResolution? resolution)
                || resolution.Observed is not { } observed)
            {
                continue;
            }

            if (builder.Length > 0)
            {
                _ = builder.Append("; ");
            }

            _ = builder.Append(observed.PostScriptName)
                .Append(' ')
                .Append(observed.NameTableVersion)
                .Append(" upem=")
                .Append(observed.UnitsPerEm.ToString(CultureInfo.InvariantCulture));
        }

        return builder.Length > 0 ? builder.ToString() : "none";
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

    private static void DrawText(
        SKCanvas canvas,
        RasterTransform transform,
        SceneText text,
        Dictionary<(string Family, bool Bold), PinnedFontResolution> fonts)
    {
        if (text.Style.TextColour is not { } colour)
        {
            // R8.3 D6: the scene resolved no text colour. The renderer never
            // chooses one, so there is nothing to paint.
            return;
        }

        SKRect bounds = ToSkRect(transform, text.TextBounds);
        PinnedFontResolution resolution = fonts[(text.Style.FontFamily!, text.Style.Bold == true)];
        float sizePx = transform.Length(text.Style.FontSizePt!.Value);

        SKFont font = CreateFont(resolution.Typeface!, sizePx);
        try
        {
            SKFontMetrics metrics = font.Metrics;
            float baseline = ((bounds.Top + bounds.Bottom) / 2f) - ((metrics.Ascent + metrics.Descent) / 2f);
            SKTextAlign alignment = MapAlignment(text.Alignment);
            float inset = text.LeftInsetPt > 0 ? transform.Length(text.LeftInsetPt) : 0f;
            float x = alignment == SKTextAlign.Left
                ? bounds.Left + inset
                : alignment == SKTextAlign.Center
                    ? bounds.MidX
                    : bounds.Right;

            using SKPaint paint = new() { IsAntialias = true, Color = ToSkColor(colour) };

            // The host clips a label's text to its textbox; the raster renderer
            // does the same so a glyph can never spill past the resolved box.
            _ = canvas.Save();
            canvas.ClipRect(bounds);
            canvas.DrawText(text.Text, x, baseline, alignment, font, paint);
            canvas.Restore();
        }
        finally
        {
            font.Dispose();
        }
    }

    private static SKFont CreateFont(SKTypeface typeface, float sizePx) => new(typeface, sizePx, 1f, 0f);

    private static SKTextAlign MapAlignment(GanttTextAlignment alignment) => alignment switch
    {
        GanttTextAlignment.Left => SKTextAlign.Left,
        GanttTextAlignment.Centre => SKTextAlign.Center,
        GanttTextAlignment.Right => SKTextAlign.Right,
        _ => throw new ArgumentOutOfRangeException(
            nameof(alignment),
            alignment,
            "No SkiaSharp text alignment maps this value."),
    };

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
