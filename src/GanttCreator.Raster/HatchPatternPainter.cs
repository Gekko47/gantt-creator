using GanttCreator.Core;
using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// Draws a procurement hatch over a shape's area (R8.3 D8): explicit diagonal
/// lines clipped to the shape, at the pitch and stroke width the catalogue's
/// <c>HatchPitchPt</c>/<c>HatchLinePt</c> tokens name.
/// </summary>
/// <remarks>
/// <para>
/// The raster channel honours the two tokens exactly, where the Excel host's
/// fixed <c>MsoPatternType</c> set has no member controlling pitch or line
/// width (the R4.6 Step-0 probe). That makes the PNG <em>more</em> faithful
/// than the live chart, and the deviation is recorded rather than hidden.
/// </para>
/// <para>
/// The Step-0 probe measured <c>SKPathEffect.Create2DLine</c> as
/// matrix-convention dependent (contradictory output across variants), so this
/// painter draws the lines itself: fully deterministic, every pixel accounted
/// for, cost bounded by the shape area divided by the pitch. Lines run at 45
/// degrees; the perpendicular pitch is the token, and the x-intercept step is
/// that pitch times the square root of two.
/// </para>
/// </remarks>
internal static class HatchPatternPainter
{
    private const float _diagonalFactor = 1.41421356f;

    /// <summary>Paints a hatch clipped to a shape's area.</summary>
    /// <param name="canvas">The canvas to paint into.</param>
    /// <param name="bounds">The shape's mapped pixel bounds.</param>
    /// <param name="clipPath">The shape path for a non-rectangular clip, or null for a rectangle.</param>
    /// <param name="pattern">The resolved hatch pattern; <see cref="GanttHatchPattern.None"/> paints nothing.</param>
    /// <param name="colour">The hatch line colour — the resolved style's stroke colour.</param>
    /// <param name="pitchPt">The perpendicular hatch pitch in points.</param>
    /// <param name="lineWidthPt">The hatch line stroke width in points.</param>
    /// <param name="transform">The single points-to-pixels mapping.</param>
    public static void Paint(
        SKCanvas canvas,
        SKRect bounds,
        SKPath? clipPath,
        GanttHatchPattern pattern,
        SKColor colour,
        double pitchPt,
        double lineWidthPt,
        RasterTransform transform)
    {
        if (pattern == GanttHatchPattern.None || bounds.Width <= 0 || bounds.Height <= 0)
        {
            return;
        }

        float pitchPx = transform.Length(pitchPt);
        float lineWidthPx = transform.Length(lineWidthPt);
        float step = pitchPx * _diagonalFactor;

        using SKPathBuilder builder = new();
        if (pattern is GanttHatchPattern.ForwardDiagonal or GanttHatchPattern.Cross)
        {
            AddDiagonal(builder, bounds, step, forward: true);
        }

        if (pattern is GanttHatchPattern.BackwardDiagonal or GanttHatchPattern.Cross)
        {
            AddDiagonal(builder, bounds, step, forward: false);
        }

        using SKPath path = builder.Detach();
        using SKPaint paint = new()
        {
            Style = SKPaintStyle.Stroke,
            IsAntialias = true,
            Color = colour,
            StrokeWidth = lineWidthPx,
            StrokeCap = SKStrokeCap.Butt,
        };

        _ = canvas.Save();
        if (clipPath is null)
        {
            canvas.ClipRect(bounds);
        }
        else
        {
            canvas.ClipPath(clipPath);
        }

        canvas.DrawPath(path, paint);
        canvas.Restore();
    }

    private static void AddDiagonal(SKPathBuilder builder, SKRect bounds, float step, bool forward)
    {
        if (forward)
        {
            // x - y = c: top-left to bottom-right.
            for (float offset = bounds.Left - bounds.Height; offset <= bounds.Right + bounds.Height; offset += step)
            {
                builder.MoveTo(offset, bounds.Top);
                builder.LineTo(offset + bounds.Height, bounds.Bottom);
            }
        }
        else
        {
            // x + y = c: bottom-left to top-right.
            for (float offset = bounds.Left; offset <= bounds.Right + bounds.Height; offset += step)
            {
                builder.MoveTo(offset, bounds.Top);
                builder.LineTo(offset - bounds.Height, bounds.Bottom);
            }
        }
    }
}
