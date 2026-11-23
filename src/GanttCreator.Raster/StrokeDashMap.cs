using GanttCreator.Core;
using SkiaSharp;

namespace GanttCreator.Raster;

/// <summary>
/// The closed <see cref="GanttStrokeDash"/> to SkiaSharp dash mapping — the
/// ADR-0044 obligation the roadmap's Phase-8 note assigns to R8.2/R8.3.
/// </summary>
/// <remarks>
/// <para>
/// Intervals are stated in multiples of the stroke width, so a retuned width or
/// a larger export scale moves the pattern with the stroke rather than leaving
/// it fixed. The raster channel draws the pattern deterministically where
/// Excel's host-defined <c>MsoLineDashStyle</c> appearance is build-specific;
/// the goldens are the equivalence evidence, and the deviation direction is
/// recorded once, not per test.
/// </para>
/// <para>
/// The table is exhaustive over the enum and throws on an unmapped member
/// rather than drawing the nearest pattern — the same closed-mapping rule
/// <c>ExcelShapeWriter.MapDashStyle</c> follows on the Office side.
/// </para>
/// </remarks>
internal static class StrokeDashMap
{
    /// <summary>
    /// Maps a dash style to its stroke cap: a round dot needs round caps to
    /// render the dot its width, every other pattern uses butt caps.
    /// </summary>
    /// <param name="dash">The resolved dash style.</param>
    /// <returns>The cap the dash pattern is stroked with.</returns>
    public static SKStrokeCap CapFor(GanttStrokeDash dash) =>
        dash == GanttStrokeDash.RoundDot ? SKStrokeCap.Round : SKStrokeCap.Butt;

    /// <summary>
    /// Creates the dash path effect for a stroke width, or null for a solid
    /// stroke.
    /// </summary>
    /// <param name="dash">The resolved dash style.</param>
    /// <param name="strokeWidthPx">The stroke width in pixels (never below one).</param>
    /// <returns>The path effect to assign to the stroke paint, or null for solid.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when the enum member has no mapped pattern — a scene-contract
    /// fault, so the render aborts rather than drawing an invented pattern.
    /// </exception>
    public static SKPathEffect? Create(GanttStrokeDash dash, float strokeWidthPx) => dash switch
    {
        GanttStrokeDash.Solid => null,
        GanttStrokeDash.SquareDot => SKPathEffect.CreateDash([strokeWidthPx, strokeWidthPx], 0f),
        GanttStrokeDash.RoundDot => SKPathEffect.CreateDash([0.01f, strokeWidthPx * 2f], 0f),
        GanttStrokeDash.Dash => SKPathEffect.CreateDash([strokeWidthPx * 4f, strokeWidthPx * 2f], 0f),
        GanttStrokeDash.DashDot => SKPathEffect.CreateDash(
            [strokeWidthPx * 4f, strokeWidthPx * 2f, strokeWidthPx, strokeWidthPx * 2f],
            0f),
        GanttStrokeDash.LongDash => SKPathEffect.CreateDash([strokeWidthPx * 8f, strokeWidthPx * 3f], 0f),
        _ => throw new ArgumentOutOfRangeException(
            nameof(dash),
            dash,
            "No SkiaSharp dash pattern maps this value."),
    };
}
