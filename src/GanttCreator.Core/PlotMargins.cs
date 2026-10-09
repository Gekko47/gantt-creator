namespace GanttCreator.Core;

/// <summary>
/// The closed plot-margin choices stored in workbook settings. Each is a
/// Windows/Excel page-margin preset, plus a user-entered custom value (R5.2).
/// </summary>
/// <remarks>
/// <para>
/// The margin is a <em>horizontal</em> page margin applied to both sides of the
/// layout width: the usable plotting area is <c>layoutWidth - 2 x margin</c>
/// (owner ruling, R5.2). It is subtracted from the width budget, never from the
/// measured text panel, so it cannot resize a user's column (R4.7H D3).
/// </para>
/// <para>
/// The three presets are the Windows/Excel horizontal defaults, authored in
/// centimetres: Narrow 0.635cm (0.25"), Normal 1.78cm (0.7"), Wide 2.54cm
/// (1.0"). <see cref="Custom"/> carries a user value through the companion
/// <c>MarginCm</c> setting.
/// </para>
/// </remarks>
public enum GanttPlotMargin
{
    /// <summary>The Windows "Narrow" horizontal margin, 0.635cm per side.</summary>
    Narrow = 0,

    /// <summary>The Windows "Normal" horizontal margin, 1.78cm per side.</summary>
    Normal = 1,

    /// <summary>The Windows "Wide" horizontal margin, 2.54cm per side.</summary>
    Wide = 2,

    /// <summary>A user-entered margin, carried by the <c>MarginCm</c> setting.</summary>
    Custom = 3,
}

/// <summary>
/// The margin catalogue: the closed preset set, the cm value of each, a strict
/// parser, and the single cm-to-points resolution used by the plot layout.
/// </summary>
public static class GanttPlotMargins
{
    /// <summary>The default margin when a workbook names none (Windows "Normal").</summary>
    public static GanttPlotMargin Default => GanttPlotMargin.Normal;

    /// <summary>The default custom margin, in centimetres, matching Normal.</summary>
    public const double DefaultCustomCm = 1.78;

    /// <summary>The smallest permitted custom margin, in centimetres.</summary>
    public const double MinimumCustomCm = 0.0;

    /// <summary>The largest permitted custom margin, in centimetres.</summary>
    public const double MaximumCustomCm = 10.0;

    /// <summary>Gets the preset's horizontal margin in centimetres.</summary>
    /// <param name="margin">The preset.</param>
    /// <returns>The per-side margin in centimetres.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="margin"/> is <see cref="GanttPlotMargin.Custom"/>
    /// (which has no fixed value) or is not a defined member.
    /// </exception>
    public static double PresetCm(GanttPlotMargin margin) =>
        margin switch
        {
            GanttPlotMargin.Narrow => 0.635,
            GanttPlotMargin.Normal => 1.78,
            GanttPlotMargin.Wide => 2.54,
            GanttPlotMargin.Custom => throw new ArgumentOutOfRangeException(
                nameof(margin),
                margin,
                "The custom margin has no fixed preset value; resolve it through MarginCm."),
            _ => throw new ArgumentOutOfRangeException(nameof(margin), margin, "Unknown margin."),
        };

    /// <summary>Attempts to parse a stored margin selector using exact ordinal text.</summary>
    /// <param name="text">The stored setting text.</param>
    /// <param name="margin">The parsed margin when successful.</param>
    /// <returns><see langword="true"/> when the text is a known margin.</returns>
    public static bool TryParse(string? text, out GanttPlotMargin margin)
    {
        switch (text)
        {
            case nameof(GanttPlotMargin.Narrow):
                margin = GanttPlotMargin.Narrow;
                return true;
            case nameof(GanttPlotMargin.Normal):
                margin = GanttPlotMargin.Normal;
                return true;
            case nameof(GanttPlotMargin.Wide):
                margin = GanttPlotMargin.Wide;
                return true;
            case nameof(GanttPlotMargin.Custom):
                margin = GanttPlotMargin.Custom;
                return true;
            default:
                margin = default;
                return false;
        }
    }

    /// <summary>
    /// Resolves the horizontal margin in points, once, from a parsed margin and a
    /// custom centimetre value. A preset ignores the custom value; a custom margin
    /// clamps an out-of-range centimetre value to the permitted band rather than
    /// refusing, so a stale stored value cannot break a Refresh.
    /// </summary>
    /// <param name="margin">The parsed margin selector.</param>
    /// <param name="customCm">The custom centimetre value, used only when <paramref name="margin"/> is <see cref="GanttPlotMargin.Custom"/>.</param>
    /// <returns>The per-side margin in points.</returns>
    public static double MarginPt(GanttPlotMargin margin, double customCm)
    {
        var cm = margin == GanttPlotMargin.Custom
            ? Math.Clamp(customCm, MinimumCustomCm, MaximumCustomCm)
            : PresetCm(margin);
        // A corrupt stored value (NaN / Infinity) survives the clamp — it is not
        // ordered against the band — so coerce it to the default rather than let
        // it propagate and break a Refresh with a negative or infinite margin.
        if (!double.IsFinite(cm))
        {
            cm = DefaultCustomCm;
        }
        return cm * 10.0 * SizePresets.PointsPerMillimetre;
    }
}
