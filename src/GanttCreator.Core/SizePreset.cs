namespace GanttCreator.Core;

/// <summary>The page orientation of a size preset.</summary>
public enum SizeOrientation
{
    /// <summary>Portrait: taller than wide.</summary>
    Portrait = 0,

    /// <summary>Landscape: wider than tall.</summary>
    Landscape = 1,
}

/// <summary>
/// The exact output size of a rendered chart, before the data panel and margins
/// consume their share (R4.7H).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why dimensions are exact rather than approximate.</b> A4 and the two
/// presentation ratios are physical or broadcast targets. A preset that rounded to
/// "about" A4 would not print to A4, and a 16:9 preset off by a fraction would
/// letterbox on a projector, so the values are declared and tested exactly.
/// </para>
/// <para>
/// <b>Why this carries no plot width.</b> The preset is the <em>budget</em>. What
/// remains after the measured text panel and the margins is the plot, and that
/// subtraction belongs to <c>PlotGeometryResolver</c> -- one authority, so a
/// command, a renderer and a Ribbon callback cannot each derive a different plot.
/// </para>
/// <para>
/// <b>Why a preset never resizes a text column.</b> The user's readable column
/// widths are the <em>input</em>, not the output. A preset too narrow for a valid
/// plot refuses rather than shrinking the panel, the plot, or the font.
/// </para>
/// </remarks>
/// <param name="Key">The durable settings/catalogue key, stored as text.</param>
/// <param name="DisplayName">The user-facing name.</param>
/// <param name="WidthPt">The exact overall width in points.</param>
/// <param name="HeightPt">The exact overall height in points.</param>
/// <param name="Orientation">The declared orientation.</param>
public sealed record SizePreset(
    string Key,
    string DisplayName,
    double WidthPt,
    double HeightPt,
    SizeOrientation Orientation)
{
    /// <summary>
    /// The exact width-to-height ratio. Derived rather than stored, so it cannot
    /// drift from the declared dimensions.
    /// </summary>
    public double AspectRatio => WidthPt / HeightPt;
}

/// <summary>The catalogue of output size presets.</summary>
public static class SizePresets
{
    /// <summary>Points per inch, the unit the presets are authored in.</summary>
    public const double PointsPerInch = 72.0;

    /// <summary>Points per millimetre.</summary>
    public const double PointsPerMillimetre = PointsPerInch / 25.4;

    /// <summary>A4 portrait, 210 x 297 mm, in points.</summary>
    public static SizePreset A4Portrait { get; } = new(
        "A4Portrait",
        "A4 (portrait)",
        210.0 * PointsPerMillimetre,
        297.0 * PointsPerMillimetre,
        SizeOrientation.Portrait);

    /// <summary>A4 landscape, 297 x 210 mm, in points.</summary>
    public static SizePreset A4Landscape { get; } = new(
        "A4Landscape",
        "A4 (landscape)",
        297.0 * PointsPerMillimetre,
        210.0 * PointsPerMillimetre,
        SizeOrientation.Landscape);

    /// <summary>A presentation slide at exactly 16:9.</summary>
    public static SizePreset Presentation16x9 { get; } = new(
        "Presentation16x9",
        "Presentation 16:9",
        1280.0,
        720.0,
        SizeOrientation.Landscape);

    /// <summary>A presentation slide at exactly 4:3.</summary>
    public static SizePreset Presentation4x3 { get; } = new(
        "Presentation4x3",
        "Presentation 4:3",
        1024.0,
        768.0,
        SizeOrientation.Landscape);

    /// <summary>Every preset, in presentation order.</summary>
    public static IReadOnlyList<SizePreset> All { get; } =
        Array.AsReadOnly([A4Portrait, A4Landscape, Presentation16x9, Presentation4x3]);

    /// <summary>
    /// The default preset when a workbook names none.
    /// </summary>
    /// <remarks>
    /// A4 <em>portrait</em>, matching the <c>SizePreset</c> default published by
    /// <see cref="GanttCatalogues.Settings"/>. The fallback and the stored default
    /// must agree: while they disagreed, a workbook whose setting was absent
    /// rendered a landscape page while a workbook whose setting was present
    /// rendered a portrait one, from the same code and the same data.
    /// </remarks>
    public static SizePreset Default => A4Portrait;

    /// <summary>
    /// Returns the preset for a settings key.
    /// </summary>
    /// <param name="key">The stored key.</param>
    /// <returns>The preset, or <see langword="null"/> when the key is unknown.</returns>
    public static SizePreset? ByKey(string? key) =>
        key is null ? null : All.FirstOrDefault(preset =>
            string.Equals(preset.Key, key, StringComparison.Ordinal));

    /// <summary>
    /// Validates the declared catalogue: keys unique, dimensions finite and
    /// positive, and each orientation consistent with its own dimensions.
    /// </summary>
    /// <returns>The validation failures; empty when the catalogue is sound.</returns>
    /// <remarks>
    /// An orientation contradicting the dimensions would let a caller label a
    /// 16:9 chart "portrait", so the pair is checked rather than trusted.
    /// </remarks>
    public static IReadOnlyList<string> ValidateCatalogue()
    {
        List<string> failures = [];

        foreach (SizePreset preset in All)
        {
            if (string.IsNullOrWhiteSpace(preset.Key))
            {
                failures.Add("A preset carries a blank key.");
            }

            if (!double.IsFinite(preset.WidthPt) || preset.WidthPt <= 0)
            {
                failures.Add($"{preset.Key}: width must be finite and positive.");
            }

            if (!double.IsFinite(preset.HeightPt) || preset.HeightPt <= 0)
            {
                failures.Add($"{preset.Key}: height must be finite and positive.");
            }

            var isWider = preset.WidthPt > preset.HeightPt;
            if (preset.Orientation == SizeOrientation.Portrait && isWider)
            {
                failures.Add($"{preset.Key}: declared portrait but is wider than tall.");
            }

            if (preset.Orientation == SizeOrientation.Landscape && !isWider)
            {
                failures.Add($"{preset.Key}: declared landscape but is not wider than tall.");
            }
        }

        if (All.Select(static preset => preset.Key).Distinct(StringComparer.Ordinal).Count() != All.Count)
        {
            failures.Add("Two presets share a key.");
        }

        return failures;
    }
}
