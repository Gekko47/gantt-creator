namespace GanttCreator.Core;

/// <summary>
/// WCAG G18 relative-luminance contrast for theme text/fill pairs (R5.6 D2).
/// </summary>
/// <remarks>
/// <para>
/// Pure and offline: operates on <see cref="ColourHex"/> only, no Office or
/// workbook types. Warning-only by contract — the caller decides whether to
/// write; this type never replaces a colour (entity-guide rule).
/// </para>
/// <para>
/// Per Microsoft <c>Accessible text requirements</c>: contrast evaluation is
/// deterministic and does not account for hue perception, so a separate
/// red-green advisory exists alongside the luminance gate.
/// </para>
/// </remarks>
public static class ThemeContrast
{
    /// <summary>The WCAG G18 minimum luminance ratio for text on background.</summary>
    public const double MinimumRatio = 4.5;

    /// <summary>
    /// Computes the WCAG contrast ratio between text and background colours.
    /// </summary>
    /// <param name="text">The label text colour.</param>
    /// <param name="background">The fill colour behind the text.</param>
    /// <returns>The contrast ratio, 1 (no contrast) to 21 (black on white).</returns>
    public static double Ratio(ColourHex text, ColourHex background)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(background);

        double textLuminance = Luminance(text);
        double backgroundLuminance = Luminance(background);
        double lighter = Math.Max(textLuminance, backgroundLuminance);
        double darker = Math.Min(textLuminance, backgroundLuminance);
        return (lighter + 0.05) / (darker + 0.05);
    }

    /// <summary>
    /// Reports whether a text/fill pair meets the minimum luminance ratio.
    /// </summary>
    /// <param name="text">The label text colour.</param>
    /// <param name="background">The fill colour behind the text.</param>
    /// <returns>True when the ratio is at least <see cref="MinimumRatio"/>.</returns>
    public static bool MeetsMinimum(ColourHex text, ColourHex background) =>
        Ratio(text, background) >= MinimumRatio - 1e-9;

    /// <summary>
    /// Reports whether a pair is a red-green combination that hue-deficient
    /// vision may confuse even when the luminance ratio passes.
    /// </summary>
    /// <param name="first">The first colour.</param>
    /// <param name="second">The second colour.</param>
    /// <returns>True when one colour is red-dominant and the other green-dominant.</returns>
    /// <remarks>
    /// Advisory only, never blocking: hatch and outline redundancies carry the
    /// distinction. Threshold-free by design — no invented hue angle.
    /// </remarks>
    public static bool IsRedGreenPair(ColourHex first, ColourHex second)
    {
        ArgumentNullException.ThrowIfNull(first);
        ArgumentNullException.ThrowIfNull(second);

        return IsRedDominant(first) != IsRedDominant(second)
            && (IsRedDominant(first) || IsRedDominant(second))
            && (IsGreenDominant(first) || IsGreenDominant(second));
    }

    private static bool IsRedDominant(ColourHex colour)
    {
        byte red = (byte)(colour.ARGB >> 16);
        byte green = (byte)(colour.ARGB >> 8);
        byte blue = (byte)colour.ARGB;
        return red > green && red > blue;
    }

    private static bool IsGreenDominant(ColourHex colour)
    {
        byte red = (byte)(colour.ARGB >> 16);
        byte green = (byte)(colour.ARGB >> 8);
        byte blue = (byte)colour.ARGB;
        return green > red && green > blue;
    }

    private static double Luminance(ColourHex colour)
    {
        double red = Linearise((byte)(colour.ARGB >> 16));
        double green = Linearise((byte)(colour.ARGB >> 8));
        double blue = Linearise((byte)colour.ARGB);
        return (0.2126 * red) + (0.7152 * green) + (0.0722 * blue);
    }

    private static double Linearise(byte channel)
    {
        double value = channel / 255.0;
        return value <= 0.03928 ? value / 12.92 : Math.Pow((value + 0.055) / 1.055, 2.4);
    }
}
