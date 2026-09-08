using System.Globalization;

namespace GanttCreator.Raster;

/// <summary>
/// The unit of the user-entered export width.
/// </summary>
public enum ExportUnit
{
    /// <summary>Centimetres. Converted to pixels at 300 DPI.</summary>
    Centimetres,

    /// <summary>Inches. Converted to pixels at 300 DPI.</summary>
    Inches,

    /// <summary>Raw pixels. The entered value is the final width.</summary>
    Pixels
}

/// <summary>
/// A parsed, validated export-width request.
/// </summary>
/// <param name="Unit">The unit of <paramref name="Value"/>.</param>
/// <param name="Value">The positive width in <paramref name="Unit"/>.</param>
public readonly record struct WidthRequest(ExportUnit Unit, double Value);

/// <summary>
/// The calculated pixel dimensions of an export, derived from a
/// width request and the scene aspect ratio.
/// </summary>
/// <param name="PixelWidth">The rounded pixel width.</param>
/// <param name="PixelHeight">The rounded pixel height.</param>
public readonly record struct PixelDimensions(int PixelWidth, int PixelHeight);

/// <summary>
/// Parses and converts export-width requests. All parsing is
/// culture-invariant because user input such as "2.5in" must not
/// be interpreted differently on month-day-year vs day-month-year
/// locales. See docs/07-GANTT-ENTITY-GUIDE.md coordinate policy.
/// </summary>
public static class ExportSize
{
    /// <summary>Rendering resolution: 300 dots per inch.</summary>
    public const double Dpi = 300.0;

    /// <summary>Inches per centimetre.</summary>
    public const double CmPerInch = 2.54;

    /// <summary>
    /// Practical upper bound for any single exported pixel dimension. This
    /// sits far below int.MaxValue but above every rasterizer's bitmap
    /// limit, so absurd user input fails here with an actionable message
    /// instead of surfacing as a GDI+/Skia allocation failure at encode time.
    /// </summary>
    public const double MaxPixelDimension = 65_535;

    /// <summary>
    /// Parses a width string such as "10cm", "4in", or "800px"
    /// (case-insensitive, surrounding whitespace ignored) into a
    /// <see cref="WidthRequest"/>.
    /// </summary>
    /// <exception cref="FormatException">Unknown unit or invalid value.</exception>
    public static WidthRequest ParseWidth(string input)
    {
        ArgumentNullException.ThrowIfNull(input);
        input = input.Trim();

        // Switch expression: a suffix is matched exactly once and
        // unknown units fall through to a single failure point.
        return input switch
        {
            var s when s.EndsWith("cm", StringComparison.OrdinalIgnoreCase) =>
                new WidthRequest(ExportUnit.Centimetres, ParseValue(s[..^2])),
            var s when s.EndsWith("in", StringComparison.OrdinalIgnoreCase) =>
                new WidthRequest(ExportUnit.Inches, ParseValue(s[..^2])),
            var s when s.EndsWith("px", StringComparison.OrdinalIgnoreCase) =>
                new WidthRequest(ExportUnit.Pixels, ParseValue(s[..^2])),
            _ => throw new FormatException($"Unknown unit in '{input}'.")
        };
    }

    /// <summary>
    /// Converts a width request and scene dimensions (in points) to
    /// pixel dimensions. The height preserves the scene aspect ratio;
    /// it is calculated, never entered independently.
    /// </summary>
    /// <param name="request">The parsed width request. <see cref="WidthRequest.Value"/> must be finite and greater than zero.</param>
    /// <param name="sceneWidthPt">The scene width in points. Must be finite and greater than zero.</param>
    /// <param name="sceneHeightPt">The scene height in points. Must be finite and non-negative.</param>
    /// <exception cref="ArgumentOutOfRangeException">A dimension value is
    /// not finite, or sceneWidthPt is not greater than zero, or the
    /// computed pixel dimensions exceed the practical export limit of
    /// 65,535 px.
    /// </exception>
    public static PixelDimensions ToPixels(WidthRequest request, double sceneWidthPt, double sceneHeightPt)
    {
        if (!double.IsFinite(request.Value) || request.Value <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(request), request.Value, "Width request Value must be finite and greater than zero.");
        }

        if (!double.IsFinite(sceneWidthPt) || sceneWidthPt <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sceneWidthPt), sceneWidthPt, "Scene width must be finite and greater than zero.");
        }

        if (!double.IsFinite(sceneHeightPt) || sceneHeightPt < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sceneHeightPt), sceneHeightPt, "Scene height must be finite and non-negative.");
        }

        var pixelWidth = request.Unit switch
        {
            ExportUnit.Centimetres => request.Value / CmPerInch * Dpi,
            ExportUnit.Inches => request.Value * Dpi,
            ExportUnit.Pixels => request.Value,
            _ => throw new ArgumentOutOfRangeException(nameof(request))
        };

        var pixelHeight = pixelWidth * sceneHeightPt / sceneWidthPt;

        if (!double.IsFinite(pixelWidth) || pixelWidth > MaxPixelDimension)
        {
            throw new ArgumentOutOfRangeException(
                nameof(request), pixelWidth,
                $"Computed pixel width exceeds the practical export limit of {MaxPixelDimension} px.");
        }

        if (!double.IsFinite(pixelHeight) || pixelHeight <= 0 || pixelHeight > MaxPixelDimension)
        {
            // Zero/negative heights must be rejected here too: a zero height
            // would otherwise round to PixelDimensions with a zero height.
            throw new ArgumentOutOfRangeException(
                nameof(sceneHeightPt), pixelHeight,
                $"Computed pixel height must be finite, greater than zero, and within the practical export limit of {MaxPixelDimension} px.");
        }

        var pxWidth = (int)Math.Round(pixelWidth);
        var pxHeight = (int)Math.Round(pixelHeight);
        return new PixelDimensions(pxWidth, pxHeight);
    }

    private static double ParseValue(string text)
    {
        text = text.Trim();

        // Tuple switch: validates emptiness, parseability, sign, and
        // finiteness in one expression. Zero is rejected: a zero-width
        // export is meaningless and would surface as a rasterizer failure
        // far from the cause (decision recorded in the W4 work item).
        return (text.Length, double.TryParse(text, NumberStyles.Float,
                CultureInfo.InvariantCulture, out var value), value <= 0, double.IsFinite(value)) switch
        {
            (0, _, _, _) => throw new FormatException("Width value is empty."),
            (_, false, _, _) => throw new FormatException($"'{text}' is not a valid width value."),
            (_, true, true, _) => throw new FormatException("Width value must be greater than zero."),
            (_, true, _, false) => throw new FormatException($"'{text}' is not a finite number."),
            _ => value
        };
    }
}
