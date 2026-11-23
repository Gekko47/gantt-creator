using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;
using GanttCreator.Raster;

namespace GanttCreator.AddIn;

/// <summary>The evaluated state of one PNG width input against a cached export aspect.</summary>
/// <param name="Valid">Whether the input parses and converts within the safety limits.</param>
/// <param name="Request">The parsed width request, meaningful only when <paramref name="Valid"/>.</param>
/// <param name="Dimensions">The calculated pixel dimensions, meaningful only when <paramref name="Valid"/>.</param>
/// <param name="Message">The actionable message, empty when valid.</param>
internal readonly record struct PngEvaluation(
    bool Valid,
    WidthRequest Request,
    PixelDimensions Dimensions,
    string Message);

/// <summary>
/// The pure PNG width-preview math (R8.7 D1/D2/D3/D6): parse a width, convert it
/// against the cached export aspect through <see cref="ExportSize"/>, and format
/// the calculated height plus pixel dimensions.
/// </summary>
/// <remarks>
/// <para>
/// <b>Delegated math, no second formula.</b> Every number comes from
/// <see cref="ExportSize.ParseWidth"/> and <see cref="ExportSize.ToPixels"/>; this
/// type only formats their output and maps the unit/preset item ids. The equality
/// pin asserts the preview equals the parser for a table of inputs.
/// </para>
/// <para>
/// <b>Culture-invariant.</b> The width text, the calculated height and the pixel
/// dimensions all format with <see cref="CultureInfo.InvariantCulture"/>, so a
/// comma-decimal locale cannot render "29,7 cm" as "297".
/// </para>
/// </remarks>
internal static class PngPreview
{
    /// <summary>The not-available preview text before a completed render records an aspect.</summary>
    internal const string NotAvailableText = "Refresh the chart to calculate the size.";

    /// <summary>The actionable message for an empty width.</summary>
    internal const string EmptyInputMessage = "Enter a width (for example 10cm, 4in or 800px).";

    /// <summary>The default unit, centimetres (the catalogue's display convention).</summary>
    internal const ExportUnit DefaultUnit = ExportUnit.Centimetres;

    /// <summary>
    /// The default width number: the A4-landscape preset's width expressed in
    /// centimetres (29.7 cm — the first R8.4 ratified width), paired with
    /// <see cref="DefaultUnit"/>. Derived from the preset contract, never a new
    /// literal (D6).
    /// </summary>
    internal static string DefaultWidthText => CmText(SizePresets.A4Landscape.WidthPt);

    /// <summary>Converts a width in points to its centimetre text for the width box.</summary>
    /// <param name="widthPt">The width in points.</param>
    /// <returns>The invariant centimetre text the width box shows and re-parses.</returns>
    /// <remarks>
    /// Three decimals is the display precision the preset catalogue labels use
    /// as well. The residual is at most 0.0005 cm ≈ 0.0014 pt, far below the
    /// 1-pixel boundary at 300 DPI, and one formatter keeps the box text and the
    /// preset-index match in step.
    /// </remarks>
    internal static string CmText(double widthPt) =>
        (widthPt / SizePresets.PointsPerInch * ExportSize.CmPerInch)
            .ToString("0.###", CultureInfo.InvariantCulture);

    /// <summary>Evaluates a width number and its selected unit against a cached export aspect.</summary>
    /// <param name="widthText">The number the user entered in the width box.</param>
    /// <param name="unit">The unit selected in the unit dropdown.</param>
    /// <param name="aspect">The cached export aspect recorded by a completed render.</param>
    /// <returns>The evaluation; every expected failure is a result, never an exception.</returns>
    internal static PngEvaluation Evaluate(string? widthText, ExportUnit unit, ExportAspectPt aspect)
    {
        if (string.IsNullOrWhiteSpace(widthText))
        {
            return new PngEvaluation(false, default, default, EmptyInputMessage);
        }

        try
        {
            WidthRequest request = ExportSize.ParseWidth(ComposeInput(widthText, unit));
            PixelDimensions dimensions =
                ExportSize.ToPixels(request, aspect.WidthPt, aspect.HeightPt);
            return new PngEvaluation(true, request, dimensions, string.Empty);
        }
        catch (Exception ex) when (ex is FormatException or ArgumentOutOfRangeException)
        {
            // The parser and the converter each name the offending value or the
            // limit enforced, so the message is surfaced verbatim (D2).
            return new PngEvaluation(false, default, default, ex.Message);
        }
    }

    /// <summary>Whether the width number and unit form a positive, finite request.</summary>
    /// <param name="widthText">The number the user entered in the width box.</param>
    /// <param name="unit">The unit selected in the unit dropdown.</param>
    /// <returns><see langword="true"/> when the pair is a usable width request.</returns>
    /// <remarks>
    /// Parse-only: the pixel conversion needs an aspect, so this answers the
    /// input half of the readiness fact when no render has recorded one yet.
    /// </remarks>
    internal static bool IsValidInput(string? widthText, ExportUnit unit)
    {
        if (string.IsNullOrWhiteSpace(widthText))
        {
            return false;
        }

        try
        {
            _ = ExportSize.ParseWidth(ComposeInput(widthText, unit));
            return true;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    /// <summary>
    /// Composes the parser input from the number box and the unit dropdown (D3):
    /// the number means "value in the selected unit", so the unit is a label on
    /// the number and switching it never converts the value. The space is
    /// tolerated by <see cref="ExportSize"/>'s grammar and turns a unit typed into
    /// the number box into the parser's "not a valid width value" message rather
    /// than an unreadable doubled suffix.
    /// </summary>
    /// <param name="widthText">The number the user entered.</param>
    /// <param name="unit">The unit selected in the dropdown.</param>
    /// <returns>The grammar input <see cref="ExportSize.ParseWidth"/> consumes.</returns>
    private static string ComposeInput(string widthText, ExportUnit unit) =>
        string.Concat(widthText, " ", UnitLabel(unit));

    /// <summary>Formats the preview text for a valid evaluation.</summary>
    /// <param name="request">The parsed width request.</param>
    /// <param name="dimensions">The calculated pixel dimensions.</param>
    /// <returns>The width, calculated height and pixel dimensions, all invariant.</returns>
    internal static string FormatPreview(WidthRequest request, PixelDimensions dimensions)
    {
        string unit = UnitLabel(request.Unit);
        string width = request.Value.ToString("0.###", CultureInfo.InvariantCulture);
        string height = request.Unit == ExportUnit.Pixels
            ? dimensions.PixelHeight.ToString("0", CultureInfo.InvariantCulture)
            : (dimensions.PixelHeight / ExportSize.Dpi * CentimetresPerInch(request.Unit))
                .ToString("0.##", CultureInfo.InvariantCulture);

        return string.Create(
            CultureInfo.InvariantCulture,
            $"{width} {unit} wide \u2192 {height} {unit} tall \u2192 {dimensions.PixelWidth} \u00d7 {dimensions.PixelHeight} px");
    }

    /// <summary>The display label of a unit.</summary>
    /// <param name="unit">The export unit.</param>
    /// <returns><c>cm</c>, <c>in</c>, or <c>px</c>.</returns>
    internal static string UnitLabel(ExportUnit unit) => unit switch
    {
        ExportUnit.Centimetres => "cm",
        ExportUnit.Inches => "in",
        ExportUnit.Pixels => "px",
        _ => "cm",
    };

    /// <summary>The Ribbon item id of a unit's dropdown row.</summary>
    /// <param name="unit">The export unit.</param>
    /// <returns>The item id.</returns>
    internal static string UnitItemId(ExportUnit unit) => unit switch
    {
        ExportUnit.Centimetres => RibbonControlIds.PngUnitCm,
        ExportUnit.Inches => RibbonControlIds.PngUnitIn,
        ExportUnit.Pixels => RibbonControlIds.PngUnitPx,
        _ => RibbonControlIds.PngUnitCm,
    };

    /// <summary>The dropdown index of a unit.</summary>
    /// <param name="unit">The export unit.</param>
    /// <returns>The zero-based index in the centimetre, inch, pixel order.</returns>
    internal static int UnitIndex(ExportUnit unit) => unit switch
    {
        ExportUnit.Centimetres => 0,
        ExportUnit.Inches => 1,
        ExportUnit.Pixels => 2,
        _ => 0,
    };

    /// <summary>Maps a unit dropdown item id to its unit.</summary>
    /// <param name="itemId">The selected item id, or null when unprobeable.</param>
    /// <param name="unit">The resolved unit.</param>
    /// <returns><see langword="true"/> when the id names a supported unit.</returns>
    internal static bool TryParseUnitId(string? itemId, out ExportUnit unit)
    {
        switch (itemId)
        {
            case RibbonControlIds.PngUnitCm:
                unit = ExportUnit.Centimetres;
                return true;
            case RibbonControlIds.PngUnitIn:
                unit = ExportUnit.Inches;
                return true;
            case RibbonControlIds.PngUnitPx:
                unit = ExportUnit.Pixels;
                return true;
            default:
                unit = DefaultUnit;
                return false;
        }
    }

    /// <summary>The preset dropdown item id for a size preset.</summary>
    /// <param name="preset">The size preset.</param>
    /// <returns>The item id, the prefix plus the preset's durable key.</returns>
    internal static string PresetItemId(SizePreset preset) =>
        RibbonControlIds.PngSizeItemPrefix + preset.Key;

    /// <summary>Resolves a preset dropdown item id back to its size preset.</summary>
    /// <param name="itemId">The selected item id, or null when unprobeable.</param>
    /// <returns>The matching preset, or null when the id names none.</returns>
    internal static SizePreset? FindPreset(string? itemId) =>
        itemId is { } id
        && id.StartsWith(RibbonControlIds.PngSizeItemPrefix, StringComparison.Ordinal)
            ? SizePresets.All.FirstOrDefault(
                preset => string.Equals(PresetItemId(preset), id, StringComparison.Ordinal))
            : null;

    /// <summary>The preset dropdown index matching the stored width and unit, if any.</summary>
    /// <param name="widthText">The stored width text.</param>
    /// <param name="unit">The stored unit.</param>
    /// <returns>The matching preset's index, or -1 for a custom width.</returns>
    /// <remarks>
    /// A preset selection writes exactly <see cref="CmText"/> in centimetres, so
    /// the match is the text it wrote. Any other width — including a preset value
    /// typed by hand in a different unit — reads as custom rather than selecting a
    /// row the user did not choose.
    /// </remarks>
    internal static int PresetIndex(string? widthText, ExportUnit unit)
    {
        if (unit != ExportUnit.Centimetres || widthText is null)
        {
            return -1;
        }

        for (int index = 0; index < SizePresets.All.Count; index++)
        {
            if (string.Equals(CmText(SizePresets.All[index].WidthPt), widthText, StringComparison.Ordinal))
            {
                return index;
            }
        }

        return -1;
    }

    private static double CentimetresPerInch(ExportUnit unit) =>
        unit == ExportUnit.Inches ? 1.0 : ExportSize.CmPerInch;
}
