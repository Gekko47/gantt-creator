using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// One shape's resolved style expressed in the units the host's format members
/// take, ready for <see cref="ExcelShapeWriter.ApplyStyle"/> to write verbatim.
/// </summary>
/// <param name="FillVisible">Whether the shape has a fill at all.</param>
/// <param name="FillRgb">The packed foreground colour, or <see langword="null"/>.</param>
/// <param name="FillBackgroundRgb">
/// The packed patterned-fill background colour, or <see langword="null"/> when the
/// scene supplies none. Deliberately not defaulted to white: a renderer that
/// chooses a colour substitutes one, which the entity guide's "renderers consume,
/// never recalculate" rule forbids.
/// </param>
/// <param name="FillTransparency">The fill transparency, 0 opaque to 1 clear.</param>
/// <param name="HatchPattern">The resolved hatch pattern, or <see langword="null"/> for a solid fill.</param>
/// <param name="LineVisible">Whether the shape has a stroke.</param>
/// <param name="LineRgb">The packed stroke colour, or <see langword="null"/>.</param>
/// <param name="LineWeightPt">The stroke width in points, or <see langword="null"/>.</param>
/// <param name="LineTransparency">The stroke transparency, 0 opaque to 1 clear.</param>
/// <param name="TextRgb">
/// The packed label text colour, or <see langword="null"/> when the scene resolved
/// none. Packed with the same <c>blue * 65536 + green * 256 + red</c> order as
/// <paramref name="FillRgb"/> and <paramref name="LineRgb"/> so the three cannot
/// disagree about byte order.
/// </param>
/// <remarks>
/// <para>
/// Every value is a primitive or a Core-owned type, so this record and the mapper
/// that produces it never name an interop type. That is what lets the whole
/// token-to-property matrix be asserted without a live host, and it keeps the
/// interop surface confined to the one adapter that must call COM.
/// </para>
/// <para>
/// A <see langword="null"/> colour and a <see langword="false"/> visibility flag are
/// different facts and both are preserved. The first means the scene resolved no
/// colour; the second means the entity genuinely has no fill or stroke. A mapper
/// that collapsed them would make "no stroke resolved" indistinguishable from
/// "stroke is white", which is the substitution rule again.
/// </para>
/// </remarks>
public sealed record OfficeShapeStyle(
    bool FillVisible,
    int? FillRgb,
    int? FillBackgroundRgb,
    float FillTransparency,
    GanttHatchPattern? HatchPattern,
    bool LineVisible,
    int? LineRgb,
    float? LineWeightPt,
    float LineTransparency,
    int? TextRgb);

/// <summary>
/// Maps a scene-resolved shape request onto the host's fill and line property
/// values. R4.6 D3: the token is written verbatim, with no gamma correction, no
/// theme colour, and no default substituted for an absent token.
/// </summary>
public static class OfficeStyleMapper
{
    /// <summary>
    /// Maps one request's style members onto host-shaped values.
    /// </summary>
    /// <param name="request">The request whose style members are mapped.</param>
    /// <returns>The host-shaped style values.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="request"/> is <see langword="null"/>.</exception>
    public static OfficeShapeStyle Map(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        // A line has no interior: the host's line format is its entire
        // appearance, and writing a fill to it would be a value the scene never
        // resolved for that entity. Section 24's delineator and section 16's
        // critical overlay are both stroke-only, and their scene styles carry a
        // null fill - this arm keeps that null from being turned into a fill.
        //
        // The colour and pattern are nulled, not merely flagged invisible: a
        // record that reported FillVisible = false while still carrying a FillRgb
        // would be a trap for the next caller, and "there is no fill colour"
        // and "the fill is switched off" are the same fact here.
        var fillApplies = CarriesFill(request.Kind);
        GanttHatchPattern? hatch = !fillApplies || request.HatchPattern == GanttHatchPattern.None
            ? null
            : request.HatchPattern;

        return new OfficeShapeStyle(
            FillVisible: fillApplies && request.FillColour is not null,
            FillRgb: fillApplies && request.FillColour is { } fill ? ToOfficeRgb(fill) : null,
            FillBackgroundRgb: null,
            FillTransparency: fillApplies && request.FillColour is { } fillAlpha
                ? ToTransparency(fillAlpha)
                : 0f,
            HatchPattern: hatch,
            LineVisible: request.StrokeColour is not null,
            LineRgb: request.StrokeColour is { } stroke ? ToOfficeRgb(stroke) : null,
            LineWeightPt: request.LineWidthPt is { } width
                ? GeometryMath.SnapToDisplayPrecision(width)
                : null,
            LineTransparency: request.StrokeColour is { } strokeAlpha
                ? ToTransparency(strokeAlpha)
                : 0f,
            // The text colour is packed by the same ToOfficeRgb the fill and stroke
            // use, so the byte order cannot differ between the three. It is
            // deliberately not gated on CarriesFill: a line carries no text, but
            // the text-bearing kinds are exactly the ones a caller is most likely
            // to ask this of, and a null request colour must stay null rather than
            // be turned into black here.
            TextRgb: request.TextColour is { } textColour ? ToOfficeRgb(textColour) : null);
    }

    /// <summary>
    /// Packs a resolved colour into the host's integer colour value.
    /// </summary>
    /// <param name="colour">The resolved colour.</param>
    /// <returns>The packed colour.</returns>
    /// <remarks>
    /// <para>
    /// The host packs a colour least-significant-byte first as
    /// <c>blue * 65536 + green * 256 + red</c> (MS-OI29500 6.1.2.7.1.11, the
    /// <c>RGB</c> function's return value, which is also what
    /// <c>ColorFormat.RGB</c> stores). The arithmetic below spells that out
    /// rather than writing a hex mask, because the byte order is the opposite of
    /// the <c>#RRGGBB</c> form the tokens are authored in and getting it wrong
    /// swaps red and blue with nothing to catch it.
    /// </para>
    /// <para>
    /// Theme colours are never used. <c>ColorFormat.ObjectThemeColor</c> and
    /// <c>SchemeColor</c> exist and are deliberately not touched, so the written
    /// value is the token itself rather than whatever the workbook's theme
    /// currently resolves that token to (D3).
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="colour"/> is <see langword="null"/>.</exception>
    public static int ToOfficeRgb(ColourHex colour)
    {
        ArgumentNullException.ThrowIfNull(colour);

        var argb = colour.ARGB;
        var red = (int)((argb >> 16) & 0xFF);
        var green = (int)((argb >> 8) & 0xFF);
        var blue = (int)(argb & 0xFF);

        return (blue * 65536) + (green * 256) + red;
    }

    /// <summary>
    /// Converts a resolved colour's alpha channel to the host's transparency scale.
    /// </summary>
    /// <param name="colour">The resolved colour.</param>
    /// <returns>0 for an opaque colour, approaching 1 as alpha approaches 0.</returns>
    /// <remarks>
    /// The host's <c>FillFormat.Transparency</c> and <c>LineFormat.Transparency</c>
    /// are <c>Single</c> values on 0-to-1 where 0 is opaque - the inverse sense
    /// of the token's alpha, hence the subtraction. A <c>#RRGGBB</c> token parses
    /// to an alpha of <c>FF</c> and therefore to 0, so the common opaque case needs
    /// no special case and the two token forms cannot disagree.
    /// </remarks>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="colour"/> is <see langword="null"/>.</exception>
    public static float ToTransparency(ColourHex colour)
    {
        ArgumentNullException.ThrowIfNull(colour);

        var alpha = (colour.ARGB >> 24) & 0xFF;
        return 1f - (alpha / 255f);
    }

    /// <summary>
    /// Determines whether a shape kind has a fill the host can express.
    /// </summary>
    /// <param name="kind">The shape kind.</param>
    /// <returns><see langword="true"/> for the kinds with an interior.</returns>
    /// <remarks>
    /// Closed over <see cref="OfficeShapeKind"/> on purpose. The default arm is
    /// <see langword="false"/> rather than a silent <see langword="true"/>, so a
    /// kind added later is treated as having no fill until this says otherwise -
    /// the conservative direction, since a wrongly written fill is a visible
    /// substitution while a wrongly omitted one is corrected by the next test run.
    /// </remarks>
    public static bool CarriesFill(OfficeShapeKind kind) =>
        kind is OfficeShapeKind.Rectangle
            or OfficeShapeKind.TextBox
            or OfficeShapeKind.Diamond;
}
