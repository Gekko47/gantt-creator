using System.Globalization;

namespace GanttCreator.Core.Scene;

/// <summary>
/// Measures label text from the real font advances of the typography families,
/// with a small proportional safety margin (ADR-0035 D1).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this replaces the flat table.</b> <see cref="FakeTextMetrics"/> charged a
/// flat 4pt per character. Measured against a live Excel textbox at 8pt, 40
/// characters of <c>W</c> occupies 287.03pt where the flat table said 160pt --
/// <b>44 percent too small</b> -- and 40 characters of <c>i</c> occupies 78.03pt
/// where it said 160pt, 105 percent too large. The small end is what clips a label:
/// the scene sized the text box from the underestimate, and the shape writer writes
/// <c>AutoSize = msoAutoSizeNone</c> with wrapping off, so the host has no way to
/// recover the difference and simply cuts the glyphs off.
/// </para>
/// <para>
/// <b>What the padding is for.</b> The font's declared advances and what Excel
/// actually reports differ slightly: the same 40 characters measured 285.47pt from
/// the table against 287.03pt from the host, and 76.41pt against 78.03pt for
/// <c>i</c>. The shortfall is roughly 0.04pt <em>per character</em>, so a
/// proportional margin is the right shape. Two percent covers the wide-glyph case
/// with room to spare and the narrow-glyph case to within about 0.1pt on a
/// 40-character label; that residual is recorded rather than hidden, and the
/// constant is a single named value so it can be raised without touching the
/// table.
/// </para>
/// <para>
/// <b>Still Office-free.</b> The advances are read from the font binaries once, by
/// the committed generator, and committed as data. Nothing here reads a font at
/// runtime, so Core keeps no Office, GDI+ or filesystem dependency and the
/// measurement stays identical on every machine.
/// </para>
/// </remarks>
public sealed class AptosTextMetrics : ITextMetrics
{
    /// <summary>
    /// The proportional safety margin applied to the summed advances.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Three percent, not the two percent first specified.</b> Two percent was
    /// measured against a live host and found INSUFFICIENT: at 8pt, 40 characters of
    /// <c>i</c> measured 77.934pt with the margin against 78.025pt reported by the
    /// host -- a residual of <b>-0.0907pt</b>, i.e. still under-measuring, which is
    /// precisely the defect this class exists to remove. The worst ratio observed is
    /// 78.02504 / 76.40625 = 1.0212 for <c>i</c>, against 1.0055 for <c>W</c>.
    /// </para>
    /// <para>
    /// Three percent clears the measured worst case with headroom. It is still only a
    /// margin: the flat table's error was up to 44 percent, which no percentage
    /// could absorb, so the advance table does the work and this absorbs the
    /// host-versus-table difference. It is a single named constant precisely so it
    /// can be revised if a future probe disagrees.
    /// </para>
    /// <para>
    /// It is applied to the SUM rather than per character, so it cannot compound on a
    /// long label.
    /// </para>
    /// </remarks>
    public const double WidthPaddingRatio = 0.03;

    /// <summary>
    /// The line height to <em>report</em> when the caller does not pin one, in
    /// multiples of the font size.
    /// </summary>
    /// <remarks>
    /// A font's own ascent plus descent, not its typographic line gap: the host box
    /// is one worksheet row tall and the text is vertically centred in it, so the
    /// reported height describes the glyphs, not the leading the host adds.
    /// </remarks>
    private const double _defaultLineHeightEm = 1.22;

    /// <summary>
    /// The advance charged for a code point the table carries no glyph for.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A FULL em, not an arbitrary constant. An earlier draft used 1409 em, and the
    /// test caught it: the widest glyph in the ASCII run is about 1863 em, so a run
    /// of unmapped characters measured NARROWER than a run of ordinary ones -- the
    /// under-estimating direction, which is the one that clips.
    /// </para>
    /// <para>
    /// One em is >= any advance a font can express for a single glyph, so the
    /// invariant holds by construction rather than by a hand-typed figure that can
    /// drift when the table is regenerated. Over-estimating costs a little space for
    /// a glyph we cannot measure; under-estimating silently truncates a label.
    /// </para>
    /// </remarks>
    private const int _fallbackAdvanceEm = FontAdvanceTable.UnitsPerEm;

    private readonly double _fontSizePt;
    private readonly bool _bold;
    private readonly bool _narrow;

    /// <summary>Initialises the metrics for the shipped label typography.</summary>
    public AptosTextMetrics()
        : this(8d, bold: false, narrow: false) { }

    /// <summary>Initialises the metrics for one face and size.</summary>
    /// <param name="fontSizePt">The font size in points.</param>
    /// <param name="bold">Whether to measure the bold face.</param>
    /// <param name="narrow">Whether to measure the Narrow family.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="fontSizePt"/> is not finite and positive.
    /// </exception>
    public AptosTextMetrics(double fontSizePt, bool bold, bool narrow)
    {
        if (!double.IsFinite(fontSizePt) || fontSizePt <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(fontSizePt),
                fontSizePt,
                "Font size must be finite and positive.");
        }

        _fontSizePt = fontSizePt;
        _bold = bold;
        _narrow = narrow;
    }

    /// <inheritdoc />
    public bool TryMeasure(string text, out TextMeasurement? measurement)
    {
        if (text is null)
        {
            measurement = null;
            return false;
        }

        var sumEm = 0L;
        foreach (var character in text)
        {
            sumEm += AdvanceEm(character);
        }

        var widthPt = sumEm * (_fontSizePt / FontAdvanceTable.UnitsPerEm);
        measurement = new TextMeasurement(widthPt * (1d + WidthPaddingRatio), _fontSizePt * _defaultLineHeightEm);
        return true;
    }

    /// <summary>
    /// The advance for one character in design units, falling back for a glyph the
    /// table does not carry.
    /// </summary>
    /// <param name="character">The character to measure.</param>
    /// <returns>The advance in design units.</returns>
    private int AdvanceEm(char character)
    {
        // A char is UTF-16, so a surrogate pair would measure as two unknown
        // characters. That over-estimates, which is the safe direction, and the
        // fallback is the widest ASCII advance rather than zero for the same reason.
        var advance = FontAdvanceTable.AdvanceFor(character, _bold, _narrow);
        return advance > 0 ? advance : _fallbackAdvanceEm;
    }

    /// <summary>
    /// The line height this instance reports, for diagnostics and tests.
    /// </summary>
    /// <returns>The line height in points.</returns>
    internal double LineHeightPt => _fontSizePt * _defaultLineHeightEm;

    /// <inheritdoc />
    public override string ToString() =>
        string.Create(CultureInfo.InvariantCulture, $"AptosTextMetrics({_fontSizePt}pt, bold={_bold}, narrow={_narrow})");
}
