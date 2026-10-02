using GanttCreator.Core.Scene;
using Xunit.Abstractions;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Contract tests for <see cref="AptosTextMetrics"/> and the generated
/// <c>FontAdvanceTable</c> (ADR-0035 D1).
/// </summary>
/// <remarks>
/// <para>
/// The tests that matter here are the ones the flat 4pt-per-character table could
/// NOT pass: that different glyphs measure differently, that the width scales with
/// the font size, and that bold differs from regular. A flat table satisfies none of
/// them, so they are the non-vacuity guard for this change.
/// </para>
/// <para>
/// The reference figures come from a live Excel probe against a real textbox at
/// 8pt: 40 characters of <c>W</c> reported 287.03pt and 40 characters of <c>i</c>
/// 78.03pt. The table's own sums are 285.47pt and 76.41pt, so the 2 percent margin
/// exists to cover that gap.
/// </para>
/// </remarks>
public sealed class AptosTextMetricsTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static double Width(AptosTextMetrics metrics, string text)
    {
        Assert.True(metrics.TryMeasure(text, out TextMeasurement? measurement));
        return Assert.IsType<TextMeasurement>(measurement).WidthPt;
    }

    /// <summary>Wide and narrow glyphs measure differently, which a flat advance cannot do.</summary>
    [Fact]
    public void Wide_and_narrow_glyphs_measure_differently()
    {
        var metrics = new AptosTextMetrics();

        double wide = Width(metrics, new string('W', 40));
        double narrow = Width(metrics, new string('i', 40));

        _output.WriteLine($"W x40 = {wide:0.###}pt, i x40 = {narrow:0.###}pt");

        Assert.True(wide > narrow * 3, "A 40-char 'W' run must be far wider than a 40-char 'i' run.");

        // The flat table gave both exactly 160pt. This is the assertion that fails if
        // the table is ever replaced by a per-character constant again.
        Assert.NotEqual(wide, narrow);
    }

    /// <summary>The measured width scales with the font size, which a fixed advance cannot do.</summary>
    [Fact]
    public void Width_scales_with_the_font_size()
    {
        double atEight = Width(new AptosTextMetrics(8d, bold: false, narrow: false), new string('W', 40));
        double atSixteen = Width(new AptosTextMetrics(16d, bold: false, narrow: false), new string('W', 40));

        Assert.Equal(atEight * 2d, atSixteen, 3);
    }

    /// <summary>Bold measures wider than regular: the table carries both faces.</summary>
    [Fact]
    public void Bold_measures_wider_than_regular()
    {
        double regular = Width(new AptosTextMetrics(8d, bold: false, narrow: false), new string('n', 40));
        double bold = Width(new AptosTextMetrics(8d, bold: true, narrow: false), new string('n', 40));

        Assert.True(bold > regular, $"Bold ({bold:0.###}) must exceed regular ({regular:0.###}).");
    }

    /// <summary>
    /// Narrow measures narrower than the default family, which is why the face has to
    /// be selected rather than assumed.
    /// </summary>
    [Fact]
    public void Narrow_measures_narrower_than_the_default_family()
    {
        double wide = Width(new AptosTextMetrics(8d, bold: false, narrow: false), new string('n', 40));
        double narrow = Width(new AptosTextMetrics(8d, bold: false, narrow: true), new string('n', 40));

        Assert.True(narrow < wide, $"Narrow ({narrow:0.###}) must be under the default family ({wide:0.###}).");
    }

    /// <summary>
    /// The margin reaches what the host reported for real text, which is the whole
    /// point: under-measuring is the defect, because it clips.
    /// </summary>
    /// <remarks>
    /// The probe figures are ground truth: 287.03pt and 78.03pt at 8pt for 40
    /// characters. The residual is reported rather than asserted away, so a future
    /// change cannot quietly widen the gap.
    /// </remarks>
    [Theory]
    [InlineData('W', 287.02502)]
    [InlineData('i', 78.02504)]
    public void The_margin_reaches_what_the_host_reported_for_real_text(char character, double hostPt)
    {
        double measured = Width(new AptosTextMetrics(8d, bold: false, narrow: false), new string(character, 40));

        _output.WriteLine(
            $"40x'{character}': table+margin={measured:0.###} host={hostPt:0.###} residual={(measured - hostPt):0.####}");

        Assert.True(
            measured >= hostPt,
            $"40x'{character}' must measure at least what the host reported, or the label clips.");
    }

    /// <summary>
    /// The ellipsis measures, because it is the truncation marker: adding it to a
    /// string must make the string wider, never free.
    /// </summary>
    /// <remarks>
    /// Asserted through the seam rather than against the generated table, which is
    /// internal on purpose. Reaching past <see cref="ITextMetrics"/> would test the
    /// storage instead of the contract every caller depends on.
    /// </remarks>
    [Fact]
    public void Adding_the_ellipsis_widens_the_measurement()
    {
        var metrics = new AptosTextMetrics();

        double withoutMarker = Width(metrics, "x");
        double withMarker = Width(metrics, "x" + LabelText.Ellipsis);

        Assert.True(
            withMarker > withoutMarker,
            $"'{LabelText.Ellipsis}' must cost width, or a truncated label is narrower than the text it replaces ({withMarker:0.###} vs {withoutMarker:0.###}).");
    }

    /// <summary>
    /// A character with no glyph in the table still measures, and over-estimates
    /// rather than under-estimates.
    /// </summary>
    [Fact]
    public void An_unmapped_character_over_estimates_rather_than_clipping()
    {
        var metrics = new AptosTextMetrics(8d, bold: false, narrow: false);

        // A CJK ideograph is absent from a Latin-first face.
        double wide = Width(metrics, new string('W', 8));
        double unmapped = Width(metrics, new string('\u4E2D', 8));

        Assert.True(
            unmapped >= wide,
            $"An unmapped glyph must measure at least the widest ASCII advance ({unmapped:0.###} vs {wide:0.###}).");
    }

    /// <summary>A null string is refused rather than measured, matching the seam contract.</summary>
    [Fact]
    public void A_null_string_is_refused()
    {
        Assert.False(new AptosTextMetrics().TryMeasure(null!, out TextMeasurement? measurement));
        Assert.Null(measurement);
    }

    /// <summary>The empty string measures to zero width, which the truncation helper relies on.</summary>
    [Fact]
    public void The_empty_string_measures_zero_width()
    {
        Assert.Equal(0d, Width(new AptosTextMetrics(), string.Empty), 6);
    }

    /// <summary>
    /// A non-positive or non-finite font size is refused at construction.
    /// </summary>
    /// <remarks><b>Positive test for the validator</b> (AGENTS.md).</remarks>
    [Theory]
    [InlineData(0d)]
    [InlineData(-8d)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void A_non_positive_font_size_is_refused(double fontSizePt)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new AptosTextMetrics(fontSizePt, false, false));
    }

    /// <summary>
    /// The reported line height is proportional to the size, so a label box sized from
    /// it tracks the font rather than sitting at a fixed 10pt.
    /// </summary>
    [Fact]
    public void The_line_height_scales_with_the_font_size()
    {
        Assert.True(new AptosTextMetrics(8d, false, false).TryMeasure("x", out TextMeasurement? atEight));
        Assert.True(new AptosTextMetrics(16d, false, false).TryMeasure("x", out TextMeasurement? atSixteen));

        Assert.NotNull(atEight);
        Assert.NotNull(atSixteen);
        Assert.Equal(atEight!.HeightPt * 2d, atSixteen!.HeightPt, 3);
    }
}