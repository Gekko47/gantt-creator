namespace GanttCreator.Core.Tests;

/// <summary>
/// R5.2: the plot margin is a closed preset set (Windows defaults) plus a custom
/// value, resolved once to points, and the catalogue default matches the code
/// default so an unset margin renders identically to a fresh one.
/// </summary>
public sealed class PlotMarginsTests
{
    private const double TolerancePt = 1e-9;

    [Theory]
    [InlineData(GanttPlotMargin.Narrow, 0.635)]
    [InlineData(GanttPlotMargin.Normal, 1.78)]
    [InlineData(GanttPlotMargin.Wide, 2.54)]
    public void Each_preset_has_its_windows_default_centimetre_value(GanttPlotMargin margin, double expectedCm)
    {
        Assert.Equal(expectedCm, GanttPlotMargins.PresetCm(margin), TolerancePt);
    }

    [Fact]
    public void Custom_has_no_fixed_preset_value_and_throws()
    {
        // A custom margin carries no fixed value; it must be resolved through the
        // MarginCm companion, so asking for its "preset" value is a programming error.
        Assert.Throws<ArgumentOutOfRangeException>(() => GanttPlotMargins.PresetCm(GanttPlotMargin.Custom));
    }

    [Theory]
    [InlineData("Narrow", GanttPlotMargin.Narrow)]
    [InlineData("Normal", GanttPlotMargin.Normal)]
    [InlineData("Wide", GanttPlotMargin.Wide)]
    [InlineData("Custom", GanttPlotMargin.Custom)]
    public void The_parser_accepts_each_exact_member_name(string text, GanttPlotMargin expected)
    {
        Assert.True(GanttPlotMargins.TryParse(text, out var margin));
        Assert.Equal(expected, margin);
    }

    [Theory]
    [InlineData("narrow")]
    [InlineData("Normal ")]
    [InlineData("Margin")]
    [InlineData("")]
    [InlineData(null)]
    public void The_parser_refuses_unknown_or_malformed_text(string? text)
    {
        // Positive test for the closed set: a stored value that is not the exact
        // member name is refused rather than coerced to a default.
        Assert.False(GanttPlotMargins.TryParse(text, out _));
    }

    [Fact]
    public void A_preset_margin_resolves_to_its_centimetre_value_in_points()
    {
        // Narrow = 0.635cm = 6.35mm; points = mm * (72/25.4).
        double expected = 0.635 * 10.0 * SizePresets.PointsPerMillimetre;

        Assert.Equal(expected, GanttPlotMargins.MarginPt(GanttPlotMargin.Narrow, 99), TolerancePt);
    }

    [Fact]
    public void A_custom_margin_resolves_to_its_centimetre_value_in_points()
    {
        double expected = 2.5 * 10.0 * SizePresets.PointsPerMillimetre;

        Assert.Equal(expected, GanttPlotMargins.MarginPt(GanttPlotMargin.Custom, 2.5), TolerancePt);
    }

    [Theory]
    [InlineData(-5.0)]
    [InlineData(999.0)]
    [InlineData(double.NaN)]
    public void An_out_of_range_custom_margin_is_clamped_rather_than_refused(double customCm)
    {
        // A stale or corrupt stored custom value must not break a Refresh or produce a
        // negative/unbounded margin, so it is clamped into the permitted band. The NaN
        // case is caught by Math.Clamp's ordering (it returns the minimum for NaN input).
        double resolved = GanttPlotMargins.MarginPt(GanttPlotMargin.Custom, customCm);

        Assert.True(double.IsFinite(resolved));
        Assert.True(resolved >= GanttPlotMargins.MinimumCustomCm * 10.0 * SizePresets.PointsPerMillimetre - TolerancePt);
        Assert.True(resolved <= GanttPlotMargins.MaximumCustomCm * 10.0 * SizePresets.PointsPerMillimetre + TolerancePt);
    }

    [Fact]
    public void The_default_margin_is_normal_and_matches_the_catalogue_default()
    {
        // The two must agree: a workbook whose setting is absent must render exactly
        // like one whose setting carries the default (the same rule SizePresets.Default
        // established for the size preset).
        Assert.Equal(GanttPlotMargin.Normal, GanttPlotMargins.Default);

        GanttSettingDefinition stored = GanttCatalogues.Settings
            .First(setting => setting.Key == "Margin");
        Assert.True(GanttPlotMargins.TryParse(stored.DefaultValue, out var margin));
        Assert.Equal(GanttPlotMargins.Default, margin);
    }

    [Fact]
    public void The_catalogue_carries_both_margin_keys_with_a_valid_default_custom_value()
    {
        GanttSettingDefinition marginCm = GanttCatalogues.Settings
            .First(setting => setting.Key == "MarginCm");

        Assert.True(
            double.TryParse(marginCm.DefaultValue, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cm));
        Assert.True(cm >= GanttPlotMargins.MinimumCustomCm && cm <= GanttPlotMargins.MaximumCustomCm);
    }
}
