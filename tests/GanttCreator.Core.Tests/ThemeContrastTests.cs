using GanttCreator.Core;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Pins the R5.6 contrast policy: WCAG G18 4.5:1 luminance gate plus a
/// non-blocking red-green hue advisory, with shipped defaults grandfathered.
/// </summary>
public sealed class ThemeContrastTests
{
    [Fact]
    public void Black_on_white_exceeds_the_minimum_ratio()
    {
        var ratio = ThemeContrast.Ratio(ColourHex.Parse("#000000"), ColourHex.Parse("#FFFFFF"));
        Assert.Equal(21.0, ratio, 1);
        Assert.True(ThemeContrast.MeetsMinimum(ColourHex.Parse("#000000"), ColourHex.Parse("#FFFFFF")));
    }

    [Fact]
    public void White_on_delay_red_fails_the_minimum_ratio()
    {
        // DelayText #FFFFFF on DelayFill #FF0000 ≈ 4.0:1 — below 4.5 by design,
        // grandfathered for shipped defaults; a user edit creating this pair warns.
        Assert.False(ThemeContrast.MeetsMinimum(ColourHex.Parse("#FFFFFF"), ColourHex.Parse("#FF0000")));
    }

    [Fact]
    public void Shipped_green_blue_red_defaults_are_unchanged_inputs_not_warnings()
    {
        // Grandfather pin: these are the approved catalogue values, not warnings.
        Assert.Equal("#92D050", GanttCatalogues.GetPreset("AsPlannedActivity").FillColour);
        Assert.Equal("#00B0F0", GanttCatalogues.GetPreset("AsBuiltActivity").FillColour);
        Assert.Equal("#FF0000", GanttCatalogues.GetPreset("DelayEvent").FillColour);
        Assert.Equal("#FF0000", GanttCatalogues.GetPreset("CriticalInterval").StrokeColour);
    }

    [Fact]
    public void Red_green_pair_is_advisory_even_when_luminance_passes()
    {
        Assert.True(ThemeContrast.IsRedGreenPair(ColourHex.Parse("#FF0000"), ColourHex.Parse("#00B000")));
        Assert.False(ThemeContrast.IsRedGreenPair(ColourHex.Parse("#000000"), ColourHex.Parse("#FFFFFF")));
    }

    [Fact]
    public void Null_colours_throw_before_any_ratio_math()
    {
        Assert.Throws<ArgumentNullException>(() => ThemeContrast.Ratio(null!, ColourHex.Parse("#FFFFFF")));
        Assert.Throws<ArgumentNullException>(() => ThemeContrast.Ratio(ColourHex.Parse("#000000"), null!));
        Assert.Throws<ArgumentNullException>(() => ThemeContrast.IsRedGreenPair(null!, ColourHex.Parse("#FFFFFF")));
    }
}
