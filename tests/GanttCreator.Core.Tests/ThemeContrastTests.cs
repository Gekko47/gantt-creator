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

    /// <summary>
    /// A pair the user did NOT change is grandfathered: it is not reported.
    /// This is R5.6 D2's "no warning on open, warning only on a user edit
    /// that creates a violation".
    /// </summary>
    /// <remarks>
    /// The shipped Delay Event is white on red at about 4.0:1, which is below
    /// the 4.5:1 gate. Without the baseline the dialog would warn about it
    /// every time it was opened and accepted, because the dialog emits an edit
    /// for every style it holds rather than only the ones the user touched.
    /// </remarks>
    [Fact]
    public void An_unchanged_low_contrast_pair_is_not_reported()
    {
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("DelayEvent", "#FF0000", "#C00000", "#FFFFFF")],
            new Dictionary<string, double>());

        var baseline = new List<ThemeStyleEdit>
        {
            new("DelayEvent", "#FF0000", "#C00000", "#FFFFFF"),
        };

        Assert.Empty(ThemeEditRequest.ContrastWarnings(request, baseline));
    }

    /// <summary>
    /// A user edit that CREATES a low-contrast pair is reported. This is the
    /// positive control for the grandfathering: without it, "no warning" would
    /// be indistinguishable from "the check never fires".
    /// </summary>
    [Fact]
    public void A_user_edit_that_creates_a_low_contrast_pair_is_reported()
    {
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("DelayEvent", "#FF0000", "#C00000", "#FFFFFF")],
            new Dictionary<string, double>());

        // The user changed the TEXT colour to black, which passes contrast, so
        // the pair no longer violates and must not be reported.
        var passing = new ThemeEditRequest(
            [new ThemeStyleEdit("DelayEvent", "#FF0000", "#C00000", "#000000")],
            new Dictionary<string, double>());
        var passingBaseline = new List<ThemeStyleEdit>
        {
            new("DelayEvent", "#FF0000", "#C00000", "#FFFFFF"),
        };
        Assert.Empty(ThemeEditRequest.ContrastWarnings(passing, passingBaseline));

        // The user changed the FILL to white, leaving white text on white:
        // a violation the user created, so it is reported.
        var violating = new ThemeEditRequest(
            [new ThemeStyleEdit("DelayEvent", "#FFFFFF", "#C00000", "#FFFFFF")],
            new Dictionary<string, double>());
        var violatingBaseline = new List<ThemeStyleEdit>
        {
            new("DelayEvent", "#FF0000", "#C00000", "#FFFFFF"),
        };
        Assert.NotEmpty(ThemeEditRequest.ContrastWarnings(violating, violatingBaseline));
    }

    /// <summary>
    /// A style the baseline does not carry is treated as changed, and is
    /// reported. The conservative reading: the caller cannot know whether the
    /// user touched it, so it is not grandfathered.
    /// </summary>
    [Fact]
    public void A_style_absent_from_the_baseline_is_reported()
    {
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("DelayEvent", "#FF0000", "#C00000", "#FFFFFF")],
            new Dictionary<string, double>());

        var warnings = ThemeEditRequest.ContrastWarnings(request, []);

        Assert.NotEmpty(warnings);
    }
}
