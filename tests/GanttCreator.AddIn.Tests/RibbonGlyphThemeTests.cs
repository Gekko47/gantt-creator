namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The palette resolution chain, asserted against the values that were actually observed
/// on 2026-11-16 rather than against a guess about what Office writes.
/// </summary>
public sealed class RibbonGlyphThemeTests
{
    [Theory]
    [InlineData(RibbonGlyphTheme.OfficeDarkGray)]
    [InlineData(RibbonGlyphTheme.OfficeBlack)]
    public void A_dark_office_theme_resolves_to_the_dark_palette(int officeTheme) =>
        // The system says light on purpose: the Office choice must win.
        Assert.Equal(RibbonGlyphPalette.Dark, RibbonGlyphTheme.Resolve(false, true, officeTheme, true));

    [Theory]
    [InlineData(RibbonGlyphTheme.OfficeColorful)]
    [InlineData(RibbonGlyphTheme.OfficeWhite)]
    public void A_light_office_theme_resolves_to_the_light_palette(int officeTheme) =>
        Assert.Equal(RibbonGlyphPalette.Light, RibbonGlyphTheme.Resolve(false, false, officeTheme, false));

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void Use_system_setting_defers_to_the_windows_theme(bool systemIsLight, bool expectDark)
    {
        // The expectation is a bool because a public test signature cannot expose the
        // internal palette enum.
        RibbonGlyphPalette expected = expectDark
            ? RibbonGlyphPalette.Dark
            : RibbonGlyphPalette.Light;

        Assert.Equal(
            expected,
            RibbonGlyphTheme.Resolve(false, true, RibbonGlyphTheme.OfficeUseSystemTheme, systemIsLight));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(42)]
    [InlineData(-1)]
    public void An_unrecognised_office_value_falls_through_to_the_windows_theme(int officeTheme) =>
        Assert.Equal(RibbonGlyphPalette.Dark, RibbonGlyphTheme.Resolve(false, true, officeTheme, false));

    [Fact]
    public void An_unrecognised_office_value_with_a_light_windows_theme_resolves_to_light() =>
        Assert.Equal(RibbonGlyphPalette.Light, RibbonGlyphTheme.Resolve(false, false, 42, true));

    [Fact]
    public void An_unreadable_probe_and_no_windows_value_resolve_to_light() =>
        Assert.Equal(RibbonGlyphPalette.Light, RibbonGlyphTheme.Resolve(false, false, null, null));

    [Fact]
    public void An_unreadable_office_probe_still_honours_a_dark_windows_theme() =>
        Assert.Equal(RibbonGlyphPalette.Dark, RibbonGlyphTheme.Resolve(false, true, null, false));

    [Fact]
    public void High_contrast_overrides_an_explicit_dark_office_theme() =>
        Assert.Equal(
            RibbonGlyphPalette.Light,
            RibbonGlyphTheme.Resolve(true, true, RibbonGlyphTheme.OfficeBlack, false));

    [Fact]
    public void High_contrast_with_a_dark_window_overrides_an_explicit_light_office_theme() =>
        Assert.Equal(
            RibbonGlyphPalette.Dark,
            RibbonGlyphTheme.Resolve(true, false, RibbonGlyphTheme.OfficeColorful, true));

    [Fact]
    public void The_environment_probe_always_answers_with_a_defined_palette() =>
        Assert.True(Enum.IsDefined(RibbonGlyphEnvironment.Current));

    [Fact]
    public void The_office_theme_probe_is_total()
    {
        // The probe must be total: whatever this machine holds, the call returns.
        int? value = RibbonGlyphEnvironment.ReadOfficeTheme();
        Assert.True(value is null or >= 0);
    }

    [Fact]
    public void The_windows_theme_probe_is_repeatable()
    {
        // A three-state answer is inherently total; what matters is that two reads agree,
        // so a flapping probe cannot make the ribbon change palette on its own.
        bool? first = RibbonGlyphEnvironment.ReadSystemAppsUseLight();
        bool? second = RibbonGlyphEnvironment.ReadSystemAppsUseLight();
        Assert.Equal(first, second);
    }
}
