namespace GanttCreator.AddIn;

/// <summary>
/// Which stroke palette a ribbon glyph is drawn in.
/// </summary>
internal enum RibbonGlyphPalette
{
    /// <summary>The near-black stroke, for the Colorful and White Office themes.</summary>
    Light,

    /// <summary>The near-white stroke, for the Dark Gray and Black Office themes.</summary>
    Dark,
}

/// <summary>
/// The one place the Office theme value becomes a palette.
/// </summary>
/// <remarks>
/// <para>
/// <b>The mapping is observed, never inferred.</b> <c>HKCU\Software\Microsoft\Office\{version}\Common</c>
/// <c>UI Theme</c> was read once per theme on 2026-11-16: Colorful <c>7</c>, White
/// <c>5</c>, Dark Gray <c>3</c>, Black <c>4</c>, Use system setting <c>6</c>. Microsoft
/// publishes the four theme names and the statement that only Black and Dark Gray are
/// the dark pair, but no numeric mapping, so those five values are the only authority
/// this type may use. See ADR-0063.
/// </para>
/// <para>
/// <b>High contrast is resolved before anything Office-specific.</b> Under a
/// high-contrast theme Office replaces its own chrome with system colours, so the window
/// colour is the only reliable signal; an Office-theme read would be answering a
/// question nobody asked.
/// </para>
/// <para>
/// <b>An unknown value falls through rather than guesses.</b> A value outside the
/// observed set, a key that cannot be read, or a Windows setting that is absent all
/// resolve through the system theme and finally to light, which is the palette every
/// other assumption in this add-in already makes.
/// </para>
/// </remarks>
internal static class RibbonGlyphTheme
{
    /// <summary>The <c>UI Theme</c> value observed for the Colorful theme.</summary>
    internal const int OfficeColorful = 7;

    /// <summary>The <c>UI Theme</c> value observed for the White theme.</summary>
    internal const int OfficeWhite = 5;

    /// <summary>The <c>UI Theme</c> value observed for the Dark Gray theme.</summary>
    internal const int OfficeDarkGray = 3;

    /// <summary>The <c>UI Theme</c> value observed for the Black theme.</summary>
    internal const int OfficeBlack = 4;

    /// <summary>The <c>UI Theme</c> value observed for "Use system setting".</summary>
    internal const int OfficeUseSystemTheme = 6;

    /// <summary>
    /// Resolves the palette from the environment facts.
    /// </summary>
    /// <param name="highContrast">Whether a high-contrast theme is active.</param>
    /// <param name="highContrastBackgroundIsLight">
    /// Whether the high-contrast window colour is light, which decides the stroke when
    /// <paramref name="highContrast"/> is true.
    /// </param>
    /// <param name="officeTheme">The observed <c>UI Theme</c> value, or null when unreadable.</param>
    /// <param name="systemAppsUseLight">
    /// The Windows app-theme flag, or null when the value is absent. True means light.
    /// </param>
    /// <returns>The palette to serve.</returns>
    internal static RibbonGlyphPalette Resolve(
        bool highContrast,
        bool highContrastBackgroundIsLight,
        int? officeTheme,
        bool? systemAppsUseLight)
    {
        if (highContrast)
        {
            return highContrastBackgroundIsLight
                ? RibbonGlyphPalette.Light
                : RibbonGlyphPalette.Dark;
        }

        switch (officeTheme)
        {
            case OfficeColorful:
            case OfficeWhite:
                return RibbonGlyphPalette.Light;

            case OfficeDarkGray:
            case OfficeBlack:
                return RibbonGlyphPalette.Dark;

            default:
                break;
        }

        // Use system setting, an unrecognised value, or an unreadable key: the Office
        // choice carries no answer of its own, so defer to the system theme.
        return systemAppsUseLight == false
            ? RibbonGlyphPalette.Dark
            : RibbonGlyphPalette.Light;
    }
}
