using System.Security;
using Microsoft.Win32;

namespace GanttCreator.AddIn;

/// <summary>
/// Reads the three environment facts the glyph palette depends on, and never throws.
/// </summary>
/// <remarks>
/// <para>
/// <b>One registry value, one Windows value, two system colours.</b> The Office UI theme
/// is the only fact that differs from what Windows reports, which is exactly why the
/// probe exists: a user can run Windows Light with Office Black, and that user is the
/// reason the old fixed-black glyphs disappeared.
/// </para>
/// <para>
/// <b>No COM.</b> The Office key is version-scoped, and the version segment is tried
/// against the known Office installs in order rather than read from
/// <c>Application.Version</c>, so answering an image request never touches Excel's object
/// model. <c>16.0</c> is the verified present case; <c>15.0</c> and <c>14.0</c> are tried
/// only if it is absent.
/// </para>
/// <para>
/// <b>A refusal is a fact, not an error.</b> A missing key, a denied read, or a value of
/// an unexpected type all return null so the resolver can fall through its documented
/// chain. A ribbon callback that throws takes the whole ribbon down, so nothing here is
/// allowed to propagate.
/// </para>
/// </remarks>
internal static class RibbonGlyphEnvironment
{
    private const string _officeCommonParent = @"Software\Microsoft\Office";
    private const string _officeThemeValueName = "UI Theme";
    private const string _personaliseKey = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
    private const string _appsUseLightThemeValueName = "AppsUseLightTheme";

    /// <summary>
    /// The Office version keys to try, most likely first. <c>16.0</c> was verified present
    /// on the development host; the older keys cover an install that predates it.
    /// </summary>
    private static readonly string[] _officeVersionCandidates = ["16.0", "15.0", "14.0"];

    /// <summary>The luminance midpoint between a light and a dark window colour.</summary>
    private const double _luminanceMidpoint = 127.5;

    /// <summary>
    /// Gets the palette for the current environment.
    /// </summary>
    internal static RibbonGlyphPalette Current =>
        RibbonGlyphTheme.Resolve(
            IsHighContrast(),
            IsHighContrastBackgroundLight(),
            ReadOfficeTheme(),
            ReadSystemAppsUseLight());

    /// <summary>
    /// Reads the Office <c>UI Theme</c> value from the first version key that holds one.
    /// </summary>
    /// <returns>The observed value, or null when no key could be read.</returns>
    internal static int? ReadOfficeTheme()
    {
        foreach (string version in _officeVersionCandidates)
        {
            object? value = ReadRegistryValue($@"{_officeCommonParent}\{version}\Common", _officeThemeValueName);
            if (value is int theme)
            {
                return theme;
            }
        }

        return null;
    }

    /// <summary>
    /// Reads the Windows app-theme flag.
    /// </summary>
    /// <returns>True for light, false for dark, or null when the value is absent.</returns>
    internal static bool? ReadSystemAppsUseLight()
    {
        object? value = ReadRegistryValue(_personaliseKey, _appsUseLightThemeValueName);
        return value is int flag ? flag != 0 : null;
    }

    /// <summary>Reports whether a high-contrast theme is active.</summary>
    /// <returns>True when a high-contrast theme is in use.</returns>
    internal static bool IsHighContrast()
    {
        try
        {
            return SystemInformation.HighContrast;
        }
        catch (InvalidOperationException)
        {
            // No window station: treat as an ordinary theme rather than failing the image.
            return false;
        }
    }

    /// <summary>
    /// Reports whether the high-contrast window background is light.
    /// </summary>
    /// <returns>True when the window colour is lighter than the midpoint.</returns>
    internal static bool IsHighContrastBackgroundLight()
    {
        Color window;
        try
        {
            window = SystemColors.Window;
        }
        catch (InvalidOperationException)
        {
            return true;
        }

        double luminance = (0.2126 * window.R) + (0.7152 * window.G) + (0.0722 * window.B);
        return luminance > _luminanceMidpoint;
    }

    /// <summary>Reads one value from <c>HKEY_CURRENT_USER</c>, returning null on any refusal.</summary>
    /// <param name="subKey">The key path below <c>HKEY_CURRENT_USER</c>.</param>
    /// <param name="name">The value name.</param>
    /// <returns>The value, or null when it cannot be read.</returns>
    private static object? ReadRegistryValue(string subKey, string name)
    {
        try
        {
            return Registry.GetValue(@"HKEY_CURRENT_USER\" + subKey, name, null);
        }
        catch (SecurityException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
        catch (IOException)
        {
            return null;
        }
        catch (ArgumentException)
        {
            return null;
        }
    }
}
