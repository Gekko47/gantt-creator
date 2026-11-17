namespace GanttCreator.Core;

/// <summary>
/// The chart-chrome colour roles: the parts of the chart frame that carry a
/// colour of their own, as opposed to an entity's style colours.
/// </summary>
/// <remarks>
/// <para>
/// Each role is backed by a code-owned colour token from the entity guide's
/// "Shared colour and typography tokens" table, and is stored per workbook as
/// a setting, so the chart's background, alternating bands, grid, and header
/// bands are editable without a code change.
/// </para>
/// <para>
/// The roles are deliberately <em>chart</em>-scoped. An entity's fill, stroke,
/// and label colours belong to a named style and are edited on the style
/// (see <see cref="GanttStyleRegistry"/>); a chrome role is the frame the
/// entities sit on, and it belongs to no entity at all.
/// </para>
/// </remarks>
public enum GanttChartChromeRole
{
    /// <summary>The chart background behind the plot and the header bands.</summary>
    Background = 0,

    /// <summary>The band drawn behind alternate plot periods.</summary>
    AlternateBand = 1,

    /// <summary>The minor grid lines and lane separators.</summary>
    MinorGrid = 2,

    /// <summary>The major grid lines and the plot frame boundary.</summary>
    MajorGrid = 3,

    /// <summary>The year header band.</summary>
    YearHeader = 4,

    /// <summary>The period header band.</summary>
    PeriodHeader = 5,
}

/// <summary>
/// The mapping between a chart-chrome colour role, the code-owned colour token
/// that names it, and the setting key it is stored under.
/// </summary>
/// <remarks>
/// <para>
/// <b>One authority per value.</b> The token names the approved colour and the
/// entity guide publishes its value; the setting key is where the workbook
/// stores the user's choice. The catalogue's setting default and the token
/// value are pinned equal by a test, so the two cannot drift apart silently.
/// </para>
/// <para>
/// The period header band has no token of its own: the entity guide's section 6
/// says its "alternating neutral fills may follow plot bands", and the
/// reference scene gives it the band fill. It therefore borrows
/// <c>AlternateBandFill</c> while keeping its own setting key, so a user can
/// still give the period band a colour of its own.
/// </para>
/// </remarks>
public static class GanttChartChrome
{
    /// <summary>
    /// The chrome roles, in the order the dialog presents them.
    /// </summary>
    public static IReadOnlyList<GanttChartChromeRole> Roles { get; } =
    [
        GanttChartChromeRole.Background,
        GanttChartChromeRole.AlternateBand,
        GanttChartChromeRole.MinorGrid,
        GanttChartChromeRole.MajorGrid,
        GanttChartChromeRole.YearHeader,
        GanttChartChromeRole.PeriodHeader,
    ];

    /// <summary>
    /// The colour token that names a chrome role.
    /// </summary>
    /// <param name="role">The chrome role.</param>
    /// <returns>The catalogue colour token name.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the role is not defined.</exception>
    public static string TokenNameFor(GanttChartChromeRole role) =>
        role switch
        {
            GanttChartChromeRole.Background => "ChartBackground",
            GanttChartChromeRole.AlternateBand => "AlternateBandFill",
            GanttChartChromeRole.MinorGrid => "MinorGridStroke",
            GanttChartChromeRole.MajorGrid => "MajorGridStroke",
            GanttChartChromeRole.YearHeader => "YearHeaderFill",
            GanttChartChromeRole.PeriodHeader => "AlternateBandFill",
            _ => throw new ArgumentOutOfRangeException(nameof(role), role, "Unknown chart chrome role."),
        };

    /// <summary>
    /// The setting key a chrome role is stored under.
    /// </summary>
    /// <param name="role">The chrome role.</param>
    /// <returns>The setting key.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the role is not defined.</exception>
    public static string SettingKeyFor(GanttChartChromeRole role) =>
        role switch
        {
            GanttChartChromeRole.Background => "ChartBackgroundColour",
            GanttChartChromeRole.AlternateBand => "AlternateBandColour",
            GanttChartChromeRole.MinorGrid => "MinorGridColour",
            GanttChartChromeRole.MajorGrid => "MajorGridColour",
            GanttChartChromeRole.YearHeader => "YearHeaderColour",
            GanttChartChromeRole.PeriodHeader => "PeriodHeaderColour",
            _ => throw new ArgumentOutOfRangeException(nameof(role), role, "Unknown chart chrome role."),
        };

    /// <summary>
    /// The catalogue default for a chrome role: the approved token value.
    /// </summary>
    /// <param name="role">The chrome role.</param>
    /// <returns>The uppercase <c>#RRGGBB</c> default.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the role is not defined.</exception>
    public static string DefaultHexFor(GanttChartChromeRole role)
    {
        string tokenName = TokenNameFor(role);
        foreach (GanttColourToken token in GanttCatalogues.Colours)
        {
            if (string.Equals(token.Name, tokenName, StringComparison.Ordinal))
            {
                return token.HexValue;
            }
        }

        throw new ArgumentOutOfRangeException(nameof(role), role, "The chrome role names no catalogue colour token.");
    }

    /// <summary>
    /// Determines whether a setting key names a chart-chrome colour.
    /// </summary>
    /// <param name="settingKey">The setting key.</param>
    /// <returns>True when the key stores a chrome colour.</returns>
    public static bool IsColourSettingKey(string settingKey)
    {
        ArgumentNullException.ThrowIfNull(settingKey);
        foreach (GanttChartChromeRole role in Roles)
        {
            if (string.Equals(SettingKeyFor(role), settingKey, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// The user-facing name of a chrome role.
    /// </summary>
    /// <param name="role">The chrome role.</param>
    /// <returns>The display name.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the role is not defined.</exception>
    public static string DisplayLabelFor(GanttChartChromeRole role) =>
        role switch
        {
            GanttChartChromeRole.Background => "Background",
            GanttChartChromeRole.AlternateBand => "Alternate band",
            GanttChartChromeRole.MinorGrid => "Minor grid",
            GanttChartChromeRole.MajorGrid => "Major grid",
            GanttChartChromeRole.YearHeader => "Year header",
            GanttChartChromeRole.PeriodHeader => "Period header",
            _ => throw new ArgumentOutOfRangeException(nameof(role), role, "Unknown chart chrome role."),
        };
}
