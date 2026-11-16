namespace GanttCreator.Office;

/// <summary>
/// Reads the user's automatic-refresh switch without being able to change it.
/// </summary>
/// <remarks>
/// A read-only view, so the scheduler can consult the switch without being able to flip it. The
/// mutable side is <see cref="AutoRefreshPreferences"/>.
/// </remarks>
public interface IAutoRefreshPreferenceSource
{
    /// <summary>
    /// Gets a value indicating whether a chart-affecting edit or a successful chart-setting
    /// command may schedule the coalesced automatic refresh.
    /// </summary>
    /// <remarks>
    /// <see langword="false"/> means only the explicit Refresh command rebuilds the chart. It never
    /// weakens a pass that does run: an enabled automatic pass and an explicit pass both read and
    /// validate the complete table before mutating a shape (ADR-0050 D6).
    /// </remarks>
    bool AutoRefreshEnabled { get; }
}

/// <summary>
/// The Excel session's view of the workbook's <c>AutoRefreshEnabled</c> setting (R5.20).
/// </summary>
/// <remarks>
/// <para>
/// <b>Persisted, not invented.</b> The value is the workbook's own
/// <see cref="Core.GanttAutoRefresh.SettingKey"/> row, written when the user presses the Refresh
/// menu's toggle and re-read when a workbook is opened or refreshed, so reopening a file restores
/// the choice the user made in it. The default is <see langword="true"/>, which is the behaviour
/// every release shipped and what an absent row means.
/// </para>
/// <para>
/// <b>A cached copy, because the scheduler is asked on every cell edit.</b> The scheduler is on a
/// hot path — Excel raises `SheetChange` for every edit — so it reads this one boolean rather than
/// reading the hidden configuration sheet per event. The cache is refreshed by the same activation
/// read that seeds the rest of the ribbon state, which makes the workbook the single authority and
/// this object its snapshot.
/// </para>
/// <para>
/// Mutable and explicitly shared, like <see cref="ProtectionPreferences"/>: it lives as long as the
/// Excel session, is touched only on the main thread, and is re-seeded from whichever workbook is
/// active. It deliberately has no <c>Reset()</c> on a workbook switch — the re-seed is the reset,
/// and dropping to the default would switch the feature back on in a workbook that had it off.
/// </para>
/// </remarks>
public sealed class AutoRefreshPreferences : IAutoRefreshPreferenceSource
{
    /// <summary>The Excel session's single switch.</summary>
    public static AutoRefreshPreferences Shared { get; } = new();

    /// <summary>
    /// Gets or sets a value indicating whether the automatic refresh is enabled. Defaults to
    /// <see langword="true"/>, the value every workbook written before the setting existed has.
    /// </summary>
    public bool AutoRefreshEnabled { get; set; } = true;

    /// <summary>Sets the switch from the workbook's stored text.</summary>
    /// <param name="stored">The <c>AutoRefreshEnabled</c> cell text, or null when absent.</param>
    /// <remarks>
    /// The parse belongs to <see cref="Core.GanttAutoRefresh.Parse"/>, so the reader in the state
    /// service and this one cannot disagree about what an absent or unreadable row means.
    /// </remarks>
    public void ApplyStored(string? stored) => AutoRefreshEnabled = Core.GanttAutoRefresh.Parse(stored);
}
