namespace GanttCreator.Core;

/// <summary>
/// The automatic-refresh switch stored in <c>tblGanttSettings</c> (schema version 25, R5.20).
/// </summary>
/// <remarks>
/// <para>
/// <b>One key, because the switch is one fact.</b> <see cref="SettingKey"/> records whether a
/// chart-affecting table edit or a successful chart-setting command may rebuild the chart
/// automatically. It is the user's switch, persisted with the workbook, so the choice survives
/// opening the file in another instance — the same reason <see cref="GanttColumnVisibility.SettingKey"/>
/// is a key rather than session state.
/// </para>
/// <para>
/// <b>Absent and unreadable both mean enabled.</b> Every workbook written before this key existed
/// behaves exactly as it did when it was written, which is the value this class returns for a blank
/// cell. A blank value is also the catalogue's default, so a freshly initialised workbook is
/// enabled without the writer having to invent a row. The predicate matches
/// <c>ProtectGanttSheet</c>'s reading in <c>GanttRefreshOrchestrator</c>: only an explicit
/// <see cref="DisabledValue"/> turns the feature off, so a typo in the hidden sheet cannot silently
/// stop every automatic refresh.
/// </para>
/// <para>
/// <b>This key is not an invariant relaxation.</b> The explicit Refresh command remains a full
/// validation and repair pass whatever the switch says, and the automatic pass, when enabled,
/// still reads and validates the complete table before any shape mutation. The switch only decides
/// whether the coalesced pass is scheduled at all (ADR-0050 D6).
/// </para>
/// </remarks>
public static class GanttAutoRefresh
{
    /// <summary>The settings key that stores the automatic-refresh switch (schema version 25).</summary>
    public const string SettingKey = "AutoRefreshEnabled";

    /// <summary>The stored value meaning automatic refresh is scheduled.</summary>
    public const string EnabledValue = "TRUE";

    /// <summary>The stored value meaning only an explicit Refresh rebuilds the chart.</summary>
    public const string DisabledValue = "FALSE";

    /// <summary>Returns the stored text for a switch state.</summary>
    /// <param name="enabled">Whether automatic refresh is enabled.</param>
    /// <returns>The single stored form, so two workbooks' text differs only when the choice does.</returns>
    public static string Format(bool enabled) => enabled ? EnabledValue : DisabledValue;

    /// <summary>Reads the stored switch.</summary>
    /// <param name="stored">The stored cell text, or null/blank when the key is absent.</param>
    /// <returns>
    /// <see langword="true"/> unless the value is explicitly <see cref="DisabledValue"/>.
    /// </returns>
    /// <remarks>
    /// Deliberately fail-open: a workbook that never wrote the key, and a value that is not the
    /// disabled token, both keep the behaviour every release has shipped. Reading an unreadable
    /// value as "off" would stop the automatic refresh on a workbook nobody switched off, which is
    /// the one failure a user could not see or undo from the ribbon.
    /// </remarks>
    public static bool Parse(string? stored) =>
        !string.Equals(stored?.Trim(), DisabledValue, StringComparison.OrdinalIgnoreCase);
}
