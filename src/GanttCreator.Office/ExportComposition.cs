using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// The shared export-composition rule: which composition an export covers.
/// </summary>
/// <remarks>
/// <para>
/// <b>One authority for the selection.</b> The exact same rule
/// (<see cref="SelectExportSelection"/>) decides the export coverage for Copy
/// Editable, Send to PowerPoint and the Export PNG command. It moved here from the
/// AddIn pre-flight so the Office-side composition and the AddIn-side commands cannot
/// disagree about which composition they describe.
/// </para>
/// <para>
/// The per-refresh export-aspect record that once lived beside this rule (R8.7 D4)
/// was withdrawn with the R8.7 width/preview surface (owner ruling 2026-11-23). The
/// Export PNG command sizes itself from its own pre-flight's bounds, so no cached
/// aspect exists to go stale.
/// </para>
/// </remarks>
public static class ExportComposition
{
    /// <summary>
    /// The settings key that selects the export composition (catalogue default
    /// <c>TRUE</c>).
    /// </summary>
    public const string ExportIncludeDataPanelSettingKey = "ExportIncludeDataPanel";

    /// <summary>Maps the stored inclusion setting to the export selection.</summary>
    /// <param name="settings">The workbook's effective settings map.</param>
    /// <returns>Chart-only on an explicit case-insensitive <c>FALSE</c>; otherwise table-plus-chart.</returns>
    /// <remarks>
    /// Fail-open toward the catalogue default (<c>TRUE</c>), matching the
    /// automatic-refresh parse precedent: a typo cannot silently change what an
    /// export covers, and the catalogue remains the key's authority.
    /// </remarks>
    public static ExportSelection SelectExportSelection(IReadOnlyDictionary<string, string> settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return settings.TryGetValue(ExportIncludeDataPanelSettingKey, out string? stored)
            && string.Equals(stored?.Trim(), "FALSE", StringComparison.OrdinalIgnoreCase)
                ? ExportSelection.ChartOnly
                : ExportSelection.TablePlusChart;
    }
}
