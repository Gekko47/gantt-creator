using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// The Excel session's view of the workbook's PNG-export settings (schema version 28, R8.8).
/// </summary>
/// <remarks>
/// <para>
/// <b>Persisted, not invented.</b> <see cref="Target"/> is the workbook's own
/// <see cref="GanttPngExport.TargetSettingKey"/> row, written when the user toggles the
/// Export PNG split button's menu and re-read when a workbook is opened or refreshed, so
/// reopening a file restores the choice the user made in it. <see cref="Folder"/> is the
/// <see cref="GanttPngExport.FolderSettingKey"/> row, written after a successful file
/// export.
/// </para>
/// <para>
/// <b>A cached copy.</b> The ribbon's pressed-state getters and the command's target read run
/// on every ribbon invalidation and on every click, so they read this snapshot rather than the
/// hidden configuration sheet. The cache is refreshed by the same activation read that seeds
/// the automatic-refresh switch, which makes the workbook the single authority.
/// </para>
/// <para>
/// Mutable and explicitly shared, like <see cref="AutoRefreshPreferences"/>: it lives as long as
/// the Excel session, is touched only on the main thread, and is re-seeded from whichever
/// workbook is active.
/// </para>
/// </remarks>
public sealed class PngExportPreferences
{
    /// <summary>The Excel session's single PNG-export preference set.</summary>
    public static PngExportPreferences Shared { get; } = new();

    /// <summary>
    /// Gets or sets where the Export PNG main action delivers the image. Defaults to the
    /// clipboard, the value every workbook written before the setting existed has.
    /// </summary>
    public PngExportTarget Target { get; set; } = PngExportTarget.Clipboard;

    /// <summary>
    /// Gets or sets the folder the last successful file export wrote to, or an empty string
    /// when none is remembered. Never assumed to exist: the consumer falls back to the
    /// workbook's folder when the remembered path is absent.
    /// </summary>
    public string Folder { get; set; } = string.Empty;

    /// <summary>Adopts the workbook's stored target and remembered folder.</summary>
    /// <param name="settings">The settings map the activation read produced.</param>
    /// <remarks>
    /// The parse belongs to <see cref="GanttPngExport.Parse"/>, so an absent or
    /// unreadable row means the clipboard, exactly as it did before the keys existed.
    /// </remarks>
    public void ApplyStored(IReadOnlyDictionary<string, string> settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        Target = GanttPngExport.Parse(
            settings.TryGetValue(GanttPngExport.TargetSettingKey, out string? storedTarget)
                ? storedTarget
                : null);
        Folder = settings.TryGetValue(GanttPngExport.FolderSettingKey, out string? storedFolder)
            ? storedFolder?.Trim() ?? string.Empty
            : string.Empty;
    }
}
