namespace GanttCreator.Core;

/// <summary>
/// Where the Export PNG command delivers the image (schema version 28, R8.8).
/// </summary>
public enum PngExportTarget
{
    /// <summary>The rendered image is copied to the clipboard (the default).</summary>
    Clipboard = 0,

    /// <summary>The rendered image is written to a verified file.</summary>
    File = 1,
}

/// <summary>
/// The PNG-export preferences stored in <c>tblGanttSettings</c> (schema version 28, R8.8).
/// </summary>
/// <remarks>
/// <para>
/// <b>Two keys, because the export remembers two facts.</b>
/// <see cref="TargetSettingKey"/> records whether the split button's main action copies to
/// the clipboard or writes a file, and <see cref="FolderSettingKey"/> records the folder
/// the last successful file write used, so the save dialog reopens there. Both are the
/// user's choices persisted with the workbook, for the same reason the automatic-refresh
/// switch and the hidden-column list are: the choice survives opening the file in another
/// instance.
/// </para>
/// <para>
/// <b>Absent and unreadable both mean the safe default.</b> <see cref="Parse"/> fails open
/// to <see cref="PngExportTarget.Clipboard"/>, so a workbook written before these keys
/// existed behaves exactly as a fresh one, and a typo in the hidden sheet cannot silently
/// change where an export goes. The settings reader still validates the stored target
/// against the closed two-value set and refuses anything else, so the fail-open parse is
/// only the consumer's last line of defence, mirroring <see cref="GanttAutoRefresh"/>.
/// </para>
/// </remarks>
public static class GanttPngExport
{
    /// <summary>The settings key that stores the export target (schema version 28).</summary>
    public const string TargetSettingKey = "PngExportTarget";

    /// <summary>The settings key that stores the last successful file-export folder (schema version 28).</summary>
    public const string FolderSettingKey = "LastPngExportFolder";

    /// <summary>The stored value meaning the image is copied to the clipboard.</summary>
    public const string ClipboardValue = "clipboard";

    /// <summary>The stored value meaning the image is written to a file.</summary>
    public const string FileValue = "file";

    /// <summary>Returns the stored text for a target.</summary>
    /// <param name="target">The export target.</param>
    /// <returns>The single stored form, so two workbooks' text differs only when the choice does.</returns>
    public static string Format(PngExportTarget target) =>
        target == PngExportTarget.File ? FileValue : ClipboardValue;

    /// <summary>Reads the stored target, failing open to the clipboard.</summary>
    /// <param name="stored">The stored cell text, or null/blank when the key is absent.</param>
    /// <returns>The stored target, or <see cref="PngExportTarget.Clipboard"/> when unreadable.</returns>
    public static PngExportTarget Parse(string? stored) =>
        TryParse(stored, out PngExportTarget target) ? target : PngExportTarget.Clipboard;

    /// <summary>Strictly parses the stored target, for the settings reader's validation.</summary>
    /// <param name="stored">The stored cell text.</param>
    /// <param name="target">The parsed target when the value is one of the closed set.</param>
    /// <returns><see langword="true"/> when the value is exactly <c>clipboard</c> or <c>file</c>.</returns>
    public static bool TryParse(string? stored, out PngExportTarget target)
    {
        if (string.Equals(stored?.Trim(), FileValue, StringComparison.OrdinalIgnoreCase))
        {
            target = PngExportTarget.File;
            return true;
        }

        if (string.Equals(stored?.Trim(), ClipboardValue, StringComparison.OrdinalIgnoreCase))
        {
            target = PngExportTarget.Clipboard;
            return true;
        }

        target = PngExportTarget.Clipboard;
        return false;
    }
}
