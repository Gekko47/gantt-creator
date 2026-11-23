namespace GanttCreator.AddIn;

/// <summary>
/// The production <see cref="ISaveFileDialogPort"/>: a WinForms
/// <see cref="SaveFileDialog"/> configured for PNG (R8.8 D2).
/// </summary>
/// <remarks>
/// <para>
/// <b>STA, on the caller's thread.</b> <see cref="CommonDialog.ShowDialog()"/> runs the
/// dialog on the calling thread's message loop. The ribbon command already runs on Excel's
/// main STA thread, so no thread is created, nothing is marshalled, and COM is never
/// touched from another apartment (the architecture's COM rule).
/// </para>
/// <para>
/// <b>Overwrite is asked, then the writer remains the mechanical authority.</b>
/// <see cref="SaveFileDialog.OverwritePrompt"/> asks the user before an existing file is
/// replaced; the R8.6 atomic writer still performs the replace and validates the bytes it
/// wrote, so a success is never reported on the strength of the dialog alone.
/// </para>
/// </remarks>
internal sealed class WindowsSaveFileDialogPort : ISaveFileDialogPort
{
    /// <inheritdoc/>
    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Microsoft.Globalization",
        "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    public string? PickPath(string suggestedFileName)
    {
        using SaveFileDialog dialog = new()
        {
            Title = "Export PNG",
            Filter = "PNG image (*.png)|*.png",
            DefaultExt = "png",
            AddExtension = true,
            OverwritePrompt = true,
            FileName = suggestedFileName,
        };

        return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
    }
}
