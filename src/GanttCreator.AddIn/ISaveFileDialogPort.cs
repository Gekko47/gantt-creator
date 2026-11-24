namespace GanttCreator.AddIn;

/// <summary>
/// Picks the file the PNG export writes (R8.8 D2). The one production implementation
/// wraps a WinForms <see cref="SaveFileDialog"/> and is called on
/// Excel's main STA thread from the ribbon command — never through <c>Task.Run</c>, which
/// is forbidden around COM.
/// </summary>
/// <remarks>
/// The port exists so the command's stage order (and its dialog-cancelled silence) is
/// testable without a host: a unit test substitutes a fixed path or a cancel, and the
/// production adapter is the only place a dialog is shown.
/// </remarks>
internal interface ISaveFileDialogPort
{
    /// <summary>Shows the save dialog and returns the chosen path.</summary>
    /// <param name="suggestedFileName">The file name the dialog is pre-filled with.</param>
    /// <param name="initialDirectory">
    /// The folder the dialog opens at — the remembered export folder, else the workbook's own
    /// folder (R8.8 D9) — or null/blank to let the host choose.
    /// </param>
    /// <returns>The chosen path, or <see langword="null"/> when the user cancelled.</returns>
    string? PickPath(string suggestedFileName, string? initialDirectory);
}
