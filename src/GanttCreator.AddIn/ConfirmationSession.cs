namespace GanttCreator.AddIn;

/// <summary>
/// Session-scoped state for the destructive-command confirmation
/// opt-out (ADR-0042, amending ADR-0008 D3). The delete-row
/// confirmation dialog offers a "don't ask again this session"
/// verification checkbox; once the user opts in and confirms,
/// <see cref="DontAskAgain"/> is set and later deletes in the
/// same Excel session skip the confirmation dialog.
/// </summary>
/// <remarks>
/// <para>
/// "Session" is the add-in's loaded lifetime (one Excel process).
/// The state is process-scoped and held in memory only: it is not
/// persisted to the workbook or to the <c>_GanttCreatorConfig</c>
/// worksheet (which stores catalogues, styles, metrics and
/// configuration metadata, never session state), and it resets
/// when the add-in reloads.
/// </para>
/// <para>
/// A mutable reference type so the confirmation dialog can record
/// the user's choice and the command can read it without a static
/// singleton leaking into tests; tests construct a fresh instance
/// per case.
/// </para>
/// </remarks>
internal sealed class ConfirmationSession
{
    /// <summary>
    /// Whether the user asked not to be prompted again this
    /// session. While true, a delete-row command proceeds without
    /// showing its confirmation dialog.
    /// </summary>
    public bool DontAskAgain { get; set; }

    /// <summary>
    /// Records the user's opt-out choice after a destructive
    /// command's confirmation dialog closed. Pure logic, unit-tested
    /// without a dialog host: the choice is recorded only when the
    /// user actually confirmed the change; a declined dialog leaves
    /// the session untouched so the guard stays in force.
    /// </summary>
    /// <param name="confirmed">Whether the user confirmed the change.</param>
    /// <param name="dontAskAgain">
    /// Whether the user checked "don't ask again this session".
    /// </param>
    public void RecordChoice(bool confirmed, bool dontAskAgain)
    {
        if (confirmed)
        {
            DontAskAgain = dontAskAgain;
        }
    }
}
