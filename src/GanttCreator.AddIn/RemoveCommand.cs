using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Removes the row under the active cell (R5.4 D6): the row is
/// deleted from <c>tblGanttData</c> and its children are re-homed
/// to its parent. Destructive under ADR-0008, so the command
/// confirms first and No is the safe default that mutates nothing.
/// The command owns the confirmation and outcome-to-message
/// translation; the selection resolution, planning and cell writes
/// belong to the Office mutator port, exactly as the promote,
/// demote, move and add-row commands delegate theirs.
/// </summary>
internal static class RemoveCommand
{
    /// <summary>
    /// The session-scoped destructive-confirmation opt-out
    /// (ADR-0042, amending ADR-0008 D3). Process-scoped and
    /// in-memory only: once the user confirms a delete with
    /// "don't ask again this session" checked, later deletes
    /// in the same Excel session skip the confirmation dialog.
    /// It is not persisted and resets when the add-in reloads.
    /// </summary>
    private static readonly ConfirmationSession _session = new();

    /// <summary>Runs the production remove command in the current Excel session.</summary>
    internal static void RunForExcel() =>
        Run(
            new ExcelParentChildMutator(
                ExcelDna.Integration.ExcelDnaUtil.Application,
                outlineGroups: new ExcelOutlineGroupWriter(ExcelDna.Integration.ExcelDnaUtil.Application)),
            CommandErrorDialog.Confirm,
            CommandErrorDialog.Show,
            RefreshSheetCommand.RunForExcel);

    /// <summary>Runs an injected remove command.</summary>
    /// <param name="mutator">The parent/child mutation port.</param>
    /// <param name="confirm">
    /// Asks the user to approve the destructive change with the
    /// confirmation body and the session whose opt-out checkbox
    /// the dialog drives; the dialog renders the shared
    /// destructive-command heading
    /// (<see cref="DestructiveCommandPolicy.ConfirmationMainInstruction"/>).
    /// </param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="onMutated">
    /// Invoked after the row was actually deleted (R5.4
    /// D11-amended): the delete host-removes the anchored
    /// shapes, so the chart must be rebuilt immediately. Not
    /// invoked on a declined confirmation or a refusal.
    /// </param>
    internal static void Run(
        IParentChildMutator mutator,
        Func<string, ConfirmationSession?, bool> confirm,
        Action<string> presenter,
        Action? onMutated = null)
        => Run(mutator, confirm, presenter, onMutated, _session);

    /// <summary>
    /// Runs an injected remove command against an explicit
    /// session, so tests exercise the opt-out in isolation
    /// without touching the process-scoped session.
    /// </summary>
    /// <param name="mutator">The parent/child mutation port.</param>
    /// <param name="confirm">
    /// Asks the user to approve the destructive change with the
    /// confirmation body and the session whose opt-out checkbox
    /// the dialog drives; the dialog renders the shared
    /// destructive-command heading
    /// (<see cref="DestructiveCommandPolicy.ConfirmationMainInstruction"/>).
    /// </param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="onMutated">
    /// Invoked after the row was actually deleted (R5.4
    /// D11-amended): the delete host-removes the anchored
    /// shapes, so the chart must be rebuilt immediately. Not
    /// invoked on a declined confirmation or a refusal.
    /// </param>
    /// <param name="session">
    /// The session whose opt-out gates the confirmation. The
    /// production entry point passes the process-scoped
    /// session; tests pass a fresh instance.
    /// </param>
    internal static void Run(
        IParentChildMutator mutator,
        Func<string, ConfirmationSession?, bool> confirm,
        Action<string> presenter,
        Action? onMutated,
        ConfirmationSession session)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(confirm);
        ArgumentNullException.ThrowIfNull(presenter);
        ArgumentNullException.ThrowIfNull(session);

        // ADR-0042 D4: once the user has opted out for this
        // session, the delete proceeds without the confirmation
        // dialog. The opt-out is recorded only by an explicit
        // "don't ask again this session" confirmation
        // (ConfirmationSession.RecordChoice); a fresh session
        // always re-arms the ADR-0008 D3 guard.
        if (!session.DontAskAgain)
        {
            // ADR-0008 D2/D3: the confirmation names the command
            // and states that it cannot be undone. Declining is
            // the safe default: the mutator is never called and
            // nothing is mutated.
            ConfirmationText confirmation = DestructiveCommandPolicy.BuildConfirmationText(
                DestructiveCommandPolicy.CommandClass.RemoveRow);
            if (!confirm(confirmation.Body, session))
            {
                return;
            }
        }

        ParentChildMutationOutcome outcome = mutator.Remove();

        // R2.2/R2.6: success is silent; the ribbon state refresh
        // handles any enabled-state consequences (R1.5 D3).
        if (outcome.Succeeded)
        {
            // R5.4 D11-amended: the delete removed a worksheet
            // row and its anchored shapes, so the chart is
            // rebuilt now — even when the rehomed-children
            // resequence wrote zero cells (a lone row delete
            // still changed the sheet).
            onMutated?.Invoke();
            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
        Present(presenter, ParentChildRefusalTranslator.TranslateRemove(outcome));
    }

    /// <summary>
    /// Shows a message, degrading to silence when the host refuses to present it. A
    /// dialog failure must never turn a completed command into an escaping exception.
    /// </summary>
    /// <param name="presenter">The user-safe message presenter.</param>
    /// <param name="message">The message to present.</param>
    private static void Present(Action<string> presenter, string message)
    {
        // CA1031: presenter failure is outside command behaviour.
#pragma warning disable CA1031
        try
        {
            presenter(message);
        }
        catch
        {
            // Intentionally empty: a dialog failure must not escape into Excel.
        }
#pragma warning restore CA1031
    }
}
