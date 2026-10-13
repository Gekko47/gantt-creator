using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Moves the row under the active cell down one position (R5.4 D5): the
/// mirror of <see cref="MoveUpCommand"/> — a top-level row swaps its window
/// with the group below it, a child swaps with its next sibling, and the
/// bottom of the set is a silent success with zero cells written (R5.4 D4).
/// The command owns outcome-to-message translation; the selection
/// resolution, planning and cell writes belong to the Office mutator port.
/// </summary>
internal static class MoveDownCommand
{
    /// <summary>Runs the production move-down command in the current Excel session.</summary>
    internal static void RunForExcel() =>
        Run(
            new ExcelParentChildMutator(
                ExcelDna.Integration.ExcelDnaUtil.Application,
                outlineGroups: new ExcelOutlineGroupWriter(ExcelDna.Integration.ExcelDnaUtil.Application)),
            new ExcelInsertedRowSelector(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Show,
            RefreshSheetCommand.RunForExcel);

    /// <summary>Runs an injected move-down command.</summary>
    /// <param name="mutator">The parent/child mutation port.</param>
    /// <param name="selector">Moves the selection onto the moved row (R5.4 D9).</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="onMutated">
    /// Invoked after a move that actually wrote cells (R5.4
    /// D11-amended): the move rotates content only, so the chart
    /// must be rebuilt immediately for the shapes to follow. Not
    /// invoked on a boundary no-op, which changed nothing.
    /// </param>
    internal static void Run(
        IParentChildMutator mutator,
        IInsertedRowSelector selector,
        Action<string> presenter,
        Action? onMutated = null)
    {
        ArgumentNullException.ThrowIfNull(mutator);
        ArgumentNullException.ThrowIfNull(selector);
        ArgumentNullException.ThrowIfNull(presenter);

        ParentChildMutationOutcome outcome = mutator.MoveDown();

        // R2.2/R2.6 and R5.4 D4: success is silent — including the boundary
        // no-op, which writes zero cells and reports nothing because the
        // user asked for a move that cannot happen.
        if (outcome.Succeeded)
        {
            // R5.4 D9: the selection follows the row to its new position so
            // the next move repeats on the same row. A boundary no-op moved
            // nothing and leaves the selection where it is. A selection
            // failure degrades to "cursor stays put", exactly as the
            // add-row command's selection move does.
            if (outcome.MovedToRowNumber is { } movedTo)
            {
                SelectMovedRow(selector, movedTo);
            }

            // R5.4 D11-amended: a content rotation leaves the chart
            // stale, so a move that wrote anything rebuilds it now.
            // The selection moved first, so the refresh's
            // ADR-0020 scope captures it and restores it onto the
            // moved row.
            if (outcome.CellsWritten > 0)
            {
                onMutated?.Invoke();
            }

            return;
        }

        // CA1031: presenter failure is outside command behaviour; it degrades to no dialog.
        Present(presenter, ParentChildRefusalTranslator.TranslateMove(outcome, moveUp: false));
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

    /// <summary>
    /// Moves the selection onto the moved row, degrading to no selection
    /// change when the host refuses. A selection failure must never turn a
    /// completed move into a user-visible error or escape into Excel.
    /// </summary>
    /// <param name="selector">The selection port.</param>
    /// <param name="bodyIndex">The moved row's new one-based body index.</param>
    private static void SelectMovedRow(IInsertedRowSelector selector, int bodyIndex)
    {
        // CA1031: a host selection failure is outside command behaviour.
#pragma warning disable CA1031
        try
        {
            selector.SelectBodyRow(bodyIndex);
        }
        catch
        {
            // Intentionally empty: the row moved; only the cursor stays put.
        }
#pragma warning restore CA1031
    }
}
