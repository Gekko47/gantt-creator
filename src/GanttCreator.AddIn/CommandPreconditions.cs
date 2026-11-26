using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The execution-time preconditions the single-row, table-writing commands share.
/// </summary>
/// <remarks>
/// The ribbon gate may be stale by the time a command runs: the selection can change between
/// render and click, so every setter command re-reads it and refuses anything that is not
/// exactly one table body row (the per-entity ruling). The check is identical in each command
/// and only the refusal wording differs, so the wording stays with the command's own
/// translator and the check lives here.
/// </remarks>
internal static class CommandPreconditions
{
    /// <summary>
    /// Requires the current selection to be exactly one table body row, presenting the
    /// command's own refusal when it is not.
    /// </summary>
    /// <param name="selection">The standalone selection identifier, re-read rather than taken from the snapshot.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    /// <param name="refusalMessage">Translates the offending facts into this command's wording.</param>
    /// <returns><see langword="true"/> when the command may proceed.</returns>
    /// <exception cref="ArgumentNullException">An argument is <see langword="null"/>.</exception>
    internal static bool TryRequireSingleBodyRow(
        ISelectionIdentifier selection,
        Action<string> presenter,
        Func<SelectionFacts, string> refusalMessage)
    {
        ArgumentNullException.ThrowIfNull(selection);
        ArgumentNullException.ThrowIfNull(presenter);
        ArgumentNullException.ThrowIfNull(refusalMessage);

        SelectionFacts facts = selection.Identify();
        if (facts.IsSingleBodyRow)
        {
            return true;
        }

        CommandBoundary.SafePresent(presenter, refusalMessage(facts));
        return false;
    }
}
