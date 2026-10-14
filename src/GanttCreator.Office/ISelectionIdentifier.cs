namespace GanttCreator.Office;

/// <summary>
/// Where the worksheet selection sits relative to the visible
/// <c>tblGanttData</c> table. Immutable. No interop type crosses the port.
/// </summary>
/// <param name="HasSelection">Whether a selection could be resolved at all.</param>
/// <param name="IsSingleBodyRow">Whether exactly one table body row is selected.</param>
/// <param name="BodyIndex">The one-based body index when a single row is selected.</param>
/// <param name="IsHeaderRow">Whether the selection is on the table header row.</param>
/// <param name="SelectedRowCount">How many table body rows the selection spans.</param>
public sealed record SelectionFacts(
    bool HasSelection,
    bool IsSingleBodyRow,
    int? BodyIndex,
    bool IsHeaderRow,
    int SelectedRowCount)
{
    /// <summary>Gets facts for "no resolvable selection".</summary>
    public static SelectionFacts None { get; } = new(false, false, null, false, 0);
}

/// <summary>
/// Standalone selection identifier over the Excel application object (R5.5).
/// Answers "which table row is selected" on demand and pushes selection-change
/// notifications via Excel's native <c>Application.SheetSelectionChange</c>.
/// </summary>
/// <remarks>
/// Standalone by owner ruling 2026-10-14: the workbook-state adapter keeps its
/// single responsibility, and this port is reused by R5.10 for fill/stroke
/// without a second probe. The surface is primitive — no interop type crosses
/// it — so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c> (the CS0433 discipline on
/// <see cref="IExcelApplicationAdapter"/>).
/// </remarks>
public interface ISelectionIdentifier
{
    /// <summary>
    /// Identifies the current selection against <c>tblGanttData</c>.
    /// Never throws: an indeterminable selection reports
    /// <see cref="SelectionFacts.None"/>.
    /// </summary>
    SelectionFacts Identify();

    /// <summary>
    /// Subscribes <paramref name="handler"/> to worksheet selection changes.
    /// </summary>
    /// <param name="handler">
    /// Invoked for each change, on the thread that raised the Excel event (the
    /// Excel main thread). The handler must not throw; the adapter guards the
    /// invocation so a throw cannot escape into Excel's event dispatch.
    /// </param>
    /// <returns>
    /// A subscription that detaches exactly once when disposed; disposing it
    /// again is a no-op. When no application object is available the returned
    /// subscription is already inert, so callers never special-case it.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="handler"/> is null.</exception>
    IDisposable SubscribeSelectionChanged(Action handler);
}
