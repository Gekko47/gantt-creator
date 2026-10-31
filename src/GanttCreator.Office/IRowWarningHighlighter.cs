using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// Fills or clears row- and cell-level highlight fills on the visible
/// <c>tblGanttData</c> body, replacing the cell-note approach used by
/// <see cref="IGanttValidationReporter"/>. Fills are cosmetic
/// cell-interior fills — they never touch cell values,
/// formulas, structure, or the VeryHidden configuration sheet.
/// </summary>
/// <remarks>
/// <para>
/// Severity maps to cosmetic cell-interior colour:
/// <list type="table">
/// <listheader><term>Severity</term><description>Colour</description></listheader>
/// <item><term><see cref="GanttValidationSeverity.Error"/></term><description>Pastel red/pink — blocking issue</description></item>
/// <item><term><see cref="GanttValidationSeverity.Warning"/></term><description>Pastel amber/yellow — non-blocking warning</description></item>
/// </list>
/// Clearing a cell or row restores <c>xlColorIndexNone</c>.
/// </para>
/// <para>
/// <b>Row scope</b>: <see cref="HighlightRow"/> fills the entire visible row
/// of <c>tblGanttData</c> body cells (all table columns). <b>Cell scope</b>:
/// <see cref="HighlightCell"/> fills a single user-visible data cell.
/// </para>
/// <para>
/// <b>Differential reconciliation.</b> <see cref="ReconcileHighlights"/> is
/// the preferred entry point for refreshes. It compares the supplied targets
/// against the previously highlighted addresses and performs only the
/// minimal set of COM writes: clear addresses that are no longer needed,
/// fill addresses that are newly needed, and skip addresses that are
/// unchanged. This avoids the full-table clear-then-rewrite pattern that made
/// auto-refresh slower than manual refresh.
/// </para>
/// <para>
/// The port is deliberately primitive — no interop type crosses it — so the
/// AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c>.
/// </para>
/// </remarks>
public interface IRowWarningHighlighter
{
    /// <summary>
    /// Fills or clears the entire visible row of <c>tblGanttData</c> body
    /// cells for <paramref name="rowNumber"/>.
    /// </summary>
    /// <param name="rowNumber">The one-based body-row index.</param>
    /// <param name="severity">
    /// <see cref="GanttValidationSeverity.Error"/> for pastel red fill,
    /// <see cref="GanttValidationSeverity.Warning"/> for pastel amber fill,
    /// or any other value to clear.
    /// </param>
    /// <returns>
    /// The count of cells filled or cleared, or a typed refusal. On a refusal
    /// nothing was mutated.
    /// </returns>
    GanttValidationReportOutcome HighlightRow(int rowNumber, GanttValidationSeverity severity);

    /// <summary>
    /// Fills or clears a single user-visible data cell at
    /// <paramref name="rowNumber"/>, <paramref name="column"/>.
    /// </summary>
    /// <param name="rowNumber">The one-based body-row index.</param>
    /// <param name="column">The 1-based body-column index into <c>tblGanttData</c>.</param>
    /// <param name="severity">
    /// <see cref="GanttValidationSeverity.Error"/> for pastel red fill,
    /// <see cref="GanttValidationSeverity.Warning"/> for pastel amber fill,
    /// or any other value to clear.
    /// </param>
    /// <returns>
    /// The count of cells filled or cleared, or a typed refusal. On a refusal
    /// nothing was mutated.
    /// </returns>
    GanttValidationReportOutcome HighlightCell(int rowNumber, int column, GanttValidationSeverity severity);

    /// <summary>
    /// Clears the highlight fill from <paramref name="rowNumber"/>, or from a
    /// single cell when <paramref name="column"/> is supplied.
    /// </summary>
    /// <param name="rowNumber">The one-based body-row index.</param>
    /// <param name="column">The optional 1-based body-column index.</param>
    /// <returns>
    /// The count of cells cleared, or a typed refusal. On a refusal nothing
    /// was mutated.
    /// </returns>
    GanttValidationReportOutcome ClearHighlight(int rowNumber, int? column = null);

    /// <summary>
    /// Reconciles the highlight fills against the supplied targets, performing
    /// only the minimal set of COM writes needed to bring the sheet in line
    /// with <paramref name="currentTargets"/>. Addresses that were highlighted
    /// in a previous call but are absent from <paramref name="currentTargets"/>
    /// are cleared; addresses present in <paramref name="currentTargets"/> but
    /// not previously highlighted are filled; addresses in both sets are
    /// skipped.
    /// </summary>
    /// <param name="currentTargets">
    /// The complete set of row/cell targets that should be highlighted after
    /// this refresh, in any order. An empty list clears all previous highlights
    /// within <paramref name="rowsToReconcile"/> (or all rows when null).
    /// </param>
    /// <param name="rowsToReconcile">
    /// Optional scope: only rows in this list are reconciled. Highlight entries
    /// for rows outside the scope are left untouched. When <see langword="null"/>,
    /// all previously highlighted rows are reconciled (the full-refresh
    /// behaviour).
    /// </param>
    /// <returns>
    /// The count of cells cleared or filled, or a typed refusal. On a refusal
    /// nothing was mutated.
    /// </returns>
    GanttValidationReportOutcome ReconcileHighlights(
        IReadOnlyList<HighlightTarget> currentTargets,
        IReadOnlyList<int>? rowsToReconcile = null);
}

/// <summary>
/// One highlight target for <see cref="IRowWarningHighlighter.ReconcileHighlights"/>.
/// </summary>
/// <param name="Row">The one-based body-row index.</param>
/// <param name="Column">
/// The 1-based body-column index into <c>tblGanttData</c>, or
/// <see langword="null"/> when the entire row should be highlighted.
/// </param>
/// <param name="Severity">The severity that determines the fill colour.</param>
public readonly record struct HighlightTarget(
    int Row,
    int? Column,
    GanttValidationSeverity Severity);
