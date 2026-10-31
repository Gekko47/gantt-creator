namespace GanttCreator.Office;

/// <summary>
/// Analyzes a worksheet cell-change address to decide whether it should
/// trigger an automatic Gantt refresh. Only changes to the <c>Id</c> and
/// <c>Type</c> columns trigger auto-refresh; <c>Start</c>/<c>Finish</c>
/// require a manual Refresh, and <c>Duration</c> is engine-driven.
/// </summary>
/// <remarks>
/// The analyzer is Office-aware: it maps the changed range's Excel address
/// to a table column index and compares it against the schema's <c>Id</c>
/// and <c>Type</c> column positions.
/// </remarks>
public interface ISheetChangeAnalyzer
{
    /// <summary>
    /// Gets whether a change at <paramref name="changedAddress"/> should
    /// trigger an automatic Gantt refresh.
    /// </summary>
    /// <param name="changedAddress">
    /// The Excel-style address of the changed range (for example
    /// <c>"$C$5"</c>), or <see langword="null"/> when the address cannot be
    /// determined. Returns <see langword="false"/> when the address is
    /// null, unparseable, or outside <c>tblGanttData</c>'s <c>Id</c>/<c>Type</c>
    /// columns.
    /// </param>
    /// <returns>
    /// <see langword="true"/> when the change warrants an automatic refresh;
    /// <see langword="false"/> otherwise.
    /// </returns>
    bool ShouldAutoRefresh(string? changedAddress);
}
