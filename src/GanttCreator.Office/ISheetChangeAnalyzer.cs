namespace GanttCreator.Office;

/// <summary>
/// Analyzes a worksheet cell-change address to decide whether it should
/// trigger an automatic Gantt refresh. Chart-affecting table body columns
/// trigger auto-refresh; engine-derived <c>Duration</c> does not.
/// </summary>
/// <remarks>
/// The analyzer is Office-aware: it maps the changed range's Excel address
/// to table columns and compares the changed range with the chart-affecting
/// schema fields.
/// </remarks>
public interface ISheetChangeAnalyzer
{
    /// <summary>
    /// Gets whether a change at <paramref name="changedAddress"/> should
    /// trigger an automatic Gantt refresh.
    /// </summary>
    /// <param name="changedAddress">
    /// The Excel-style address of the changed range (for example
    /// <c>"Book.xlsx|Gantt Data!$C$5"</c>), or <see langword="null"/> when the
    /// address cannot be determined. Returns <see langword="false"/> when the
    /// address is null, unparseable, outside the active workbook's
    /// <c>tblGanttData</c> body, or affects only engine-derived columns.
    /// </param>
    /// <returns>
    /// <see langword="true"/> when the change warrants an automatic refresh;
    /// <see langword="false"/> otherwise.
    /// </returns>
    bool ShouldAutoRefresh(string? changedAddress);
}
