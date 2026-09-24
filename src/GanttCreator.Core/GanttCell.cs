namespace GanttCreator.Core;

/// <summary>One typed neutral worksheet cell with its preserved raw state.</summary>
/// <typeparam name="T">The supported value type for the column.</typeparam>
/// <param name="State">The raw cell state.</param>
/// <param name="Value">The normalized value when <paramref name="State"/> is <see cref="GanttCellState.Value"/>.</param>
/// <param name="ErrorCode">The neutral Excel error code when applicable.</param>
public readonly record struct GanttCell<T>(
    GanttCellState State,
    T? Value,
    GanttExcelErrorCode? ErrorCode = null)
{
    /// <summary>Returns whether this cell has a supported typed value.</summary>
    public bool HasValue => State == GanttCellState.Value;
}
