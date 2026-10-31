namespace GanttCreator.Core;

/// <summary>
/// Records worksheet cell changes since the last successful refresh, so an
/// auto-refresh can target only the affected rows instead of re-reading,
/// re-validating, and re-rendering the whole table.
/// </summary>
/// <remarks>
/// <para>
/// Each entry is <c>{RowNumber, Column, OldValue, NewValue}</c>. Column is the
/// 1-based column index into <c>tblGanttData</c> body cells, matching the
/// <see cref="GanttTableColumn"/> ordinal positions. <c>OldValue</c> is
/// captured before the edit; <c>NewValue</c> is the post-edit cell value.
/// </para>
/// <para>
/// The tracker is reset (<see cref="Reset"/>) after each successful refresh so
/// it only ever records changes since the last refresh completed.
/// </para>
/// <para>
/// The Core project has no Office, Excel-DNA, or UI dependency, so this type
/// names no interop types and holds values as <see langword="object"/>.
/// </para>
/// </remarks>
public sealed class ChangeTracker
{
    private readonly List<CellChange> _changes = [];

    /// <summary>The changes recorded since the last reset.</summary>
    public IReadOnlyList<CellChange> Changes => _changes;

    /// <summary>Whether any change has been recorded since the last reset.</summary>
    public bool HasChanges => _changes.Count > 0;

    /// <summary>Records one cell change.</summary>
    /// <param name="rowNumber">The one-based body-row index.</param>
    /// <param name="column">The 1-based body-column index.</param>
    /// <param name="oldValue">The value before the edit.</param>
    /// <param name="newValue">The value after the edit.</param>
    public void Record(int rowNumber, int column, object? oldValue, object? newValue) => _changes.Add(new CellChange(rowNumber, column, oldValue, newValue));

    /// <summary>Clears all recorded changes.</summary>
    public void Reset() => _changes.Clear();

    /// <summary>Returns the distinct row numbers that have at least one recorded change.</summary>
    public IReadOnlyList<int> ChangedRows()
    {
        List<int> rows = [];
        foreach (CellChange change in _changes)
        {
            if (!rows.Contains(change.RowNumber))
            {
                rows.Add(change.RowNumber);
            }
        }
        rows.Sort();
        return rows;
    }

    /// <summary>Returns whether a specific row has any recorded change.</summary>
    public bool RowChanged(int rowNumber) => _changes.Any(c => c.RowNumber == rowNumber);
}

/// <summary>One recorded cell change.</summary>
/// <param name="RowNumber">The one-based body-row index.</param>
/// <param name="Column">The 1-based body-column index.</param>
/// <param name="OldValue">The value before the edit.</param>
/// <param name="NewValue">The value after the edit.</param>
public readonly record struct CellChange(int RowNumber, int Column, object? OldValue, object? NewValue);
