namespace GanttCreator.Office;

/// <summary>Appends one defaulted row to the visible Gantt data table.</summary>
public interface IGanttRowInserter
{
    /// <summary>Appends one row and returns its body index, or a typed refusal.</summary>
    /// <param name="type">The catalogue type to create.</param>
    /// <param name="nextId">The one-shot stable row-ID generator.</param>
    /// <returns>The insertion result.</returns>
    GanttRowInsertOutcome Insert(Core.GanttEntityType type, Func<Core.GanttRowId> nextId);
}
