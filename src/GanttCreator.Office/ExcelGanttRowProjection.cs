using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>
/// Projects a raw Gantt table body into the neutral <see cref="GanttEvent"/>
/// list the Core hierarchy and outline planners consume. Shared by the
/// parent/child mutator (R5.4 D12 <c>Finish</c>) and the group-aware row
/// inserter (R5.13), so both project the table through one authority rather
/// than two drifting copies.
/// </summary>
/// <remarks>
/// <para>
/// The projection is pure: it reads the raw cell matrix the Office adapters
/// already bulk-read and populates only the fields the hierarchy and outline
/// planners read (<c>RowNumber</c>, <c>Id</c>, <c>Type</c>, <c>ParentId</c>);
/// the rest stay at their defaults.
/// </para>
/// <para>
/// A row whose <c>Id</c> cannot be parsed makes the table unusable. The
/// mutator maps that to a typed refusal (<c>SelectionUnknown</c> for the
/// selected row, <c>TableInvalid</c> for any other); the inserter degrades
/// to the pre-D12 behaviour -- the insert stands and the next Refresh
/// rebuilds the outline -- rather than failing work that already landed.
/// </para>
/// </remarks>
internal static class ExcelGanttRowProjection
{
    /// <summary>
    /// Projects the raw body into the neutral events the Core planner
    /// consumes. Only the fields the hierarchy and outline planners read
    /// are populated; the rest stay at their defaults.
    /// </summary>
    /// <param name="body">The raw body rows, in body order.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="failedIndex">
    /// The zero-based index of the first row whose <c>Id</c> could not be
    /// parsed, when the projection returns <see langword="null"/>.
    /// </param>
    /// <returns>The projected events, or <see langword="null"/> when a row's <c>Id</c> is unparseable.</returns>
    internal static List<GanttEvent>? ProjectRows(
        List<object?[]> body,
        int[] columnMap,
        out int? failedIndex)
    {
        failedIndex = null;
        var events = new List<GanttEvent>(body.Count);
        for (var index = 0; index < body.Count; index++)
        {
            var row = body[index];
            if (!GanttRowId.TryParse(CellText(row, columnMap, "Id"), out GanttRowId? id) || id is null)
            {
                failedIndex = index;
                return null;
            }

            _ = EntityTypeCatalog.TryParse(
                CellText(row, columnMap, "Type"),
                out GanttEntityType type);
            var parentText = CellText(row, columnMap, "ParentId");
            GanttRowId? parentId =
                GanttRowId.TryParse(parentText, out GanttRowId? parsedParent) && parsedParent is not null
                    ? parsedParent
                    : null;

            events.Add(new GanttEvent(
                index + 1,
                id,
                null,
                null,
                type,
                null,
                null,
                null,
                parentId,
                null,
                null,
                null,
                null,
                true,
                null));
        }

        return events;
    }

    /// <summary>
    /// Reads one text cell from a raw body row by schema column name.
    /// </summary>
    /// <param name="row">The raw row, one value per table column.</param>
    /// <param name="columnMap">The schema-column to table-column index map.</param>
    /// <param name="columnName">The schema column name.</param>
    /// <returns>The cell's text, or <see cref="string.Empty"/>.</returns>
    internal static string CellText(
        object?[] row,
        int[] columnMap,
        string columnName)
    {
        var schemaIndex = SchemaIndexOf(columnName);
        var tableIndex = columnMap[schemaIndex];
        return tableIndex < 1 || tableIndex > row.Length
            ? string.Empty
            : row[tableIndex - 1] as string ?? string.Empty;
    }

    /// <summary>
    /// Locates a schema column's index by name.
    /// </summary>
    /// <param name="columnName">The schema column name.</param>
    /// <returns>The zero-based schema index.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown for a name that is not a Gantt schema column. A new column
    /// name is a schema change, not a value this method may guess.
    /// </exception>
    internal static int SchemaIndexOf(string columnName)
    {
        for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            if (GanttTableSchema.Default.Columns[index].Name == columnName)
            {
                return index;
            }
        }

        throw new InvalidOperationException($"'{columnName}' is not a Gantt schema column.");
    }
}