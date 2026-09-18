namespace GanttCreator.Core;

/// <summary>
/// One named column of the visible <c>tblGanttData</c> worksheet table.
/// Immutable.
/// </summary>
public sealed class GanttTableColumn
{
    /// <summary>
    /// Initialises a column descriptor.
    /// </summary>
    /// <param name="name">The exact header text; a schema value.</param>
    /// <param name="isRequired">Whether the schema requires the column.</param>
    public GanttTableColumn(string name, bool isRequired)
    {
        ArgumentNullException.ThrowIfNull(name);

        Name = name;
        IsRequired = isRequired;
    }

    /// <summary>
    /// The exact header text. This value is part of the workbook schema;
    /// renaming it is a schema migration.
    /// </summary>
    public string Name { get; }

    /// <summary>Whether the schema requires the column.</summary>
    public bool IsRequired { get; }
}

/// <summary>
/// The ordered column schema of the visible Gantt data table
/// (<c>tblGanttData</c>). Header names and their order are part of the
/// workbook schema; optional columns beyond <see cref="Default"/>'s
/// <c>SortOrder</c> require an approved schema ADR and a schema-version bump
/// (see <see cref="GanttSchemaVersion.CurrentSchemaVersion"/>).
/// </summary>
public sealed class GanttTableSchema
{
    /// <summary>The Excel Table name of the visible Gantt data table.</summary>
    public const string TableName = "tblGanttData";

    /// <summary>
    /// Initialises a schema from ordered column descriptors.
    /// </summary>
    /// <param name="columns">The columns in header-row order.</param>
    /// <exception cref="ArgumentException">Thrown for an empty column list or an empty, whitespace, or duplicate column name.</exception>
    public GanttTableSchema(IReadOnlyList<GanttTableColumn> columns)
    {
        ArgumentNullException.ThrowIfNull(columns);

        if (columns.Count == 0)
        {
            throw new ArgumentException("Schema must contain at least one column.", nameof(columns));
        }

        foreach (GanttTableColumn column in columns)
        {
            ArgumentNullException.ThrowIfNull(column, nameof(columns));
            if (string.IsNullOrWhiteSpace(column.Name))
            {
                throw new ArgumentException("Column name must not be empty.", nameof(columns));
            }
        }

        IGrouping<string, GanttTableColumn>? duplicate = columns
            .GroupBy(c => c.Name, StringComparer.Ordinal)
            .FirstOrDefault(g => g.Count() > 1);
        if (duplicate is not null)
        {
            throw new ArgumentException($"Duplicate column name '{duplicate.Key}'.", nameof(columns));
        }

        Columns = [.. columns];
    }

    /// <summary>Gets the columns in header-row order.</summary>
    public IReadOnlyList<GanttTableColumn> Columns { get; }

    /// <summary>
    /// Resolves a column by exact header name (Ordinal comparison). Case
    /// variants and differently-cased Excel headers do not resolve.
    /// </summary>
    /// <param name="name">The header text to look up.</param>
    /// <param name="column">The resolved column when the method returns <see langword="true"/>.</param>
    /// <returns><see langword="true"/> when a column with the exact name exists.</returns>
    public bool TryGetColumn(string? name, out GanttTableColumn? column)
    {
        if (string.IsNullOrEmpty(name))
        {
            column = null;
            return false;
        }

        foreach (GanttTableColumn candidate in Columns)
        {
            if (string.Equals(candidate.Name, name, StringComparison.Ordinal))
            {
                column = candidate;
                return true;
            }
        }

        column = null;
        return false;
    }

    /// <summary>
    /// The first-release schema: the 13 required columns in contract order,
    /// followed by the optional <c>SortOrder</c> column.
    /// </summary>
    public static GanttTableSchema Default { get; } = new(
    [
        new GanttTableColumn("Id", isRequired: true),
        new GanttTableColumn("LaneId", isRequired: true),
        new GanttTableColumn("StackIndex", isRequired: true),
        new GanttTableColumn("Type", isRequired: true),
        new GanttTableColumn("Description", isRequired: true),
        new GanttTableColumn("Start", isRequired: true),
        new GanttTableColumn("Finish", isRequired: true),
        new GanttTableColumn("ParentId", isRequired: true),
        new GanttTableColumn("StyleKey", isRequired: true),
        new GanttTableColumn("LabelPosition", isRequired: true),
        new GanttTableColumn("FillColour", isRequired: true),
        new GanttTableColumn("StrokeColour", isRequired: true),
        new GanttTableColumn("Visible", isRequired: true),
        new GanttTableColumn("SortOrder", isRequired: false),
    ]);
}
