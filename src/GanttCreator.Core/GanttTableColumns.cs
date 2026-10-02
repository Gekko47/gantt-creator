namespace GanttCreator.Core;

/// <summary>
/// How the user may see and change one managed column (R4.7C D1, ADR-0029 D8).
/// </summary>
/// <remarks>
/// <para>
/// The classification is a contract with a test, not documentation. It is the
/// single input to the initialiser's hidden/locked writes, so a column cannot
/// be presented one way and repaired another.
/// </para>
/// <para>
/// <b>Why <c>Locked</c> is not applied protection.</b> The add-in never
/// protects the worksheet: its protection guards <em>refuse</em> a protected
/// target (ADR-0008 D4). A cell's <c>Locked</c> flag is a stored format that only
/// takes effect while the sheet is protected, so setting it here defines the
/// behaviour the sheet will have if the analyst chooses to protect it — the
/// authoring columns stay editable and the engine columns do not. Nothing here
/// stops the user from editing a locked cell on an unprotected sheet, which is
/// why the engine columns are also <em>hidden</em>: hidden is visible state,
/// <c>Locked</c> alone is not.
/// </para>
/// </remarks>
public enum GanttColumnAccess
{
    /// <summary>
    /// User-visible and editable: <c>Type</c>, <c>Description</c>, <c>Start</c>,
    /// <c>Finish</c>. Unlocked, so protecting the sheet still allows authoring.
    /// </summary>
    Authoring = 0,

    /// <summary>
    /// User-visible and locked: <c>Duration</c>. The analyst reads it and the
    /// add-in writes it (R4.7F); nobody edits it by hand (ADR-0029 D3).
    /// </summary>
    ReadOnlyVisible = 1,

    /// <summary>
    /// Engine-hidden and locked: <c>Id</c>, <c>ParentId</c>, <c>SiblingOrder</c>,
    /// <c>LaneId</c>, <c>StackIndex</c>, <c>StyleKey</c>, <c>LabelPosition</c>,
    /// <c>FillColour</c>, <c>StrokeColour</c>, <c>Visible</c>, <c>SortOrder</c>.
    /// They stay physically in <c>tblGanttData</c> so one logical row is one
    /// table row (ADR-0029 D7).
    /// </summary>
    EngineHidden = 2,
}

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
    /// <param name="access">How the user may see and change the column.</param>
    public GanttTableColumn(string name, bool isRequired, GanttColumnAccess access)
    {
        ArgumentNullException.ThrowIfNull(name);

        Name = name;
        IsRequired = isRequired;
        Access = access;
    }

    /// <summary>
    /// The exact header text. This value is part of the workbook schema;
    /// renaming it is a schema migration.
    /// </summary>
    public string Name { get; }

    /// <summary>Whether the schema requires the column.</summary>
    public bool IsRequired { get; }

    /// <summary>How the user may see and change this column.</summary>
    public GanttColumnAccess Access { get; }

    /// <summary>
    /// Whether the column is hidden from the user. Only
    /// <see cref="GanttColumnAccess.EngineHidden"/> is.
    /// </summary>
    public bool IsHidden => Access == GanttColumnAccess.EngineHidden;

    /// <summary>
    /// Whether the column's cells carry the locked format. Every class except
    /// <see cref="GanttColumnAccess.Authoring"/> is locked.
    /// </summary>
    public bool IsLocked => Access != GanttColumnAccess.Authoring;
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

        Columns = Array.AsReadOnly(columns.ToArray());
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
    /// The first-release schema: the required columns in contract order, followed
    /// by the optional <c>SortOrder</c> column. R4.7A added <c>SiblingOrder</c>
    /// immediately after <c>ParentId</c> and advanced the schema version 3 → 4.
    /// R4.7C added <c>Duration</c> immediately after <c>Finish</c>, classified
    /// every column, and advanced 4 → 5.
    /// </summary>
    /// <remarks>
    /// The authoring columns are deliberately grouped in the middle of the
    /// engine columns rather than leading the table: <c>Id</c>, <c>LaneId</c> and
    /// <c>StackIndex</c> are hidden, so the first <em>visible</em> column the user
    /// sees is <c>Type</c>. Moving them would have been a second, unrecorded
    /// column-order change in the same schema bump.
    /// </remarks>
    public static GanttTableSchema Default { get; } = new(
    [
        new GanttTableColumn("Id", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("LaneId", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("StackIndex", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("Type", isRequired: true, GanttColumnAccess.Authoring),
        new GanttTableColumn("Description", isRequired: true, GanttColumnAccess.Authoring),
        new GanttTableColumn("Start", isRequired: true, GanttColumnAccess.Authoring),
        new GanttTableColumn("Finish", isRequired: true, GanttColumnAccess.Authoring),
        new GanttTableColumn("Duration", isRequired: true, GanttColumnAccess.ReadOnlyVisible),
        new GanttTableColumn("ParentId", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("SiblingOrder", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("StyleKey", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("LabelPosition", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("FillColour", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("StrokeColour", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("Visible", isRequired: true, GanttColumnAccess.EngineHidden),
        new GanttTableColumn("SortOrder", isRequired: false, GanttColumnAccess.EngineHidden),
    ]);
}
