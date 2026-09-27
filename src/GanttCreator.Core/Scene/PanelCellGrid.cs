namespace GanttCreator.Core.Scene;

/// <summary>One measured data-panel column.</summary>
/// <param name="Name">
/// The exact schema display name, from <see cref="GanttTableSchema.Default"/>.
/// </param>
/// <param name="WidthPt">The measured column width in points.</param>
/// <param name="Alignment">The resolved body-text alignment for this column.</param>
/// <remarks>
/// The alignment is <see cref="GanttTextAlignment"/> rather than a panel-local
/// type so the grid's measured alignment flows straight into the emitted
/// <see cref="SceneText"/> without a conversion that could drift (ADR-0018).
/// </remarks>
public sealed record PanelColumn(string Name, double WidthPt, GanttTextAlignment Alignment = GanttTextAlignment.Left);

/// <summary>The reason a measured cell grid was refused.</summary>
public enum PanelCellGridRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>No columns were supplied.</summary>
    NoColumns = 1,

    /// <summary>A column width was not finite, or was not positive.</summary>
    NonPositiveWidth = 2,

    /// <summary>The row height was not finite, or was not positive.</summary>
    NonPositiveRowHeight = 3,

    /// <summary>A column name was blank or whitespace.</summary>
    BlankColumnName = 4,

    /// <summary>Two columns carried the same schema display name.</summary>
    DuplicateColumn = 5,

    /// <summary>A column alignment was not a defined value.</summary>
    UndefinedAlignment = 6,

    /// <summary>A required schema column was absent from the grid.</summary>
    MissingRequiredColumn = 7,
}

/// <summary>The typed result of validating a measured cell grid.</summary>
/// <param name="Grid">The validated grid, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record PanelCellGridCreationOutcome(PanelCellGrid? Grid, PanelCellGridRefusal? Refusal)
{
    /// <summary>Gets whether the grid was valid.</summary>
    public bool Succeeded => Grid is not null;
}

/// <summary>
/// The measured data-panel input for the §3/§4 export composition: the
/// included columns with their exact point widths, the row height, and the
/// required column names that must appear exactly once.
/// </summary>
/// <remarks>
/// <para>
/// This is an <em>input</em> model, not a measuring one. The builder never
/// measures anything: a Phase 4 or Phase 5 adapter reads the live column widths
/// and row heights and constructs this, and tests inject fixtures. Keeping the
/// model minimal is deliberate, so the contract cannot drift away from what a
/// live adapter can actually produce.
/// </para>
/// <para>
/// Column order is the caller's order and is significant: §4's header follows
/// the schema display names in schema order, and the emitted cell bounds depend
/// on the order, so it is preserved rather than sorted.
/// </para>
/// </remarks>
public sealed record PanelCellGrid
{
    private PanelCellGrid(
        IReadOnlyList<PanelColumn> columns,
        double rowHeightPt,
        IReadOnlyList<string> requiredColumns)
    {
        Columns = columns;
        RowHeightPt = rowHeightPt;
        RequiredColumns = requiredColumns;
    }

    /// <summary>The included columns, in caller order.</summary>
    public IReadOnlyList<PanelColumn> Columns { get; }

    /// <summary>The measured body row height in points.</summary>
    public double RowHeightPt { get; }

    /// <summary>
    /// The schema column names that must be present exactly once. §3 requires
    /// required columns to exist once, and §4 requires unique header names.
    /// </summary>
    public IReadOnlyList<string> RequiredColumns { get; }

    /// <summary>The total panel width in points, the sum of the column widths.</summary>
    public double TotalWidthPt
    {
        get
        {
            double total = 0;
            foreach (PanelColumn column in Columns)
            {
                total += column.WidthPt;
            }

            return total;
        }
    }

    /// <summary>Validates and creates a measured cell grid.</summary>
    /// <param name="columns">The included columns in caller order.</param>
    /// <param name="rowHeightPt">The measured body row height in points.</param>
    /// <param name="requiredColumns">The schema names that must appear exactly once.</param>
    /// <returns>A typed result or refusal.</returns>
    public static PanelCellGridCreationOutcome TryCreate(
        IReadOnlyList<PanelColumn>? columns,
        double rowHeightPt,
        IReadOnlyList<string>? requiredColumns)
    {
        if (columns is null)
        {
            return Refused(PanelCellGridRefusal.NullRequest);
        }

        if (columns.Count == 0)
        {
            return Refused(PanelCellGridRefusal.NoColumns);
        }

        if (!double.IsFinite(rowHeightPt) || rowHeightPt <= 0)
        {
            return Refused(PanelCellGridRefusal.NonPositiveRowHeight);
        }

        foreach (PanelColumn column in columns)
        {
            if (column is null || string.IsNullOrWhiteSpace(column.Name))
            {
                return Refused(PanelCellGridRefusal.BlankColumnName);
            }

            if (!double.IsFinite(column.WidthPt) || column.WidthPt <= 0)
            {
                return Refused(PanelCellGridRefusal.NonPositiveWidth);
            }

            if (!Enum.IsDefined(column.Alignment))
            {
                return Refused(PanelCellGridRefusal.UndefinedAlignment);
            }
        }

        // A repeated schema name would emit two header cells with identical text
        // at different bounds, which §4's unique-name rule forbids, so it is
        // refused rather than collapsed.
        if (columns.Select(column => column.Name).Distinct(StringComparer.Ordinal).Count() != columns.Count)
        {
            return Refused(PanelCellGridRefusal.DuplicateColumn);
        }

        HashSet<string> present = new(columns.Select(column => column.Name), StringComparer.Ordinal);
        foreach (var required in requiredColumns ?? [])
        {
            if (string.IsNullOrWhiteSpace(required) || !present.Contains(required))
            {
                return Refused(PanelCellGridRefusal.MissingRequiredColumn);
            }
        }

        return new PanelCellGridCreationOutcome(
            new PanelCellGrid([.. columns], rowHeightPt, [.. requiredColumns ?? []]),
            null);
    }

    private static PanelCellGridCreationOutcome Refused(PanelCellGridRefusal refusal) => new(null, refusal);
}
