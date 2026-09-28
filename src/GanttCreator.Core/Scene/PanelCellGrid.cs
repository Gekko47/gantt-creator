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

    /// <summary>No body row height was supplied.</summary>
    NoRows = 4,

    /// <summary>The header row height was not finite, or was not positive.</summary>
    NonPositiveHeaderHeight = 5,

    /// <summary>A column name was blank or whitespace.</summary>
    BlankColumnName = 6,

    /// <summary>Two columns carried the same schema display name.</summary>
    DuplicateColumn = 7,

    /// <summary>A column alignment was not a defined value.</summary>
    UndefinedAlignment = 8,

    /// <summary>A required schema column was absent from the grid.</summary>
    MissingRequiredColumn = 9,
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
        IReadOnlyList<double> rowHeightsPt,
        double headerHeightPt,
        IReadOnlyList<string> requiredColumns)
    {
        Columns = columns;
        RowHeightsPt = rowHeightsPt;
        HeaderHeightPt = headerHeightPt;
        RequiredColumns = requiredColumns;
    }

    /// <summary>The included columns, in caller order.</summary>
    public IReadOnlyList<PanelColumn> Columns { get; }

    /// <summary>
    /// The measured body row heights in points, in worksheet order, one per body row.
    /// </summary>
    /// <remarks>
    /// A list, not a single value, because a worksheet body does not have to be
    /// uniform. Entity guide §3 requires the panel to reproduce "the exact measured
    /// cell bounds in points", and a single sample height cannot express a body whose
    /// rows differ — it can only either refuse a legitimate table or lay it out wrong.
    /// The list is positional: entry <c>n</c> is the height of the <c>n</c>th body row,
    /// which is what <see cref="PanelBuildRequest.Rows"/> must match in order and count.
    /// </remarks>
    public IReadOnlyList<double> RowHeightsPt { get; }

    /// <summary>
    /// The measured header row height in points, separate from the body rows.
    /// </summary>
    /// <remarks>
    /// §4 says the data-panel header "follows live header-cell bounds". The header row
    /// is its own Excel row with its own height, so reusing a body row's height is an
    /// assumption about the worksheet rather than a measurement of it.
    /// </remarks>
    public double HeaderHeightPt { get; }

    /// <summary>
    /// The schema column names that must be present exactly once. §3 requires
    /// required columns to exist once, and §4 requires unique header names.
    /// </summary>
    public IReadOnlyList<string> RequiredColumns { get; }

    /// <summary>The total body height in points, the sum of the row heights.</summary>
    public double TotalRowHeightPt
    {
        get
        {
            double total = 0;
            foreach (var height in RowHeightsPt)
            {
                total += height;
            }

            return total;
        }
    }

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
    /// <param name="rowHeightsPt">The measured body row heights, in worksheet order.</param>
    /// <param name="headerHeightPt">The measured header row height.</param>
    /// <param name="requiredColumns">The schema names that must appear exactly once.</param>
    /// <returns>A typed result or refusal.</returns>
    public static PanelCellGridCreationOutcome TryCreate(
        IReadOnlyList<PanelColumn>? columns,
        IReadOnlyList<double>? rowHeightsPt,
        double headerHeightPt,
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

        if (rowHeightsPt is null || rowHeightsPt.Count == 0)
        {
            return Refused(PanelCellGridRefusal.NoRows);
        }

        if (!double.IsFinite(headerHeightPt) || headerHeightPt <= 0)
        {
            return Refused(PanelCellGridRefusal.NonPositiveHeaderHeight);
        }

        foreach (var rowHeight in rowHeightsPt)
        {
            if (!double.IsFinite(rowHeight) || rowHeight <= 0)
            {
                return Refused(PanelCellGridRefusal.NonPositiveRowHeight);
            }
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
            new PanelCellGrid([.. columns], [.. rowHeightsPt], headerHeightPt, [.. requiredColumns ?? []]),
            null);
    }

    private static PanelCellGridCreationOutcome Refused(PanelCellGridRefusal refusal) => new(null, refusal);
}
