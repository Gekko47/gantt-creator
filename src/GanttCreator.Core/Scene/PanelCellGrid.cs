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

    /// <summary>
    /// A row height was not finite, or was negative. (ADR-0034)
    /// </summary>
    /// <remarks>
    /// This was <c>NonPositiveRowHeight</c>, and zero was refused. Zero is not a
    /// missing measurement: Excel reports a <b>hidden</b> row's height as zero, and a
    /// collapsed outline group hides its child rows, so refusing zero made a
    /// collapse-then-Refresh fail the whole measurement instead of laying the sheet
    /// out the way it is actually displayed. A negative height is still refused
    /// because it would place a row above the one before it.
    /// </remarks>
    NegativeRowHeight = 3,

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

    /// <summary>
    /// The measured absolute origin was absent, non-finite, or negative.
    /// </summary>
    /// <remarks>
    /// <b>Refused rather than defaulted.</b> A guessed origin is precisely the silent
    /// misalignment ADR-0030 exists to remove: the scene would place lanes at a
    /// plausible but wrong vertical position, and nothing would report it.
    /// </remarks>
    InvalidOrigin = 10,

    /// <summary>
    /// A padding row height was not finite, or was negative.
    /// </summary>
    /// <remarks>
    /// Zero is allowed and means "no padding row", which is the export case. A
    /// negative height is refused because it would place the chart frame's edge
    /// inside its own content, producing an inverted chart that no later stage
    /// would catch.
    /// </remarks>
    InvalidPaddingHeight = 11,
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
        IReadOnlyList<string> requiredColumns,
        double originTopPt,
        double originLeftPt,
        double topPaddingHeightPt,
        double bottomPaddingHeightPt)
    {
        Columns = columns;
        RowHeightsPt = rowHeightsPt;
        HeaderHeightPt = headerHeightPt;
        RequiredColumns = requiredColumns;
        OriginTopPt = originTopPt;
        OriginLeftPt = originLeftPt;
        TopPaddingHeightPt = topPaddingHeightPt;
        BottomPaddingHeightPt = bottomPaddingHeightPt;

        // The cumulative tops are derived ONCE, here, so a lane anchor and a panel
        // row cannot disagree about where a body row starts. Entry n is
        // OriginTopPt plus the heights of every row above n, which is the row's
        // absolute worksheet top edge; a zero-height row therefore contributes no
        // space and the rows below it share its top (ADR-0034).
        var tops = new double[rowHeightsPt.Count];
        var running = originTopPt;
        for (var index = 0; index < rowHeightsPt.Count; index++)
        {
            tops[index] = running;
            running += rowHeightsPt[index];
        }

        RowTopsPt = tops;
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
    /// The absolute worksheet top edge of each body row, in points, in worksheet
    /// order, one per body row (ADR-0034 D1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Entry <c>n</c> is <see cref="OriginTopPt"/> plus the sum of the heights of the
    /// rows above it. This is the value that makes ADR-0026 D3's
    /// <c>Excel Top == Scene lane Top</c> representable: a lane is anchored to the row
    /// it renders on, rather than to the running sum of the lanes above it, so a body
    /// row that owns no lane (a <c>Delineator</c>, a projected child) leaves its own
    /// band empty instead of pulling every later lane upwards.
    /// </para>
    /// <para>
    /// <b>A zero-height row consumes no space, deliberately.</b> Excel reports a
    /// hidden row's height as zero, and a collapsed outline group hides its child
    /// rows, so a zero entry is a real measurement rather than a missing one: the rows
    /// below it share its top, exactly as the worksheet lays them out.
    /// </para>
    /// </remarks>
    public IReadOnlyList<double> RowTopsPt { get; }

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

    /// <summary>
    /// The absolute worksheet Y of the <em>first body row's top edge</em>, in points.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Why this is the input the live chart's vertical origin was missing.</b>
    /// Until ADR-0030 the scene had no idea where the body began on the sheet: the
    /// plot's top came from the size preset's page coordinates (a token sum of
    /// 58pt), and <c>LaneLayoutBuilder</c> stacked lanes down from that. A lane
    /// therefore could not coincide with its own row, because nothing in the model
    /// said where the row was. This is the measurement that makes ADR-0026 D3's
    /// <c>Excel Top == Scene lane Top</c> representable rather than merely unmet.
    /// </para>
    /// <para>
    /// It is the first <em>body</em> row, so the header row's height is already
    /// accounted for and this value is directly the plot's top (D1). It is not the
    /// header's top and not the reserved row's.
    /// </para>
    /// </remarks>
    public double OriginTopPt { get; }

    /// <summary>
    /// The absolute worksheet X of the panel's left edge, in points.
    /// </summary>
    /// <remarks>
    /// The table does not start at column A — the engine columns preceding
    /// <c>Type</c> are hidden but still occupy worksheet positions — so the panel's
    /// left edge is a measurement, not a constant. It is what lets the plot's left
    /// edge be placed against the panel's right edge in real sheet coordinates.
    /// </remarks>
    public double OriginLeftPt { get; }

    /// <summary>
    /// The measured height of the top padding row, in points (ADR-0031 D2).
    /// </summary>
    /// <remarks>
    /// <b>Measured, not assumed.</b> This is the chart frame's top margin. Reading it
    /// from the sheet is what makes the chart's top edge land on a row boundary the
    /// user can see and drag; substituting the band token would put the frame's edge
    /// at a number that does not correspond to any row and would drift the moment
    /// the user changed the row.
    /// </remarks>
    public double TopPaddingHeightPt { get; }

    /// <summary>
    /// The measured height of the bottom padding row, in points (ADR-0031 D2).
    /// </summary>
    /// <remarks>
    /// The bottom counterpart to <see cref="TopPaddingHeightPt"/>, and measured from
    /// the row after the last activity row rather than assumed to match the top one.
    /// The two rows are independent: a user can resize either, and assuming they are
    /// equal would silently disagree with the sheet on whichever one they changed.
    /// </remarks>
    public double BottomPaddingHeightPt { get; }

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
    /// <param name="originTopPt">
    /// The absolute worksheet Y of the first body row's top edge, in points
    /// (ADR-0030 D3). Defaults to zero, which is correct for an EXPORT composition
    /// that builds its own origin from zero; a LIVE composition must pass the real
    /// measurement, because a default of zero is exactly the page-origin assumption
    /// ADR-0030 removed.
    /// </param>
    /// <param name="originLeftPt">
    /// The absolute worksheet X of the panel's left edge, in points
    /// (ADR-0030 D3).
    /// </param>
    /// <param name="topPaddingHeightPt">
    /// The measured top padding row height, in points (ADR-0031 D2). Defaults to
    /// zero, which is correct for an EXPORT composition: there are no worksheet rows
    /// there, so a zero margin is the honest figure rather than a missing one.
    /// </param>
    /// <param name="bottomPaddingHeightPt">
    /// The measured bottom padding row height, in points (ADR-0031 D2). Defaults to
    /// zero on the same reasoning as <paramref name="topPaddingHeightPt"/>.
    /// </param>
    /// <returns>A typed result or refusal.</returns>
    public static PanelCellGridCreationOutcome TryCreate(
        IReadOnlyList<PanelColumn>? columns,
        IReadOnlyList<double>? rowHeightsPt,
        double headerHeightPt,
        IReadOnlyList<string>? requiredColumns,
        double originTopPt = 0d,
        double originLeftPt = 0d,
        double topPaddingHeightPt = 0d,
        double bottomPaddingHeightPt = 0d)
    {
        // A padding height is a MEASUREMENT, so it is validated like every other
        // measurement here rather than defaulted when absent. Zero is a legitimate
        // value (the export case) and is therefore allowed; a negative or
        // non-finite one is not, because it would build a chart whose frame extends
        // past its own content on that side and nothing downstream would report it.
        if (!double.IsFinite(topPaddingHeightPt) || topPaddingHeightPt < 0)
        {
            return Refused(PanelCellGridRefusal.InvalidPaddingHeight);
        }

        if (!double.IsFinite(bottomPaddingHeightPt) || bottomPaddingHeightPt < 0)
        {
            return Refused(PanelCellGridRefusal.InvalidPaddingHeight);
        }

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
            // Zero is legal: it is what the host reports for a hidden row, whose space
            // the rows below share. Only a negative height is refused (ADR-0034).
            if (!double.IsFinite(rowHeight) || rowHeight < 0)
            {
                return Refused(PanelCellGridRefusal.NegativeRowHeight);
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

        return IsValidOrigin(originTopPt) && IsValidOrigin(originLeftPt)
            ? new PanelCellGridCreationOutcome(
                new PanelCellGrid(
                    [.. columns],
                    [.. rowHeightsPt],
                    headerHeightPt,
                    [.. requiredColumns ?? []],
                    originTopPt,
                    originLeftPt,
                    topPaddingHeightPt,
                    bottomPaddingHeightPt),
                null)
            : Refused(PanelCellGridRefusal.InvalidOrigin);
    }

    /// <summary>
    /// Whether a measured origin is usable: finite and non-negative.
    /// </summary>
    /// <param name="value">The measured coordinate in points.</param>
    /// <returns><see langword="true"/> when the value can be trusted as a position.</returns>
    /// <remarks>
    /// Zero is <b>valid</b>: it is a real position — the sheet's own origin — and an
    /// export composition legitimately starts there. Only a negative or non-finite
    /// measurement is refused.
    /// </remarks>
    private static bool IsValidOrigin(double value) => double.IsFinite(value) && value >= 0;

    private static PanelCellGridCreationOutcome Refused(PanelCellGridRefusal refusal) => new(null, refusal);
}
