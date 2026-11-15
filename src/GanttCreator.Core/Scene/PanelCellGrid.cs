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

    /// <summary>
    /// The measured leading gutter width was not finite, or was negative.
    /// </summary>
    /// <remarks>
    /// Zero is legitimate and is the default: an export composition has no worksheet
    /// gutter, and a live sheet whose gutter has not been measured yet must still
    /// compose. A negative width is refused on the same rule as a negative padding
    /// row — it would place the chart frame's left edge inside the table it is meant
    /// to enclose — rather than clamped, because clamping would hide a measurement
    /// bug behind a plausible number.
    /// </remarks>
    InvalidLeadingGutter = 12,
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
        double bottomPaddingHeightPt,
        double anchorRowHeightPt,
        double leadingGutterPt,
        double? plotOriginRightPt)
    {
        Columns = columns;
        RowHeightsPt = rowHeightsPt;
        HeaderHeightPt = headerHeightPt;
        RequiredColumns = requiredColumns;
        OriginTopPt = originTopPt;
        OriginLeftPt = originLeftPt;
        TopPaddingHeightPt = topPaddingHeightPt;
        BottomPaddingHeightPt = bottomPaddingHeightPt;
        AnchorRowHeightPt = anchorRowHeightPt;
        LeadingGutterPt = leadingGutterPt;
        PlotOriginRightPt = plotOriginRightPt;

        // The cumulative tops are derived ONCE, here, so a lane anchor and a panel
        // row cannot disagree about where a body row starts. Entry n is
        // OriginTopPt plus the heights of every row above n, which is the row's
        // absolute worksheet top edge; a zero-height row therefore contributes no
        // space and the rows below it share its top (ADR-0034).
        double[] tops = new double[rowHeightsPt.Count];
        double running = originTopPt;
        for (int index = 0; index < rowHeightsPt.Count; index++)
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

    /// <summary>
    /// The measured height of the reserved anchor row, in points (ADR-0038 D1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The row immediately below the last body row, which the plot-spanning shapes
    /// paint through so Excel resolves their bottom cell anchor below the insert
    /// point. It is measured rather than assumed for the same reason
    /// <see cref="BottomPaddingHeightPt"/> is: the scene needs the real row height to
    /// place the closing line, and a token substituted here would put the line at a
    /// Y that corresponds to no row the moment a user resized it.
    /// </para>
    /// <para>
    /// <b>Zero is legal</b> and means "no anchor row", which is the export case: an
    /// export has no worksheet rows to anchor to, so it paints no overhang and emits
    /// no closing line below the plot.
    /// </para>
    /// </remarks>
    public double AnchorRowHeightPt { get; }

    /// <summary>
    /// The measured width of the add-in's leading gutter column, in points (R5.18k).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Column A is not a column of <c>tblGanttData</c>
    /// (<see cref="GanttSheetLayout.LeadingGutterColumnIndex"/>); it is the spacer the
    /// chart frame encloses on the left, so the frame has the same margin on all four
    /// sides instead of sitting flush against the first table column.
    /// </para>
    /// <para>
    /// <b>It is a separate measurement from <see cref="OriginLeftPt"/> and must stay
    /// one.</b> The origin is the table's own left edge and is where the live title
    /// textbox is placed; folding the gutter into it would move the title onto the
    /// gutter. The panel rectangle is what subtracts the gutter, because the panel is
    /// the thing the frame is drawn around.
    /// </para>
    /// </remarks>
    public double LeadingGutterPt { get; }

    /// <summary>
    /// The panel rectangle's left edge: the table's own left edge moved left by the
    /// leading gutter (R5.18k).
    /// </summary>
    /// <remarks>
    /// This is the value the FRAME is drawn around, so it is published rather than
    /// recomputed at the call site: the subtraction is the whole effect of the gutter,
    /// and a second copy of it in the builder is where an off-by-one-gutter would hide.
    /// <see cref="OriginLeftPt"/> is deliberately unchanged — it is where the live title
    /// textbox sits, and a title moved onto the gutter would be the visible defect this
    /// property exists to avoid.
    /// </remarks>
    public double PanelLeftPt => OriginLeftPt - LeadingGutterPt;

    /// <summary>
    /// The panel rectangle's right edge: the table's origin plus the summed column
    /// widths, which is the right edge of the table's last column.
    /// </summary>
    /// <remarks>
    /// <b>Deliberately not <see cref="PanelLeftPt"/> plus the width.</b> The gutter widens
    /// the panel leftwards only, so the right edge stays measured from the origin. Two
    /// callers need that same number and must not recompute it: the frame's right rule,
    /// and the live plot's left edge, which begins exactly here — on the column boundary,
    /// which is the alignment the owner ruled (2026-11-14). The defect this property
    /// closes is that the plot was placed at the summed widths alone, which on a sheet
    /// whose table starts in column B put it one measured origin — a gutter — inside the
    /// table's last column.
    /// </remarks>
    public double PanelRightPt => PlotOriginRightPt ?? (OriginLeftPt + TotalWidthPt);

    /// <summary>
    /// The MEASURED right edge of the panel, when the host supplied one.
    /// </summary>
    /// <remarks>
    /// On the live sheet this is the PLOT ANCHOR COLUMN's left edge, not the sum of the
    /// visible columns' widths: the table carries twelve hidden engine columns after the
    /// visible ones, so the sum stops at the last visible column and the plot began one
    /// hidden range too early (owner report 2026-11-15: the chart started in column G
    /// while the plot anchor is in column S). Null on every composition that does not
    /// measure one, which keeps export geometry byte-identical.
    /// </remarks>
    public double? PlotOriginRightPt { get; }

    /// <summary>
    /// The chart's total bottom margin: the anchor row plus the padding row
    /// (ADR-0038 D1/D4).
    /// </summary>
    /// <remarks>
    /// <b>Summed rather than either row alone.</b> The strip below the body is two
    /// rows now, and the chart frame still closes at the bottom of that whole strip
    /// (ADR-0031 D2). Reading the padding row alone would close the frame
    /// <c>AnchorRowHeightPt</c> too high, inside the anchor row itself.
    /// </remarks>
    public double TotalBottomMarginHeightPt => AnchorRowHeightPt + BottomPaddingHeightPt;

    /// <summary>The total body height in points, the sum of the row heights.</summary>
    public double TotalRowHeightPt
    {
        get
        {
            double total = 0;
            foreach (double height in RowHeightsPt)
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
    /// <param name="anchorRowHeightPt">
    /// The measured reserved anchor row height, in points (ADR-0038 D1) -- the row
    /// directly below the body that the plot-spanning shapes paint through. Defaults
    /// to zero on the same reasoning: an EXPORT composition has no worksheet rows to
    /// anchor to, so it paints no overhang and emits no closing line.
    /// </param>
    /// <param name="allowEmptyRows">Allows an intentionally empty live chart grid; defaults to the established no-rows refusal.</param>
    /// <param name="leadingGutterPt">
    /// The measured width of the add-in's leading gutter column, in points (R5.18k).
    /// Defaults to zero, which is correct for an EXPORT composition and for a live
    /// sheet whose gutter has not been measured: the frame then starts at the table's
    /// own left edge, which is the behaviour every composition had before the gutter
    /// existed.
    /// </param>
    /// <param name="plotOriginRightPt">
    /// The MEASURED right edge of the panel, in points, when the caller measured one - the
    /// plot anchor column's left edge on the live sheet. Null keeps the summed-width edge.
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
        double bottomPaddingHeightPt = 0d,
        double anchorRowHeightPt = 0d,
        bool allowEmptyRows = false,
        double leadingGutterPt = 0d,
        double? plotOriginRightPt = null)
    {
        // Validated on the same rule as the padding rows and with its own reason: a
        // negative gutter would invert the frame, and folding it into the padding
        // refusal would tell a reader to go and look at the wrong rows.
        if (!double.IsFinite(leadingGutterPt) || leadingGutterPt < 0)
        {
            return Refused(PanelCellGridRefusal.InvalidLeadingGutter);
        }
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

        // The anchor row is validated on the SAME rule as the padding rows, and with
        // its own check rather than folded into one of theirs. Folding it in would
        // have made a negative anchor row indistinguishable from a negative padding
        // row in the refusal a user sees, and the two are fixed by different actions.
        if (!double.IsFinite(anchorRowHeightPt) || anchorRowHeightPt < 0)
        {
            return Refused(PanelCellGridRefusal.InvalidPaddingHeight);
        }

        // A supplied right edge is a POSITION, so it is validated like an origin:
        // finite and positive. Zero would put the plot on the sheet's own edge, which is
        // the defect this measurement exists to remove.
        if (plotOriginRightPt is { } suppliedRight
            && (!double.IsFinite(suppliedRight) || suppliedRight <= 0))
        {
            return Refused(PanelCellGridRefusal.InvalidOrigin);
        }

        if (columns is null)
        {
            return Refused(PanelCellGridRefusal.NullRequest);
        }

        if (columns.Count == 0)
        {
            return Refused(PanelCellGridRefusal.NoColumns);
        }

        if (rowHeightsPt is null || (!allowEmptyRows && rowHeightsPt.Count == 0))
        {
            return Refused(PanelCellGridRefusal.NoRows);
        }

        if (!double.IsFinite(headerHeightPt) || headerHeightPt <= 0)
        {
            return Refused(PanelCellGridRefusal.NonPositiveHeaderHeight);
        }

        foreach (double rowHeight in rowHeightsPt)
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
        foreach (string required in requiredColumns ?? [])
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
                    bottomPaddingHeightPt,
                    anchorRowHeightPt,
                    leadingGutterPt,
                    plotOriginRightPt),
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
