namespace GanttCreator.Core;

/// <summary>
/// The <b>single authority</b> for the live Gantt worksheet's vertical layout and
/// its plot anchor (R4.7I slice 1 D1, ADR-0030 D4/D5).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why one authority.</b> The header row used to be worksheet row 1 by
/// convention, and three independent files each built the plot anchor's
/// <c>refersTo</c> string with a hardcoded <c>$1</c> literal: the initialiser, the
/// integrity checker, and the plot-anchor repairer — each with its own copy of the
/// A1 column conversion. ADR-0030 adds a reserved row above the table, which moves
/// the header to row 2. Editing three literals in step is exactly the arrangement
/// that made the disagreement possible: had the initialiser written <c>$O$2</c>
/// while the checker still expected <c>$O$1</c>, the integrity check would report
/// <c>PlotAnchorDisagreement</c> on every healthy workbook — a false finding
/// manufactured by the fix. One helper makes them incapable of disagreeing.
/// </para>
/// <para>
/// <b>The reserved row is one row, not two.</b> The owner asked for two rows above
/// the table, one for the year band and one for the period band. ADR-0030 D4/D5
/// implements <em>one</em> inserted row, because entity guide §4 already requires
/// the data-panel header's height to align with the period header's bottom — the
/// header row <em>is</em> the period band's row. A second inserted row would strand
/// the header between the period band and the first lane, leaving a visible empty
/// strip one header-row tall over the plot columns. The divergence is recorded in
/// ADR-0030 rather than taken silently; changing it is a guide amendment.
/// </para>
/// <para>
/// Pure and Office-free: this is arithmetic over the schema, testable without a
/// live host.
/// </para>
/// </remarks>
public static class GanttSheetLayout
{
    /// <summary>
    /// How many rows are reserved above the table's header row.
    /// </summary>
    /// <remarks>
    /// <b>Two</b>, per ADR-0031 D1: a top padding row and the row carrying the table
    /// title and the year band. The period band still needs no row of its own
    /// (ADR-0030 D5) because it shares the header row.
    /// </remarks>
    public const int ReservedRowCount = 2;

    /// <summary>
    /// The top padding row: worksheet row 1, above everything the add-in draws. It
    /// exists so the chart's top margin is a real row the user can see and align
    /// with, rather than a sub-row sliver of chart chrome.
    /// </summary>
    public const int TopPaddingRowIndex = 1;

    /// <summary>
    /// The reserved row directly above the header: it carries the table title across
    /// the panel columns and the year band across the plot columns.
    /// </summary>
    public const int ReservedRowIndex = TopPaddingRowIndex + 1;

    /// <summary>
    /// The table's header row, and therefore the row the period band sits on.
    /// </summary>
    /// <remarks>
    /// <b>Counted from the block's start, not from its last member.</b> This read
    /// <c>ReservedRowIndex + ReservedRowCount</c>, which was correct only while
    /// <see cref="ReservedRowIndex"/> sat at the <em>first</em> row of the reserved
    /// block. Adding the top padding row inside that block made the two overlap, and
    /// the sum then skipped a row: the header landed on 4 while the reserved rows
    /// were 1 and 2, stranding an empty row between the title row and the header.
    /// The block is contiguous by definition, so the header is simply the row after
    /// the block's last row — which is <see cref="ReservedRowIndex"/> itself now
    /// that it is the last member.
    /// </remarks>
    public const int HeaderRowIndex = ReservedRowIndex + 1;

    /// <summary>
    /// The worksheet row the first body row of the table occupies.
    /// </summary>
    /// <remarks>
    /// <b>Stated as an invariant rather than a literal.</b> The plot's top edge is
    /// this row's top edge (ADR-0030 D1), so a change to
    /// <see cref="ReservedRowCount"/> must move it without anyone editing a second
    /// place.
    /// </remarks>
    public const int FirstBodyRowIndex = HeaderRowIndex + 1;

    /// <summary>
    /// Gets the one-based worksheet row of the reserved title row.
    /// </summary>
    public static int TitleRowIndex => ReservedRowIndex;

    /// <summary>
    /// Gets the one-based worksheet column the table title starts in: the first
    /// AUTHORING column.
    /// </summary>
    /// <remarks>
    /// <b>Why the title is not in column A.</b> Columns A-C are
    /// <see cref="GanttColumnAccess.EngineHidden"/> bookkeeping
    /// (<c>Id</c>/<c>LaneId</c>/<c>StackIndex</c>). They are hidden but they still
    /// occupy worksheet positions, so a title written to column A would sit above
    /// columns the user cannot see and would appear to start over the middle of the
    /// table. The owner asked for D2:H2, which is exactly the visible span: the first
    /// authoring column through the last visible one.
    /// </remarks>
    public static int TitleStartColumnIndex
    {
        get
        {
            IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
            for (var i = 0; i < columns.Count; i++)
            {
                if (columns[i].Access != GanttColumnAccess.EngineHidden)
                {
                    return i + 1;
                }
            }

            throw new InvalidOperationException("The table schema has no user-visible column to title.");
        }
    }

    /// <summary>
    /// Gets the one-based worksheet column the table title ends in: the last
    /// user-visible column.
    /// </summary>
    public static int TitleEndColumnIndex
    {
        get
        {
            IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
            for (var i = columns.Count - 1; i >= 0; i--)
            {
                if (columns[i].Access != GanttColumnAccess.EngineHidden)
                {
                    return i + 1;
                }
            }

            throw new InvalidOperationException("The table schema has no user-visible column to title.");
        }
    }

    /// <summary>
    /// Gets how many worksheet columns the table title spans.
    /// </summary>
    public static int TitleColumnSpan => TitleEndColumnIndex - TitleStartColumnIndex + 1;

    /// <summary>
    /// Gets the one-based worksheet row of the bottom padding row: the row directly
    /// below the last activity row.
    /// </summary>
    /// <param name="bodyRowCount">The number of body rows in the table.</param>
    /// <returns>The bottom padding row's one-based worksheet index.</returns>
    /// <remarks>
    /// <b>Derived, not a constant.</b> The top padding row is a fixed index because
    /// it sits above the table, but this one sits below a body whose length the user
    /// controls. Publishing it as a constant would be a second source of truth that
    /// is wrong for every table except one particular length - the same defect
    /// <see cref="FirstBodyRowIndex"/> is derived to avoid.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="bodyRowCount"/> is not positive. A table with no
    /// body has no last activity row, so there is no row for this to name, and
    /// returning a plausible index would place the chart's bottom margin inside the
    /// table.
    /// </exception>
    public static int BottomPaddingRowIndex(int bodyRowCount)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(bodyRowCount, 1);
        return FirstBodyRowIndex + bodyRowCount;
    }

    /// <summary>
    /// Gets the one-based worksheet column of the plot anchor: the column one to the
    /// right of the table's last column.
    /// </summary>
    public static int PlotAnchorColumnIndex => GanttTableSchema.Default.Columns.Count + 1;

    /// <summary>
    /// Builds the plot anchor's <c>refersTo</c> string: the cell one column right of
    /// the table's last column, <b>on the header row</b>.
    /// </summary>
    /// <param name="sheetName">
    /// The Gantt worksheet's label. May contain an apostrophe, which is doubled
    /// inside the quoted reference.
    /// </param>
    /// <returns>The absolute A1 reference, for example <c>='Gantt Data'!$O$2</c>.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown when <paramref name="sheetName"/> is null, empty, or whitespace. An
    /// anchor naming no sheet is not a repairable value, so it is refused rather
    /// than written as a string that could never resolve.
    /// </exception>
    public static string BuildPlotAnchorRefersTo(string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        var escaped = sheetName.Replace("'", "''", StringComparison.Ordinal);
        return string.Create(
            System.Globalization.CultureInfo.InvariantCulture,
            $"='{escaped}'!${ToA1Column(PlotAnchorColumnIndex)}${HeaderRowIndex}");
    }

    /// <summary>
    /// Builds a one-based column/row pair's A1 cell reference, for example
    /// <c>(4, 2)</c> becomes <c>D2</c>.
    /// </summary>
    /// <param name="columnIndex">The one-based column index.</param>
    /// <param name="rowIndex">The one-based row index.</param>
    /// <returns>The A1 reference.</returns>
    /// <remarks>
    /// <b>Why the title test needs this.</b> The owner asked for the title in
    /// <c>D2:H2</c>. Writing that expectation as two string literals would pass even
    /// if the layout's own column derivation were wrong, because both the code and
    /// the test would be reading the same hand-typed letters. Deriving the reference
    /// from the same indices the layout publishes makes the test check the derivation
    /// instead of restating it.
    /// </remarks>
    public static string ToA1Cell(int columnIndex, int rowIndex)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(rowIndex, 1);
        return string.Create(
            System.Globalization.CultureInfo.InvariantCulture,
            $"{ToA1Column(columnIndex)}{rowIndex}");
    }

    /// <summary>
    /// Converts a one-based column index to its A1 letters, so <c>1</c> becomes
    /// <c>A</c> and <c>27</c> becomes <c>AA</c>.
    /// </summary>
    /// <param name="columnIndex">The one-based column index.</param>
    /// <returns>The A1 column letters.</returns>
    /// <remarks>
    /// <b>Why it is not <c>((char)('A' + index - 1))</c>.</b> That single-cast form
    /// is correct only to column 26 and emits a punctuation character from column 27
    /// onward, producing a reference that names no cell. The schema has 14 columns
    /// today, so the bug would be dormant until an approved schema ADR adds enough
    /// columns to reach <c>AA</c> — and an anchor written at that point would be
    /// wrong in a way nothing would catch.
    /// </remarks>
    public static string ToA1Column(int columnIndex)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(columnIndex, 1);

        var builder = new System.Text.StringBuilder();
        var remaining = columnIndex;
        while (remaining > 0)
        {
            var digit = (remaining - 1) % 26;
            _ = builder.Insert(0, (char)('A' + digit));
            remaining = (remaining - 1) / 26;
        }

        return builder.ToString();
    }
}
