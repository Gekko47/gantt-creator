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
    /// <b>One</b>, per ADR-0030 D4: the row carrying the table title and the year
    /// band. The period band needs no row of its own (D5).
    /// </remarks>
    public const int ReservedRowCount = 1;

    /// <summary>
    /// The reserved row directly above the header: it carries the table title across
    /// the panel columns and the year band across the plot columns.
    /// </summary>
    public const int ReservedRowIndex = 1;

    /// <summary>
    /// The table's header row, and therefore the row the period band sits on.
    /// </summary>
    public const int HeaderRowIndex = ReservedRowIndex + ReservedRowCount;

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
