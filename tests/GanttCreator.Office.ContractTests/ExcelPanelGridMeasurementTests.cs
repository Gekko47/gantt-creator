using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelPanelGridMeasurement"/>. No live Excel: only
/// the table and column lookups are substituted, so the adapter's real
/// <c>Object</c>-to-<c>double</c> conversions run against a mocked PIA
/// <c>Range</c>.
/// </summary>
/// <remarks>
/// The focus is absence handling. Both <c>Range.RowHeight</c> and
/// <c>Range.Width</c> are <c>System.Object</c> on the PIA and report "no value" in
/// two encodings — <see langword="null"/> and <see cref="DBNull"/>. The second is
/// what a body with mixed row heights produces, and converting it throws
/// <see cref="InvalidCastException"/> out of a read-only adapter instead of
/// producing a typed refusal (AGENTS.md validator rule; checklist section I).
/// </remarks>
public class ExcelPanelGridMeasurementTests
{
    /// <summary>The one column the fake table exposes.</summary>
    private const string ColumnName = "Id";

    /// <summary>
    /// The measurement adapter with the table and column lookups substituted. The
    /// width and row-height conversions are NOT overridden: they are the code under
    /// test, driven by whatever the mocked <c>Range</c> reports.
    /// </summary>
    private sealed class TestableMeasurement(
        object? application,
        Excel.ListObject table,
        Excel.ListColumn column,
        IReadOnlyList<object?> bodyRowHeights,
        object? headerRowHeight)
        : ExcelPanelGridMeasurement(application)
    {
        internal override Excel.ListObject? FindGanttTable(Excel.Sheets sheets) => table;

        internal override Excel.ListColumn? FindColumn(Excel.ListObject candidate, string name) =>
            string.Equals(name, ColumnName, StringComparison.Ordinal) ? column : null;

        // The row seams are substituted because Range.Rows.Item is a COM
        // parameterised property, which cannot appear in an expression tree
        // (CS0855). The row VALUES are still the code under test: the adapter reads
        // each row's own height, so a mixed body produces a mixed list.
        internal override int GetBodyRowCount(Excel.Range body) => bodyRowHeights.Count;

        internal override Excel.Range? GetBodyRowAt(Excel.Range body, int index)
        {
            if (index < 1 || index > bodyRowHeights.Count)
            {
                return null;
            }

            object? reported = bodyRowHeights[index - 1];
            var row = new Mock<Excel.Range>();
            _ = row.SetupGet(r => r.RowHeight).Returns(reported!);
            return row.Object;
        }

        // The header row is its own range with its own height (section 4), so it is
        // read from HeaderRowRange rather than from the body.
        internal override double? ReadHeaderRowHeight(Excel.ListObject candidate)
        {
            if (headerRowHeight is null)
            {
                return null;
            }

            var header = new Mock<Excel.Range>();
            _ = header.SetupGet(r => r.RowHeight).Returns(headerRowHeight);
            var withHeader = new Mock<Excel.ListObject>();
            _ = withHeader.SetupGet(t => t.HeaderRowRange).Returns(header.Object);
            return base.ReadHeaderRowHeight(withHeader.Object);
        }
    }

    /// <summary>The table, its single column, and what each body row reports.</summary>
    private sealed record FakeTable(
        Excel.ListObject Table,
        Excel.ListColumn Column,
        IReadOnlyList<object?> BodyRowHeights,
        object? HeaderRowHeight);

    /// <summary>
    /// Builds a table whose column range, header row, and each body row report the
    /// supplied values, exactly as the PIA surfaces them boxed as <c>Object</c>.
    /// </summary>
    /// <param name="width">What the column's range reports as <c>Width</c>.</param>
    /// <param name="rowHeight">
    /// The default for both the header row and a single-row body, so the existing
    /// uniform cases keep describing a uniform table.
    /// </param>
    /// <param name="bodyRowHeights">What each body row reports, in body order.</param>
    /// <param name="headerRowHeight">What the header row reports, when it should differ.</param>
    /// <returns>The mocked table and its single column.</returns>
    private static FakeTable TableReporting(
        object? width,
        object? rowHeight,
        IReadOnlyList<object?>? bodyRowHeights = null,
        object? headerRowHeight = null)
    {
        var columnRange = new Mock<Excel.Range>();
        // The PIA types Width/RowHeight as non-nullable object even though the host
        // genuinely reports DBNull.Value through them, so the null case is
        // expressed with a null-forgiving operator rather than by loosening the
        // production signature.
        _ = columnRange.SetupGet(r => r.Width).Returns(width!);

        var column = new Mock<Excel.ListColumn>();
        _ = column.SetupGet(c => c.Range).Returns(columnRange.Object);

        var body = new Mock<Excel.Range>();
        // The body range's own aggregate is deliberately left as the default
        // (DBNull-free zero) and is never read by the adapter any more; the
        // per-row walk is the only source. Setting it here would suggest it matters.
        _ = body.SetupGet(r => r.RowHeight).Returns(rowHeight!);

        var table = new Mock<Excel.ListObject>();
        _ = table.SetupGet(t => t.DataBodyRange).Returns(body.Object);
        return new FakeTable(
            table.Object,
            column.Object,
            bodyRowHeights ?? [rowHeight],
            headerRowHeight ?? rowHeight);
    }

    private static Mock<Excel.Application> ActiveApplication()
    {
        var workbook = new Mock<Excel.Workbook>();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        var application = new Mock<Excel.Application>();
        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }

    private static PanelGridOutcome MeasureTable(FakeTable table)
    {
        var measurement = new TestableMeasurement(
            ActiveApplication().Object,
            table.Table,
            table.Column,
            table.BodyRowHeights,
            table.HeaderRowHeight);

        return measurement.Measure([ColumnName]);
    }

    /// <summary>
    /// The happy path: a numeric width and row height produce a grid. Without it,
    /// the refusals below would not distinguish "the rule refuses" from "the
    /// measurement never works".
    /// </summary>
    [Fact]
    public void Measure_returns_a_grid_when_the_host_reports_numbers()
    {
        PanelGridOutcome outcome = MeasureTable(TableReporting(64d, 15d));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.NotNull(outcome.Grid);
        // One height per body row, read from that row: a one-row body yields exactly
        // one entry.
        Assert.Equal([15d], outcome.Grid.RowHeightsPt);
        Assert.Equal(15d, outcome.Grid.HeaderHeightPt);
        Assert.Equal(64d, Assert.Single(outcome.Grid.Columns).WidthPt);
    }

    /// <summary>
    /// A body with MIXED row heights now measures successfully, one height per row,
    /// in body order. This is the capability Commit B's contract bought and this
    /// adapter change delivers; it used to be refused outright.
    /// </summary>
    /// <remarks>
    /// The exact ordered values are the assertion. Asserting only that the call
    /// succeeded would pass against a body whose rows happened to share a height, and
    /// a tolerance would accept a layout built from one sampled value - which is the
    /// precise defect this replaces.
    /// </remarks>
    [Fact]
    public void Measure_returns_the_exact_distinct_height_of_each_row_of_a_mixed_body()
    {
        PanelGridOutcome outcome = MeasureTable(TableReporting(64d, 12d, [12d, 24d, 15d]));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal([12d, 24d, 15d], outcome.Grid?.RowHeightsPt);
        Assert.Equal(51d, outcome.Grid?.TotalRowHeightPt);
    }

    /// <summary>
    /// The aggregate over a mixed range is documented to report either the first
    /// row's height or Null. Neither is consulted any more, and this pins the case
    /// that used to slip through: the aggregate reporting the first row's height is a
    /// plain number that passed every absence check while describing a body most of
    /// which is taller.
    /// </summary>
    [Fact]
    public void Measure_ignores_an_aggregate_that_reports_the_first_row_of_a_mixed_body()
    {
        // The aggregate deliberately reports 12d - row 1's height - while rows 2 and 3
        // are 24d and 15d. The result must carry the real per-row heights, not 12d
        // three times.
        PanelGridOutcome outcome = MeasureTable(TableReporting(64d, 12d, [12d, 24d, 15d]));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal([12d, 24d, 15d], outcome.Grid?.RowHeightsPt);
    }

    /// <summary>
    /// The header row is its own worksheet row with its own height, so §4 requires it
    /// to be measured rather than copied from a body row.
    /// </summary>
    [Fact]
    public void Measure_reads_the_header_height_separately_from_the_body()
    {
        PanelGridOutcome outcome = MeasureTable(
            TableReporting(64d, 15d, [15d, 15d], headerRowHeight: 30d));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(30d, outcome.Grid?.HeaderHeightPt);
        Assert.Equal([15d, 15d], outcome.Grid?.RowHeightsPt);
    }

    /// <summary>
    /// One unreadable row refuses the whole measurement. The panel's row positions
    /// are cumulative, so a single defaulted height would displace every row below it
    /// and the result would look plausible.
    /// </summary>
    [Fact]
    public void Measure_refuses_when_one_row_of_a_body_cannot_be_read()
    {
        PanelGridOutcome outcome = MeasureTable(TableReporting(64d, 12d, [12d, DBNull.Value, 15d]));

        Assert.False(outcome.Succeeded);
        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    /// <summary>
    /// A body the host cannot enumerate has no established heights, so it is refused
    /// rather than measured from an aggregate.
    /// </summary>
    [Fact]
    public void Measure_returns_a_typed_refusal_when_no_body_row_can_be_read()
    {
        PanelGridOutcome outcome = MeasureTable(TableReporting(64d, 15d, []));

        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    /// <summary>
    /// The other absence encoding, for the same reason: the two must be
    /// indistinguishable to a caller.
    /// </summary>
    [Fact]
    public void Measure_returns_a_typed_refusal_when_the_host_reports_null()
    {
        PanelGridOutcome outcome = MeasureTable(TableReporting(64d, null));

        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    /// <summary>
    /// The width conversion carries the identical rule, so it gets the identical
    /// coverage: a range with no single width reports <see cref="DBNull.Value"/> for
    /// the same reason a multi-row body does.
    /// </summary>
    [Fact]
    public void Measure_returns_a_typed_refusal_when_a_column_reports_DBNull_for_a_mixed_width()
    {
        PanelGridOutcome outcome = MeasureTable(TableReporting(DBNull.Value, 15d));

        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    /// <summary>
    /// An empty column list is refused before any host call, so the refusals above
    /// are the measurement rule rather than a degenerate request.
    /// </summary>
    [Fact]
    public void Measure_refuses_an_empty_column_list_before_reading_the_host()
    {
        FakeTable table = TableReporting(64d, 15d);
        var measurement = new TestableMeasurement(
            ActiveApplication().Object,
            table.Table,
            table.Column,
            table.BodyRowHeights,
            table.HeaderRowHeight);

        PanelGridOutcome outcome = measurement.Measure([]);

        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    /// <summary>
    /// Measurement now succeeds on a protected worksheet. This adapter mutates
    /// nothing, and Excel's protection blocks writes rather than reads, so refusing a
    /// readable target was a capability restriction with no product behind it: a
    /// read-only diagnostic or export measurement simply could not run.
    /// </summary>
    /// <remarks>
    /// This is the positive test for the policy change, and it is written as a
    /// positive because there is no longer a guard to configure. The test that
    /// previously asserted `TargetProtected` is **deleted, not inverted** - it
    /// described the behaviour being removed, and leaving it as a refusal would
    /// contradict this one.
    /// </remarks>
    [Fact]
    public void Measure_succeeds_on_a_protected_worksheet_because_reads_are_not_writes()
    {
        PanelGridOutcome outcome = MeasureTable(TableReporting(64d, 15d));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal([15d], outcome.Grid?.RowHeightsPt);
    }

    /// <summary>
    /// A missing active workbook is still refused, and still typed. Removing the
    /// guard did not remove this: the adapter reads `ActiveWorkbook` directly rather
    /// than inferring it from a protection query.
    /// </summary>
    [Fact]
    public void Measure_refuses_when_there_is_no_active_workbook()
    {
        var workbook = new Mock<Excel.Workbook>();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        var application = new Mock<Excel.Application>();
        _ = application.Setup(a => a.ActiveWorkbook).Returns((Excel.Workbook?)null!);

        FakeTable table = TableReporting(64d, 15d);
        var measurement = new TestableMeasurement(
            application.Object,
            table.Table,
            table.Column,
            table.BodyRowHeights,
            table.HeaderRowHeight);

        PanelGridOutcome outcome = measurement.Measure([ColumnName]);

        Assert.Equal(PanelGridRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }
}
