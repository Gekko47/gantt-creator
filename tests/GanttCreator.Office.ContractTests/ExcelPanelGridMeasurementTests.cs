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
        IWorksheetProtectionGuard guard,
        Excel.ListObject table,
        Excel.ListColumn column)
        : ExcelPanelGridMeasurement(application, guard)
    {
        internal override Excel.ListObject? FindGanttTable(Excel.Sheets sheets) => table;

        internal override Excel.ListColumn? FindColumn(Excel.ListObject candidate, string name) =>
            string.Equals(name, ColumnName, StringComparison.Ordinal) ? column : null;
    }

    /// <summary>The table and the single column the fake exposes.</summary>
    private sealed record FakeTable(Excel.ListObject Table, Excel.ListColumn Column);

    /// <summary>
    /// Builds a table whose column range and body range report the supplied values,
    /// exactly as the PIA surfaces them boxed as <c>Object</c>.
    /// </summary>
    /// <param name="width">What the column's range reports as <c>Width</c>.</param>
    /// <param name="rowHeight">What the body's range reports as <c>RowHeight</c>.</param>
    /// <returns>The mocked table and its single column.</returns>
    private static FakeTable TableReporting(object? width, object? rowHeight)
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
        _ = body.SetupGet(r => r.RowHeight).Returns(rowHeight!);

        var table = new Mock<Excel.ListObject>();
        _ = table.SetupGet(t => t.DataBodyRange).Returns(body.Object);
        return new FakeTable(table.Object, column.Object);
    }

    private static Mock<IWorksheetProtectionGuard> ClearGuard()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
        return guard;
    }

    private static Mock<Excel.Application> ActiveApplication()
    {
        var workbook = new Mock<Excel.Workbook>();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        var application = new Mock<Excel.Application>();
        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }

    private static PanelGridOutcome MeasureTable(FakeTable table) =>
        MeasureTable(table, ClearGuard().Object);

    private static PanelGridOutcome MeasureTable(FakeTable table, IWorksheetProtectionGuard guard)
    {
        var measurement = new TestableMeasurement(
            ActiveApplication().Object,
            guard,
            table.Table,
            table.Column);

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
        Assert.Equal(15d, outcome.Grid.RowHeightPt);
        Assert.Equal(64d, Assert.Single(outcome.Grid.Columns).WidthPt);
    }

    /// <summary>
    /// The positive test for the validator. A body with MIXED row heights has no
    /// single height, so Excel reports <see cref="DBNull.Value"/> rather than a
    /// number. That is an absent measurement, and it must become the same typed
    /// refusal a <see langword="null"/> produces — not an
    /// <see cref="InvalidCastException"/> escaping into the render command.
    /// </summary>
    [Fact]
    public void Measure_returns_a_typed_refusal_when_the_body_reports_DBNull_for_a_mixed_row_height()
    {
        PanelGridOutcome outcome = MeasureTable(TableReporting(64d, DBNull.Value));

        Assert.False(outcome.Succeeded);
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
            ClearGuard().Object,
            table.Table,
            table.Column);

        PanelGridOutcome outcome = measurement.Measure([]);

        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    /// <summary>
    /// A protected sheet is refused by the shared guard before any measurement:
    /// ADR-0008 D4's first read-only check applies even to a read-only adapter.
    /// </summary>
    [Fact]
    public void Measure_refuses_a_protected_sheet_before_measuring()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        FakeTable table = TableReporting(64d, 15d);

        PanelGridOutcome outcome = MeasureTable(table, guard.Object);

        Assert.Equal(PanelGridRefusalReason.TargetProtected, outcome.Refusal);
    }
}
