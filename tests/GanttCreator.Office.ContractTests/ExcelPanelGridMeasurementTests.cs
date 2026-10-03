using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
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
        object? headerRowHeight,
        double originTopPt,
        double originLeftPt,
        double topPaddingHeightPt = 18.0,
        double bottomPaddingHeightPt = 27.5,
        double anchorRowHeightPt = 0.25)
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

        // ADR-0030 D3: the absolute origin is substituted here too, because it is
        // read from the body's Top/Left. The VALUES are still asserted on, so a test
        // proves the adapter passed the measurement through rather than defaulting
        // it to zero.
        internal override double? ReadOriginTop(Excel.ListObject candidate) => originTopPt;

        internal override double? ReadOriginLeft(Excel.ListObject candidate) => originLeftPt;

        // ADR-0031 D2: the padding-row heights are substituted for the same reason the
        // origin is — they are read through COM, not computed. Distinct sentinel
        // values rather than one number, because "the bottom row matches the top row"
        // is exactly the assumption a shared fixture value could not detect.
        internal override double? ReadTopPaddingHeight(Excel.ListObject candidate) =>
            topPaddingHeightPt;

        internal override double? ReadBottomPaddingHeight(Excel.ListObject candidate) =>
            VerifyBottomPaddingRow
                ? base.ReadBottomPaddingHeight(candidate)
                : bottomPaddingHeightPt;

        // ADR-0038 D1: the anchor row is read and VERIFIED through the same seam on the
        // same terms as the padding row, so its absence is what makes the measurements
        // below refuse. Stubbed by default for the height-focused tests; the
        // verification tests leave this on so the real path runs.
        internal override double? ReadAnchorHeight(Excel.ListObject candidate) =>
            VerifyAnchorRow
                ? base.ReadAnchorHeight(candidate)
                : anchorRowHeightPt;

    /// <summary>
    /// Whether the REAL verified bottom-padding path runs rather than the seeded stub.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Off by default so the existing height tests keep their focus, and on for the
    /// tests that exist to cover ADR-0035 D2. Without this the new verification in
    /// <see cref="ExcelPanelGridMeasurement"/> would have no test at all: the double
    /// stubs the method it lives behind, so every measurement test would pass while
    /// the code never ran.
    /// </para>
    /// </remarks>
    internal bool VerifyBottomPaddingRow { get; set; }

    /// <summary>
    /// Whether the REAL verified anchor-row path runs rather than the seeded stub.
    /// </summary>
    /// <remarks>
    /// <b>Off by default</b> for the same reason as the padding row's flag: the double
    /// stubs the method it lives behind, so the real ADR-0038 D6 verification would have
    /// no test at all while every height test still passed.
    /// </remarks>
    internal bool VerifyAnchorRow { get; set; }

    /// <summary>The body's first worksheet row, as the host reports it.</summary>
    internal int BodyFirstRow { get; set; } = 4;

    /// <summary>Whether the reserved padding row reads as empty; null means unverified.</summary>
    internal bool? PaddingRowEmpty { get; set; } = true;

    internal override int GetRangeRow(Excel.Range range)
        {
            ArgumentNullException.ThrowIfNull(range);
            return BodyFirstRow;
        }

    /// <summary>The height a resolved layout row reads back at, by row index.</summary>
    internal Dictionary<int, double> LayoutRowHeights { get; } = [];

    internal override double? ReadRowHeightAt(Excel.ListObject candidate, int rowIndex) =>
            LayoutRowHeights.TryGetValue(rowIndex, out double height) ? height : null;

    internal override bool? IsLayoutRowEmpty(Excel.ListObject candidate, int rowIndex) => PaddingRowEmpty;
    }

    /// <summary>The table, its single column, and what each body row reports.</summary>
    private sealed record FakeTable(
        Excel.ListObject Table,
        Excel.ListColumn Column,
        IReadOnlyList<object?> BodyRowHeights,
        object? HeaderRowHeight,
        double OriginTopPt,
        double OriginLeftPt);

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
        object? headerRowHeight = null,
        double originTopPt = 0d,
        double originLeftPt = 0d)
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
            headerRowHeight ?? rowHeight,
            originTopPt,
            originLeftPt);
    }

    private static Mock<Excel.Application> ActiveApplication()
    {
        var workbook = new Mock<Excel.Workbook>();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        var application = new Mock<Excel.Application>();
        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }

    /// <summary>
    /// The measured absolute origin reaches the grid (ADR-0030 D3).
    /// </summary>
    /// <remarks>
    /// <b>This is the assertion that matters for the live chart.</b> Before ADR-0030
    /// the plot's top came from the size preset's page coordinates, so the scene had
    /// no idea where the body began and a lane could not land on its own row. The
    /// values are deliberately non-zero and asymmetric, so a grid that defaulted to
    /// zero - the page-origin assumption - cannot pass.
    /// </remarks>
    [Fact]
    public void The_measured_origin_reaches_the_grid()
    {
        FakeTable table = TableReporting(64d, 15d, originTopPt: 64.5, originLeftPt: 8.25);

        PanelGridOutcome outcome = MeasureTable(table);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(64.5, outcome.Grid!.OriginTopPt);
        Assert.Equal(8.25, outcome.Grid.OriginLeftPt);
    }

    /// <summary>
    /// The origin is the first BODY row's top, not the header row's.
    /// </summary>
    /// <remarks>
    /// Reading the header's top would place the plot a whole header-row too high -
    /// the same off-by-one-row class ADR-0030 corrects elsewhere. Stated because the
    /// value is what D1 hands the plot, and a header-row-high plot looks plausible
    /// rather than obviously broken.
    /// </remarks>
    [Fact]
    public void The_origin_is_measured_from_the_body_not_the_header()
    {
        FakeTable table = TableReporting(64d, 15d, headerRowHeight: 40d, originTopPt: 58d);

        PanelGridOutcome outcome = MeasureTable(table);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        // The header is 40pt tall, so a header-derived origin would be 40, not 58.
        Assert.Equal(58d, outcome.Grid!.OriginTopPt);
    }

    private static PanelGridOutcome MeasureTable(FakeTable table)
    {
        var measurement = new TestableMeasurement(
            ActiveApplication().Object,
            table.Table,
            table.Column,
            table.BodyRowHeights,
            table.HeaderRowHeight,
            table.OriginTopPt,
            table.OriginLeftPt);

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
            table.HeaderRowHeight,
            table.OriginTopPt,
            table.OriginLeftPt);

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
    /// A bottom padding row that cannot be VERIFIED refuses the whole measurement
    /// (ADR-0035 D2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the positive test for the validator.</b> The adapter used to read
    /// <c>BottomPaddingRowIndex(rowCount)</c> -- derived from the body length on the
    /// assumption the table sits where the layout authority says -- and reported that
    /// height to Core as the chart's bottom margin. A moved table therefore made an
    /// arbitrary user row's height become the chart's margin.
    /// </para>
    /// <para>
    /// Both refusals here are the same reporting decision: an unverifiable margin is
    /// absent, and the caller already maps absent to the typed
    /// <c>InvalidMeasurement</c> refusal. Guessing a margin would reproduce the
    /// original defect with a number attached.
    /// </para>
    /// </remarks>
    [Theory]
    [InlineData(false, 4)]   // occupied: the user typed in the row below the table
    [InlineData(true, 20)]   // not contiguous: the table has moved from where it was
    public void An_unverifiable_bottom_padding_row_refuses_the_measurement(bool isEmpty, int bodyFirstRow)
    {
        FakeTable table = TableReporting(64d, 15d);
        var measurement = new TestableMeasurement(
            ActiveApplication().Object,
            table.Table,
            table.Column,
            table.BodyRowHeights,
            table.HeaderRowHeight,
            table.OriginTopPt,
            table.OriginLeftPt)
        {
            VerifyBottomPaddingRow = true,
            BodyFirstRow = bodyFirstRow,
            PaddingRowEmpty = isEmpty,
        };

        PanelGridOutcome outcome = measurement.Measure([ColumnName]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    /// <summary>
    /// A VERIFIED padding row still measures, so the new check refuses only what it
    /// cannot prove.
    /// </summary>
    /// <remarks>
    /// The mirror of the refusal test. A validator that refused unconditionally would
    /// pass it, so the positive outcome has to be pinned separately or the guard
    /// would be indistinguishable from breaking the feature.
    /// </remarks>
    [Fact]
    public void A_verified_bottom_padding_row_still_measures()
    {
        FakeTable table = TableReporting(64d, 15d);

        // One body row starting at 4, so the ANCHOR row is 5 and the padding row is 6
        // (ADR-0038 D1). The verified path resolves those rows and then READS their
        // heights through this seam, which the plain fake cannot supply -- so they are
        // seeded here. Without them the reads are absent and the refusal branch
        // answers, which would make this test assert the opposite of what it names.
        var measurement = new TestableMeasurement(
            ActiveApplication().Object,
            table.Table,
            table.Column,
            table.BodyRowHeights,
            table.HeaderRowHeight,
            table.OriginTopPt,
            table.OriginLeftPt)
        {
            VerifyBottomPaddingRow = true,
            VerifyAnchorRow = true,
            BodyFirstRow = 4,
            PaddingRowEmpty = true,
        };
        measurement.LayoutRowHeights[GanttSheetLayout.AnchorRowIndex(1)] = 0.25d;
        measurement.LayoutRowHeights[GanttSheetLayout.BottomPaddingRowIndex(1)] = 27.5d;

        PanelGridOutcome outcome = measurement.Measure([ColumnName]);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        // The resolved margin is the value read off the reserved row, and the body
        // row is still the measured one -- so the check changed nothing else.
        Assert.Equal([15d], outcome.Grid?.RowHeightsPt);
        Assert.Equal(27.5d, outcome.Grid?.BottomPaddingHeightPt);

        // ADR-0038 D1: the anchor row is measured SEPARATELY, and the chart's bottom
        // margin is the SUM of the two rows so the frame still closes at the bottom of
        // the whole strip (ADR-0031 D2).
        Assert.Equal(0.25d, outcome.Grid?.AnchorRowHeightPt);
        Assert.Equal(27.75d, outcome.Grid?.TotalBottomMarginHeightPt);
    }

    /// <summary>
    /// An anchor row that cannot be VERIFIED refuses the whole measurement
    /// (ADR-0038 D6).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the positive test for the new validator</b>, and the mirror of the
    /// padding row's: without it, an adapter that resolved the anchor row with no check
    /// at all would pass every other test in this file while the row-height normaliser
    /// resized whatever user row happened to sit below the table.
    /// </para>
    /// <para>
    /// The second case is the non-contiguous one, and it is separate on purpose: it is
    /// the state a table that has MOVED produces, which is the defect ADR-0035 D2 was
    /// raised for.
    /// </para>
    /// </remarks>
    [Theory]
    [InlineData(false, 4)]   // occupied: the user typed in the anchor row
    [InlineData(true, 20)]   // not contiguous: the table has moved from where it was
    public void An_unverifiable_anchor_row_refuses_the_measurement(bool isEmpty, int bodyFirstRow)
    {
        FakeTable table = TableReporting(64d, 15d);
        var measurement = new TestableMeasurement(
            ActiveApplication().Object,
            table.Table,
            table.Column,
            table.BodyRowHeights,
            table.HeaderRowHeight,
            table.OriginTopPt,
            table.OriginLeftPt)
        {
            VerifyBottomPaddingRow = true,
            VerifyAnchorRow = true,
            BodyFirstRow = bodyFirstRow,
            PaddingRowEmpty = isEmpty,
        };

        PanelGridOutcome outcome = measurement.Measure([ColumnName]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
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
            table.HeaderRowHeight,
            table.OriginTopPt,
            table.OriginLeftPt);

        PanelGridOutcome outcome = measurement.Measure([ColumnName]);

        Assert.Equal(PanelGridRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }
}
