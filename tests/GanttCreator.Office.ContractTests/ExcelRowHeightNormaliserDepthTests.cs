using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Depth tests for <see cref="ExcelRowHeightNormaliser"/>: the table lookup, the
/// Type-column resolution, the boxed <c>Hidden</c> decoding, and the refusal
/// arms the existing harness substitutes.
/// </summary>
/// <remarks>
/// <see cref="ExcelRowHeightNormaliserTests"/> substitutes <c>TryFindTable</c>,
/// <c>TryGetTypeColumn</c>, <c>ReadTypeCellText</c> and <c>IsLayoutRowEmpty</c> so
/// its assertions are about the heights that reach the host. This file leaves
/// those four running and substitutes only the range reads, so the searches and
/// the boxed-value decoding are exercised. The reserved-row resolution internals
/// stay out of reach without the worksheet <c>Rows</c> plumbing; the tagged
/// live-Office test covers them.
/// </remarks>
public sealed class ExcelRowHeightNormaliserDepthTests
{
    private const double ManagedPt = 18;
    private const double SplitterPt = 24;
    private const double SpacerPt = 6;
    private const double HeaderPt = 16;
    private const double ReservedRowPt = 18;
    private const double PaddingRowPt = 22;
    private const double AnchorRowPt = 0.25;

    /// <summary>
    /// The normaliser with only the range reads substituted, so the table lookup,
    /// the Type-column resolution and the boxed-value decoding all run.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="rowHeights">Each body row's <c>RowHeight</c> as the host reports it.</param>
    /// <param name="hiddenValues">Each body row's <c>Hidden</c> as the host reports it.</param>
    /// <param name="typeTexts">Each body row's Type cell value as the host reports it.</param>
    /// <param name="columns">The table's list columns, in order.</param>
    /// <param name="rowsInTable">Whether the table reports found: false models a missing table.</param>
    private sealed class DepthNormaliser(
        object? application,
        IWorksheetProtectionGuard guard,
        IReadOnlyList<object?> rowHeights,
        IReadOnlyList<object?> hiddenValues,
        IReadOnlyList<object?> typeTexts,
        IReadOnlyList<Excel.ListColumn?> columns) : ExcelRowHeightNormaliser(application, guard)
    {
        private readonly Excel.Range _body = new Mock<Excel.Range>().Object;

        public List<int> BodyWrites { get; } = [];

        public List<(int Row, bool Hidden)> HiddenWrites { get; } = [];

        public List<int> LayoutWrites { get; } = [];

        internal override Excel.Range? GetTableBody(Excel.ListObject table) => _body;

        internal override int GetBodyRowCount(Excel.Range body) => rowHeights.Count;

        internal override Excel.Range? GetBodyRowAt(Excel.Range body, int index)
        {
            if (index < 1 || index > rowHeights.Count)
            {
                return null;
            }

            int row = index;
            Mock<Excel.Range> mock = new();
            _ = mock.SetupGet(r => r.RowHeight).Returns(rowHeights[index - 1]!);
            _ = mock.SetupSet(r => r.RowHeight = It.IsAny<object>())
                .Callback<object>(_ => BodyWrites.Add(row));
            _ = mock.SetupGet(r => r.Hidden).Returns(hiddenValues[index - 1]!);
            _ = mock.SetupSet(r => r.Hidden = It.IsAny<object>())
                .Callback<object>(value => HiddenWrites.Add(
                    (row, Convert.ToBoolean(value, System.Globalization.CultureInfo.InvariantCulture))));
            return mock.Object;
        }

        internal override int GetRangeRow(Excel.Range range) => GanttSheetLayout.FirstBodyRowIndex;

        internal override bool? IsLayoutRowEmpty(Excel.Worksheet source, int rowIndex) => true;

        internal override object? GetRangeValue2(Excel.Range range) => null;

        internal override Excel.Range? GetLayoutRow(Excel.Worksheet source, int rowIndex)
        {
            // Deliberately reports a height that never equals a token, so the layout
            // rows always write. The tests here assert on BodyWrites, not on the
            // total, so a constant here keeps the count assertions unambiguous.
            Mock<Excel.Range> mock = new();
            _ = mock.SetupGet(r => r.RowHeight).Returns(999d);
            _ = mock.SetupSet(r => r.RowHeight = It.IsAny<object>())
                .Callback<object>(_ => LayoutWrites.Add(rowIndex));
            return mock.Object;
        }

        internal override Excel.ListColumn ColumnAt(Excel.ListColumns columns1, int index) =>
            columns[index - 1]!;

        internal override object? CellValueAt(Excel.Range column, int index) =>
            index >= 1 && index <= typeTexts.Count ? typeTexts[index - 1] : null;
    }

    private static Mock<IWorksheetProtectionGuard> ClearGuard()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
        return guard;
    }

    /// <summary>
    /// A sheet collection that enumerates the supplied table through one sheet.
    /// </summary>
    /// <remarks>
    /// The PIA collection is a COM collection rather than an enumerable, so
    /// <c>GetEnumerator()</c> has to be stubbed explicitly -- an unconfigured
    /// collection enumerates as null and the adapter's lookup throws rather than
    /// reporting "not found".
    /// </remarks>
    private static Excel.Sheets SheetsContaining(Excel.ListObject table)
    {
        Mock<Excel.ListObjects> listObjects = new();
        List<Excel.ListObject> tables = [table];
        _ = listObjects.Setup(c => c.GetEnumerator())
            .Returns(() => (System.Collections.IEnumerator)tables.GetEnumerator());
        Mock<Excel.Worksheet> sheet = new();
        _ = sheet.SetupGet(s => s.ListObjects).Returns(listObjects.Object);
        Mock<Excel.Sheets> sheets = new();
        List<object> entries = [sheet.Object];
        _ = sheets.Setup(s => s.GetEnumerator())
            .Returns(() => (System.Collections.IEnumerator)entries.GetEnumerator());
        return sheets.Object;
    }

    private static Excel.Sheets EmptySheets()
    {
        Mock<Excel.Sheets> sheets = new();
        _ = sheets.Setup(s => s.GetEnumerator())
            .Returns(() => (System.Collections.IEnumerator)new List<object>().GetEnumerator());
        return sheets.Object;
    }

    private static Excel.Application ApplicationWith(Excel.Sheets sheets)
    {
        Mock<Excel.Workbook> workbook = new();
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets);
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application.Object;
    }

    private static Excel.ListColumn Column(string name, Excel.Range? body = null)
    {
        Mock<Excel.ListColumn> column = new();
        _ = column.SetupGet(c => c.Name).Returns(name);
        if (body is not null)
        {
            _ = column.SetupGet(c => c.DataBodyRange).Returns(body);
        }

        return column.Object;
    }

    private static Excel.ListObject Table(IReadOnlyList<Excel.ListColumn?> columns)
    {
        Mock<Excel.ListColumns> collection = new();
        _ = collection.SetupGet(c => c.Count).Returns(columns.Count);
        Mock<Excel.ListObject> table = new();
        // The schema name, so the sheet's own table lookup finds it: an unnamed
        // mock matches nothing and every Normalise would refuse TableMissing.
        _ = table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
        _ = table.SetupGet(t => t.ListColumns).Returns(collection.Object);
        return table.Object;
    }

    private static DepthNormaliser Normaliser(
        IReadOnlyList<object?> rowHeights,
        IReadOnlyList<object?>? hiddenValues = null,
        IReadOnlyList<object?>? typeTexts = null,
        IReadOnlyList<Excel.ListColumn?>? columns = null,
        bool tableFound = true)
    {
        Excel.ListObject table = Table(columns ?? []);
        Excel.Sheets sheets = tableFound ? SheetsContaining(table) : EmptySheets();
        return new DepthNormaliser(
            ApplicationWith(sheets),
            ClearGuard().Object,
            rowHeights,
            hiddenValues ?? [.. rowHeights.Select(static _ => (object?)false)],
            typeTexts ?? [.. rowHeights.Select(static _ => (object?)null)],
            columns ?? []);
    }

    private static RowHeightNormalisationOutcome Normalise(DepthNormaliser normaliser) =>
        normaliser.Normalise(
            ManagedPt,
            SplitterPt,
            SpacerPt,
            HeaderPt,
            ReservedRowPt,
            PaddingRowPt,
            AnchorRowPt);

    [Fact]
    public void Normalise_refuses_table_missing_when_no_sheet_carries_the_gantt_table()
    {
        DepthNormaliser normaliser = Normaliser([30d], tableFound: false);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.Equal(RowHeightNormalisationRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(normaliser.BodyWrites);
        Assert.Empty(normaliser.LayoutWrites);
    }

    [Fact]
    public void A_body_row_the_host_cannot_measure_refuses_the_whole_pass()
    {
        // DBNull is how an Object-typed Excel member reports "no single value".
        // Converting it would throw; the adapter must refuse instead.
        DepthNormaliser normaliser = Normaliser([DBNull.Value]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.Equal(RowHeightNormalisationRefusalReason.InvalidMeasurement, outcome.Refusal);
        Assert.Empty(normaliser.BodyWrites);
    }

    [Fact]
    public void A_body_row_the_host_reports_as_absent_refuses_the_whole_pass()
    {
        DepthNormaliser normaliser = Normaliser([null]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.Equal(RowHeightNormalisationRefusalReason.InvalidMeasurement, outcome.Refusal);
        Assert.Empty(normaliser.BodyWrites);
    }

    [Fact]
    public void A_hidden_row_reported_as_a_boolean_is_restored_hidden_after_its_height_write()
    {
        DepthNormaliser normaliser = Normaliser([30d], hiddenValues: [true]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal([1], normaliser.BodyWrites);
        Assert.Equal([(1, true)], normaliser.HiddenWrites);
    }

    [Fact]
    public void A_hidden_row_reported_as_an_integer_is_restored_hidden()
    {
        // Range.Hidden is Object on the PIA, so a host may report 1 rather than true.
        DepthNormaliser normaliser = Normaliser([30d], hiddenValues: [1]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal([(1, true)], normaliser.HiddenWrites);
    }

    [Fact]
    public void A_hidden_row_reported_as_a_double_is_restored_hidden()
    {
        DepthNormaliser normaliser = Normaliser([30d], hiddenValues: [1.0d]);

        Assert.True(Normalise(normaliser).Succeeded);
        Assert.Equal([(1, true)], normaliser.HiddenWrites);
    }

    [Fact]
    public void A_hidden_state_the_host_reports_as_unrecognised_leaves_the_row_visible()
    {
        // An unrecognised encoding reads as visible, so the row is left un-hidden
        // rather than re-hidden on a guess.
        DepthNormaliser normaliser = Normaliser([30d], hiddenValues: ["hidden"]);

        Assert.True(Normalise(normaliser).Succeeded);
        Assert.Empty(normaliser.HiddenWrites);
    }

    [Fact]
    public void A_zero_hidden_state_leaves_the_row_visible()
    {
        DepthNormaliser normaliser = Normaliser([30d], hiddenValues: [0]);

        Assert.True(Normalise(normaliser).Succeeded);
        Assert.Empty(normaliser.HiddenWrites);
    }

    [Fact]
    public void The_type_column_is_located_by_its_schema_name_behind_other_columns()
    {
        // The Duration column comes first, so the search has to skip it. A splitter
        // row then follows the splitter token rather than the managed height.
        Excel.ListColumn typeColumn = Column("Type", new Mock<Excel.Range>().Object);
        DepthNormaliser normaliser = Normaliser(
            [30d],
            typeTexts: ["Splitter"],
            columns: [Column("Duration"), typeColumn]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal([1], normaliser.BodyWrites);
    }

    [Fact]
    public void TryGetTypeColumn_returns_false_when_the_table_reports_no_column_collection()
    {
        DepthNormaliser normaliser = Normaliser([30d]);
        var table = new Mock<Excel.ListObject>().Object;

        bool found = normaliser.TryGetTypeColumn(table, out Excel.Range? typeColumn);

        Assert.False(found);
        Assert.Null(typeColumn);
    }

    [Fact]
    public void TryGetTypeColumn_returns_false_when_the_type_column_has_no_body_range()
    {
        DepthNormaliser normaliser = Normaliser([30d], columns: [Column("Type", body: null)]);

        bool found = normaliser.TryGetTypeColumn(Table([Column("Type", body: null)]), out Excel.Range? typeColumn);

        Assert.False(found);
        Assert.Null(typeColumn);
    }

    [Fact]
    public void ReadTypeCellText_returns_null_for_a_cell_that_is_not_text()
    {
        // A date- or number-formatted Type cell arrives as a boxed number, which is
        // not a type name, so the row is treated as managed rather than misread.
        DepthNormaliser normaliser = Normaliser([30d], typeTexts: [42d]);

        Assert.Null(normaliser.ReadTypeCellText(new Mock<Excel.Range>().Object, 1));
    }

    [Fact]
    public void The_body_row_count_is_zero_when_the_table_has_no_body()
    {
        DepthNormaliser normaliser = Normaliser([]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        // No body means no padding or anchor row was resolved, so only the three
        // unconditional layout rows are written.
        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(3, outcome.RowsWritten);
        Assert.Empty(normaliser.BodyWrites);
    }
}
