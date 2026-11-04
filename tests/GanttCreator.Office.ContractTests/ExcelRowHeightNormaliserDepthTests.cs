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

        /// <summary>Each body-row height write, with the height that was written.</summary>
        public List<(int Row, double Height)> BodyWriteHeights { get; } = [];

        public List<(int Row, bool Hidden)> HiddenWrites { get; } = [];

        public List<int> LayoutWrites { get; } = [];

        /// <summary>
        /// How many <c>Type</c> cells the adapter read through <see cref="CellValueAt"/>.
        /// </summary>
        /// <remarks>
        /// The only caller of <see cref="CellValueAt"/> is <c>ReadTypeCellText</c>, so this
        /// count is exactly the per-body-row Type reads the caller's supplied texts are
        /// meant to remove.
        /// </remarks>
        public int TypeCellReads { get; private set; }

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
                .Callback<object>(value =>
                {
                    BodyWrites.Add(row);
                    BodyWriteHeights.Add((row, Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture)));
                });
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

        internal override object? CellValueAt(Excel.Range column, int index)
        {
            TypeCellReads++;
            return index >= 1 && index <= typeTexts.Count ? typeTexts[index - 1] : null;
        }
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

    private static RowHeightNormalisationOutcome Normalise(
        DepthNormaliser normaliser,
        IReadOnlyList<string?>? suppliedTypeTexts = null) =>
        normaliser.Normalise(
            ManagedPt,
            SplitterPt,
            SpacerPt,
            HeaderPt,
            ReservedRowPt,
            PaddingRowPt,
            AnchorRowPt,
            suppliedTypeTexts);

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

    /// <summary>
    /// A hidden row reported as a boxed boolean is skipped, with neither a height write
    /// nor a hidden-state write.
    /// </summary>
    /// <remarks>
    /// The boxed encodings still matter: reading <c>Hidden</c> is what decides the skip,
    /// so each supported encoding has its own test. The write itself is deliberately
    /// gone -- a hidden row measures 0pt, and writing the token into it un-hides and
    /// re-hides the row, which moves every <c>xlMoveAndSize</c> chart shape over it
    /// (measured 2026-11-03; the splitter band drifted a row taller each refresh).
    /// </remarks>
    [Fact]
    public void A_hidden_row_reported_as_a_boolean_is_skipped_without_a_write()
    {
        DepthNormaliser normaliser = Normaliser([30d], hiddenValues: [true]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Empty(normaliser.BodyWrites);
        Assert.Empty(normaliser.HiddenWrites);
    }

    /// <summary>A hidden row reported as the integer 1 is skipped too.</summary>
    [Fact]
    public void A_hidden_row_reported_as_an_integer_is_skipped_without_a_write()
    {
        // Range.Hidden is Object on the PIA, so a host may report 1 rather than true.
        DepthNormaliser normaliser = Normaliser([30d], hiddenValues: [1]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Empty(normaliser.BodyWrites);
        Assert.Empty(normaliser.HiddenWrites);
    }

    /// <summary>A hidden row reported as the double 1.0 is skipped too.</summary>
    [Fact]
    public void A_hidden_row_reported_as_a_double_is_skipped_without_a_write()
    {
        DepthNormaliser normaliser = Normaliser([30d], hiddenValues: [1.0d]);

        Assert.True(Normalise(normaliser).Succeeded);
        Assert.Empty(normaliser.BodyWrites);
        Assert.Empty(normaliser.HiddenWrites);
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

    /// <summary>
    /// A supplied Type text decides the row's height policy with no cell read at all.
    /// </summary>
    /// <remarks>
    /// The host's own Type cell deliberately reports <see langword="null"/> while the
    /// caller supplies <c>Splitter</c>. If the adapter ignored the supplied list the row
    /// would take the managed token (18pt); it takes the splitter token (24pt) and reads
    /// no cell, so the supplied text is both used and the only source that was read.
    /// </remarks>
    [Fact]
    public void A_supplied_type_text_decides_the_height_without_reading_the_cell()
    {
        DepthNormaliser normaliser = Normaliser([30d], typeTexts: [null]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser, ["Splitter"]);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(0, normaliser.TypeCellReads);
        Assert.Equal([(1, SplitterPt)], normaliser.BodyWriteHeights);
    }

    /// <summary>
    /// A supplied list shorter than the body falls back to the column for the tail.
    /// </summary>
    /// <remarks>
    /// A partial projection must not shift later rows onto earlier types: row 1 comes from
    /// the supplied list and row 2 from the host's Type column, and each gets its own
    /// token. Exactly one cell is read -- the one row the caller did not supply.
    /// </remarks>
    [Fact]
    public void A_supplied_list_shorter_than_the_body_reads_the_remaining_rows()
    {
        Excel.ListColumn typeColumn = Column("Type", new Mock<Excel.Range>().Object);
        DepthNormaliser normaliser = Normaliser(
            [30d, 30d],
            typeTexts: [null, "Splitter"],
            columns: [typeColumn]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser, ["Spacer"]);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(1, normaliser.TypeCellReads);
        Assert.Equal([(1, SpacerPt), (2, SplitterPt)], normaliser.BodyWriteHeights);
    }

    /// <summary>
    /// A pass that writes no body row returns the heights it measured, so the panel
    /// measurement can reuse them instead of reading every row again.
    /// </summary>
    /// <remarks>
    /// The rows are already at the managed token, so the plan writes nothing -- the warm
    /// case, which is the common refresh. The figures are the sheet's state precisely
    /// because nothing in this pass changed it.
    /// </remarks>
    [Fact]
    public void A_pass_that_writes_no_body_row_returns_the_heights_it_measured()
    {
        DepthNormaliser normaliser = Normaliser([ManagedPt]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Empty(normaliser.BodyWrites);
        Assert.Equal([ManagedPt], outcome.BodyHeightsPt);
    }

    /// <summary>
    /// A pass that writes a body row returns no heights: the sheet has changed, and the
    /// measurement is no longer of the sheet as it now is.
    /// </summary>
    /// <remarks>
    /// Returning the pre-write measurement here would make the measurement reuse data
    /// this refresh has just replaced, so the panel would lay out the old heights while
    /// the sheet displays the new ones. The empty list forces the measurement to read.
    /// </remarks>
    [Fact]
    public void A_pass_that_writes_a_body_row_returns_no_heights()
    {
        DepthNormaliser normaliser = Normaliser([30d]);

        RowHeightNormalisationOutcome outcome = Normalise(normaliser);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal([1], normaliser.BodyWrites);
        Assert.Empty(outcome.BodyHeightsPt);
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
