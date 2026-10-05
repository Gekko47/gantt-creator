using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>Contract tests for the R2.8 row insertion adapter.</summary>
public class GanttRowInserterTests
{
    /// <summary>
    /// Which of the row-creation seams the host refuses, so the guards around them have
    /// positive tests. All off by default.
    /// </summary>
    /// <remarks>
    /// These three calls are guarded on the same terms as <c>WriteRow</c> and were
    /// unguarded while it was not, so each needs its own case: a guard tested only
    /// through one of its three sites would pass with the other two removed.
    /// </remarks>
    /// <param name="graph">The fake to arm.</param>
    /// <param name="seam">Which seam refuses.</param>
    private static void RefuseRowCreationSeam(Graph graph, string seam) =>
        graph.Refusal = seam switch
        {
            "AddRow" => Graph.HostRefusal.AddRow,
            "GetListRowAt" => Graph.HostRefusal.GetListRowAt,
            "SetRowHeight" => Graph.HostRefusal.SetRowHeight,
            _ => Graph.HostRefusal.None,
        };

    private sealed class StubTypeOptionsMaterialiser : ITypeOptionsMaterialiser
    {
        public TypeOptionsMaterialiseOutcome Materialise() => TypeOptionsMaterialiseOutcome.Ok();

        public TypeOptionsMaterialiseOutcome EnsureCurrent() => TypeOptionsMaterialiseOutcome.Ok();
    }

    /// <summary>
    /// A successful row-height normalisation, with the targets recorded.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The inserter now normalises the reserved rows below the body after every insert
    /// (ADR-0038 D1), and these tests fake the COM seams on the INSERTER only. Left
    /// unstubbed, the real <see cref="ExcelRowHeightNormaliser"/> would run against
    /// this file's bare <see cref="Moq"/> workbook and throw out of
    /// <c>Insert</c> -- which is not what these tests are about and not a behaviour the
    /// inserter may have: a helper on the Ribbon path must never escape into the
    /// callback (ADR-0008).
    /// </para>
    /// <para>
    /// Recording the targets is what lets
    /// <c>The_appended_row_normalises_the_reserved_rows_below_the_body</c> assert the
    /// anchor and padding tokens actually reach the normaliser rather than merely that
    /// something was called.
    /// </para>
    /// </remarks>
    private sealed class StubRowHeightNormaliser : IRowHeightNormalisationPort
    {
        /// <summary>Creates the stub, optionally refusing through a thrown exception.</summary>
        /// <param name="succeeds">Whether <c>Normalise</c> returns a successful outcome.</param>
        /// <param name="throwsComException">
        /// Whether <c>Normalise</c> throws instead of returning. The guard under test is
        /// the <c>catch</c> around the call, and a typed refusal exercises the callee's own
        /// return path rather than that guard — so the exception needs its own case.
        /// </param>
        public StubRowHeightNormaliser(bool succeeds = true, bool throwsComException = false)
        {
            Succeeds = succeeds;
            ThrowsComException = throwsComException;
        }

        public bool Succeeds { get; set; } = true;

        /// <summary>Whether <c>Normalise</c> throws the COMException a host refusal produces.</summary>
        public bool ThrowsComException { get; set; }

        public double? AnchorRowHeightPt { get; private set; }

        public double? PaddingRowHeightPt { get; private set; }

        public RowHeightNormalisationOutcome Normalise(
            double managedHeightPt,
            double splitterHeightPt,
            double spacerHeightPt,
            double headerHeightPt,
            double reservedRowHeightPt,
            double paddingRowHeightPt,
            double anchorRowHeightPt)
        {
            PaddingRowHeightPt = paddingRowHeightPt;
            AnchorRowHeightPt = anchorRowHeightPt;
            if (ThrowsComException)
            {
                throw RefusedComException();
            }

            return Succeeds
                ? RowHeightNormalisationOutcome.Ok(2)
                : RowHeightNormalisationOutcome.Refused(
                    RowHeightNormalisationRefusalReason.AnchorRowNotOwned);
        }
    }

    private sealed class TestableInserter(
        Excel.Application application,
        IWorksheetProtectionGuard guard,
        Func<Excel.Sheets, int, Excel.Worksheet> sheetAt,
        Func<Excel.ListObjects, int, Excel.ListObject> tableAt,
        Func<Excel.ListColumns, int, Excel.ListColumn> columnAt,
        Func<Excel.ListObject, Excel.ListRows> rowsAt,
        Func<Excel.ListObject, Excel.Range> tableRangeAt,
        Func<Excel.Range, Excel.Range> rangeRowsAt,
        Func<Excel.Range, int> rangeRowCountAt,
        Func<Excel.Range, int, Excel.Range> rangeAt,
        Func<Excel.Range, object?> rangeValueAt,
        Action<Excel.Range> clearRange,
        Func<Excel.ListRows, Excel.ListRow> addRow,
        Func<Excel.Application, Excel.Range?> activeCell,
        Func<Excel.ListObject, bool> tableActive,
        Func<Excel.Range, int> rangeRow,
        Func<Excel.ListRow, int> rowIndex,
        Func<Excel.ListRow, Excel.Range> rowRange,
        Action insertWorksheetRow,
        Action<Excel.Range, double> setRowHeight,
        List<string> callOrder,
        Action<Excel.ListObject, int> insertWorksheetRowAt,
        Func<Excel.ListRows, int, Excel.ListRow> listRowAt,
        List<int> insertedWorksheetRows,
        ITypeOptionsMaterialiser? typeOptionsMaterialiser = null,
        IRowHeightNormalisationPort? rowHeightNormaliser = null)
        : ExcelGanttRowInserter(application, guard, typeOptionsMaterialiser, rowHeightNormaliser)
    {
        private readonly Func<Excel.Sheets, int, Excel.Worksheet> _sheetAt = sheetAt;
        private readonly Func<Excel.ListObjects, int, Excel.ListObject> _tableAt = tableAt;
        private readonly Func<Excel.ListColumns, int, Excel.ListColumn> _columnAt = columnAt;
        private readonly Func<Excel.ListObject, Excel.ListRows> _rowsAt = rowsAt;
        private readonly Func<Excel.ListObject, Excel.Range> _tableRangeAt = tableRangeAt;
        private readonly Func<Excel.Range, Excel.Range> _rangeRowsAt = rangeRowsAt;
        private readonly Func<Excel.Range, int> _rangeRowCountAt = rangeRowCountAt;
        private readonly Func<Excel.Range, int, Excel.Range> _rangeAt = rangeAt;
        private readonly Func<Excel.Range, object?> _rangeValueAt = rangeValueAt;
        private readonly Action<Excel.Range> _clearRange = clearRange;
        private readonly Func<Excel.ListRows, Excel.ListRow> _addRow = addRow;
        private readonly Func<Excel.Application, Excel.Range?> _activeCell = activeCell;
        private readonly Func<Excel.ListObject, bool> _tableActive = tableActive;
        private readonly Func<Excel.Range, int> _rangeRow = rangeRow;
        private readonly Func<Excel.ListRow, int> _rowIndex = rowIndex;
        private readonly Func<Excel.ListRow, Excel.Range> _rowRange = rowRange;
        private readonly Action _insertWorksheetRow = insertWorksheetRow;
        private readonly Action<Excel.Range, double> _setRowHeight = setRowHeight;
        private readonly List<string> _callOrder = callOrder;
        private readonly Action<Excel.ListObject, int> _insertWorksheetRowAt = insertWorksheetRowAt;
        private readonly Func<Excel.ListRows, int, Excel.ListRow> _listRowAt = listRowAt;
        private readonly List<int> _insertedWorksheetRows = insertedWorksheetRows;

        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => _sheetAt(sheets, index);
        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) => _tableAt(listObjects, index);
        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => _columnAt(columns, index);
        internal override Excel.ListRows GetListRows(Excel.ListObject table) => _rowsAt(table);
        internal override Excel.Range GetTableRange(Excel.ListObject table) => _tableRangeAt(table);
        internal override Excel.Range GetRangeRows(Excel.Range range) => _rangeRowsAt(range);
        internal override int GetRangeRowCount(Excel.Range rows) => _rangeRowCountAt(rows);
        internal override Excel.Range GetRangeAt(Excel.Range rows, int index) => _rangeAt(rows, index);
        internal override object? GetRangeValue2(Excel.Range range) => _rangeValueAt(range);
        internal override void ClearRange(Excel.Range range) => _clearRange(range);
        internal override Excel.ListRow AddRow(Excel.ListRows rows) => _addRow(rows);
        internal override Excel.Range? GetActiveCell(Excel.Application application) => _activeCell(application);
        internal override bool GetTableActive(Excel.ListObject table) => _tableActive(table);
        internal override int GetRangeRow(Excel.Range range) => _rangeRow(range);
        internal override int GetRowIndex(Excel.ListRow row) => _rowIndex(row);
        internal override Excel.Range GetRowRange(Excel.ListRow row) => _rowRange(row);
        internal override void SetRowHeight(Excel.Range range, double heightPt)
        {
            _setRowHeight(range, heightPt);
            _callOrder.Add($"SetRowHeight({heightPt})");
        }

        /// <summary>
        /// The workbook-row insert below the table is a real COM call, so it is
        /// replaced. It cannot be left to the base implementation in a unit test:
        /// there is no host, and the base method would report "not reserved" for the
        /// absence of a host rather than for a genuine refusal, so the append branch
        /// would pass for the wrong reason.
        /// </summary>
        /// <remarks>
        /// The recorded result is a field rather than a hardcoded success so a test
        /// can make the host refuse and assert the command says so.
        /// </remarks>
        internal override bool InsertWorksheetRowBelowTable(Excel.ListObject table)
        {
            ArgumentNullException.ThrowIfNull(table);
            _insertWorksheetRow();
            _callOrder.Add("InsertWorksheetRowBelowTable");

            // Recorded here as well so a test can assert WHICH worksheet row either
            // branch displaced. The base method computes lastRow + 1 from the table
            // range, which this harness seeds as row 1 with 4 rows.
            Excel.Range? tableRange = _tableRangeAt(table);
            if (tableRange is not null)
            {
                int lastRow = _rangeRow(tableRange) + 4 - 1;
                _insertedWorksheetRows.Add(lastRow + 1);
            }

            return PaddingRowReservedByHost;
        }

        internal override bool InsertWorksheetRow(Excel.ListObject table, int worksheetRow)
        {
            ArgumentNullException.ThrowIfNull(table);
            _insertWorksheetRowAt(table, worksheetRow);
            _callOrder.Add($"InsertWorksheetRow({worksheetRow})");
            return PaddingRowReservedByHost;
        }

        internal override Excel.ListRow GetListRowAt(Excel.ListRows rows, int index) =>
            _listRowAt(rows, index);

        /// <summary>Whether the simulated host reserved the row. Defaults to success.</summary>
        internal bool PaddingRowReservedByHost { get; set; } = true;
    }

    private sealed class Graph
    {
        /// <summary>
        /// Which of the row-creation seams the host refuses, if any.
        /// </summary>
        internal enum HostRefusal
        {
            /// <summary>The host behaves normally.</summary>
            None = 0,

            /// <summary><c>AddRow</c> throws.</summary>
            AddRow = 1,

            /// <summary><c>GetListRowAt</c> throws.</summary>
            GetListRowAt = 2,

            /// <summary><c>SetRowHeight</c> throws.</summary>
            SetRowHeight = 3,
        }

        /// <summary>Which row-creation seam refuses, if any.</summary>
        internal HostRefusal Refusal { get; set; } = HostRefusal.None;
        public Mock<Excel.Application> Application { get; } = new();
        public Mock<Excel.Workbook> Workbook { get; } = new();
        public Mock<Excel.Sheets> Sheets { get; } = new();
        public Mock<Excel.Worksheet> Worksheet { get; } = new();
        public Mock<Excel.ListObjects> ListObjects { get; } = new();
        public Mock<Excel.ListObject> Table { get; } = new();
        public Mock<Excel.ListColumns> ListColumns { get; } = new();
        /// <summary>
        /// Sized from the schema rather than a literal, so adding a column cannot
        /// leave the fake short of a mock and fail with an IndexOutOfRange that
        /// names neither the column nor the contract.
        /// </summary>
        public Mock<Excel.ListColumn>[] Columns { get; } =
            [.. Enumerable.Range(0, GanttTableSchema.Default.Columns.Count).Select(_ => new Mock<Excel.ListColumn>())];
        public Mock<Excel.ListRows> ListRows { get; } = new();
        public Mock<Excel.ListRow> NewRow { get; } = new();
        public Mock<Excel.ListRow> PositionedRow { get; } = new();
        public Mock<Excel.Range> RowRange { get; } = new();
        public Mock<Excel.Range> TableRange { get; } = new();
        public Mock<Excel.Range> TableRows { get; } = new();
        public Mock<Excel.Range> InitialBlankRow { get; } = new();
        public Mock<Excel.Range> ActiveCell { get; } = new();
        public object? WrittenValue { get; private set; }
        public int AddRowCalls { get; private set; }
        public int LastInsertionPosition { get; private set; }
        public int ClearRangeCalls { get; private set; }

        /// <summary>
        /// How many real worksheet rows were inserted BELOW the table, which is what
        /// moves the reserved bottom padding row down on the append branch.
        /// </summary>
        public int WorksheetRowInserts { get; private set; }

        /// <summary>
        /// The order in which the adapter performed its two order-sensitive calls.
        /// The append branch is only correct when the worksheet-row insert happens
        /// BEFORE <c>ListRows.Add()</c>, so the sequence itself is the assertion
        /// target rather than the mere presence of either call.
        /// </summary>
        public List<string> CallOrder { get; } = [];

        /// <summary>Every row height the adapter wrote, as (range, points).</summary>
        public List<(Excel.Range Range, double HeightPt)> RowHeightsWritten { get; } = [];

        /// <summary>Worksheet row indexes passed to the in-table insert seam.</summary>
        public List<int> InsertedWorksheetRows { get; } = [];

        /// <summary>Body-row indexes the adapter read back after inserting.</summary>
        public List<int> ListRowAtCalls { get; } = [];

        /// <summary>The stub normaliser handed to every built inserter unless a test passes its
        /// own, so the reserved rows can be asserted without driving real COM.
        /// </summary>
        public StubRowHeightNormaliser RowHeightNormaliser { get; } = new();

        public void RecordWrittenValue(object value) => WrittenValue = value;

        public void RecordAddRow() => AddRowCalls++;

        public void RecordClearRange() => ClearRangeCalls++;

        public Graph()
        {
            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            _ = Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            _ = Sheets.SetupGet(s => s.Count).Returns(1);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = ListObjects.SetupGet(l => l.Count).Returns(1);
            _ = Table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = Table.SetupGet(t => t.ListColumns).Returns(ListColumns.Object);
            _ = ListColumns.SetupGet(c => c.Count).Returns(Columns.Length);
            for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
            {
                _ = Columns[index].SetupGet(c => c.Name).Returns(GanttTableSchema.Default.Columns[index].Name);
            }
            _ = Table.SetupGet(t => t.ListRows).Returns(ListRows.Object);
            _ = Table.SetupGet(t => t.Range).Returns(TableRange.Object);
            _ = Table.SetupGet(t => t.Active).Returns(false);
            _ = TableRange.SetupGet(r => r.Row).Returns(1);
            _ = ActiveCell.SetupGet(r => r.Row).Returns(1);
            _ = Application.SetupGet(a => a.ActiveCell).Returns(ActiveCell.Object);
            _ = TableRange.SetupGet(r => r.Rows).Returns(TableRows.Object);
            _ = TableRows.SetupGet(r => r.Count).Returns(2);
            _ = TableRows.SetupGet(r => r[2]).Returns(InitialBlankRow.Object);
            _ = InitialBlankRow.SetupGet(r => r.Value2)
                .Returns(Array.CreateInstance(typeof(object), [1, 14], [1, 1]));
            _ = ListRows.SetupGet(r => r.Count).Returns(1);
            _ = NewRow.SetupGet(r => r.Index).Returns(2);
            _ = NewRow.SetupGet(r => r.Range).Returns(RowRange.Object);
            _ = PositionedRow.SetupGet(r => r.Index).Returns(2);
            _ = PositionedRow.SetupGet(r => r.Range).Returns(RowRange.Object);
            _ = RowRange.SetupSet(r => r.Value2 = It.IsAny<object>())
                .Callback<object>(RecordWrittenValue);
            _ = InitialBlankRow.SetupSet(r => r.Value2 = It.IsAny<object>())
                .Callback<object>(RecordWrittenValue);
            _ = InitialBlankRow.Setup(r => r.ClearContents()).Callback(RecordClearRange);
            _ = ListRows.Setup(r => r.Add(Type.Missing)).Callback(RecordAddRow).Returns(NewRow.Object);
        }

        public TestableInserter Build(
            IWorksheetProtectionGuard guard,
            ITypeOptionsMaterialiser? typeOptionsMaterialiser = null,
            StubRowHeightNormaliser? rowHeightNormaliser = null) => new(
            Application.Object,
            guard,
            (_, _) => Worksheet.Object,
            (listObjects, _) =>
            {
                Assert.Same(ListObjects.Object, listObjects);
                return Table.Object;
            },
            (_, index) => Columns[index - 1].Object,
            table =>
            {
                Assert.Same(Table.Object, table);
                return ListRows.Object;
            },
            table =>
            {
                Assert.Same(Table.Object, table);
                return TableRange.Object;
            },
            range =>
            {
                Assert.Same(TableRange.Object, range);
                return TableRows.Object;
            },
            rows =>
            {
                Assert.Same(TableRows.Object, rows);
                return 2;
            },
            (rows, index) =>
            {
                Assert.Same(TableRows.Object, rows);
                Assert.Equal(2, index);
                return InitialBlankRow.Object;
            },
            range => range == InitialBlankRow.Object ? range.Value2 : null,
            range => ClearRangeCalls++,
            rows =>
            {
                Assert.Same(ListRows.Object, rows);
                if (Refusal == HostRefusal.AddRow)
                {
                    throw RefusedComException();
                }

                AddRowCalls++;
                CallOrder.Add("AddRow");
                return NewRow.Object;
            },
            _ => ActiveCell.Object,
            table => table.Active,
            range => range.Row,
            row =>
            {
                Assert.Same(NewRow.Object, row);
                return 2;
            },
            row =>
            {
                Assert.Same(NewRow.Object, row);
                return RowRange.Object;
            },
            () => WorksheetRowInserts++,
            (range, heightPt) =>
            {
                if (Refusal == HostRefusal.SetRowHeight)
                {
                    throw RefusedComException();
                }

                RowHeightsWritten.Add((range, heightPt));
            },
            CallOrder,
            (table, row) =>
            {
                WorksheetRowInserts++;
                InsertedWorksheetRows.Add(row);
            },
            (rows, index) =>
            {
                if (Refusal == HostRefusal.GetListRowAt)
                {
                    throw RefusedComException();
                }

                ListRowAtCalls.Add(index);
                return NewRow.Object;
            },
            InsertedWorksheetRows,
            typeOptionsMaterialiser ?? new StubTypeOptionsMaterialiser(),
            rowHeightNormaliser ?? RowHeightNormaliser);
    }

    [Fact]
    public void Insert_appends_one_bulk_row_with_exact_scaffold_values()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => id);

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.BodyIndex);
        Assert.Equal(1, graph.AddRowCalls);
        guard.Verify(g => g.Query(), Times.Once);
        var matrix = Assert.IsType<object[,]>(graph.WrittenValue);
        Assert.Equal(1, matrix.GetLength(0));
        Assert.Equal(GanttTableSchema.Default.Columns.Count, matrix.GetLength(1));

        // Unreversed headers here, so a logical column's physical position IS its
        // schema index -- looked up rather than written as a literal. The previous
        // pinned `matrix[0, 8]` for StyleKey silently became SiblingOrder's cell
        // when R4.7A inserted that column, which is how the production
        // GanttRowDefaults defect was found.
        int Of(string columnName)
        {
            for (var i = 0; i < GanttTableSchema.Default.Columns.Count; i++)
            {
                if (GanttTableSchema.Default.Columns[i].Name == columnName)
                {
                    return i;
                }
            }

            throw new InvalidOperationException($"Column '{columnName}' is not in the schema.");
        }

        Assert.Equal(id.Value, matrix[0, Of("Id")]);
        Assert.Equal("As-Planned Activity", matrix[0, Of("Type")]);
        Assert.Equal("AsPlannedActivity", matrix[0, Of("StyleKey")]);

        // SiblingOrder is engine-maintained and has no scaffold default. The
        // writer coalesces a null cell value to string.Empty, so "blank" is an
        // empty string here, not null.
        Assert.Equal(string.Empty, matrix[0, Of("SiblingOrder")]);
    }

    /// <summary>
    /// The append branch must insert the genuine worksheet row BEFORE
    /// <c>ListRows.Add()</c>, because <c>Add()</c> with no position claims the row
    /// already below the table -- the 6pt reserved padding row -- as a body row.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the regression test for the reported defect. Measured 2026-10-03
    /// (<c>scripts/probe-rowinsert-anchoring.ps1</c> Q4): with the insert AFTER
    /// <c>Add()</c>, the new body row measured <strong>6pt</strong> and stayed 6pt
    /// even after the compensating insert, because that insert created a fresh row
    /// BELOW the table while the claimed row remained a squashed body row. The same
    /// probe's Q3 shows the reversed order yields 18pt.
    /// </para>
    /// <para>
    /// The assertion is on the ORDER, not on both calls having happened. The previous
    /// code made both calls and was still wrong, so a presence check would have
    /// passed straight through the defect.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_append_branch_inserts_the_worksheet_row_before_claiming_it_with_ListRows()
    {
        var graph = new Graph();
        var guard = NotProtected();
        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => id);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, graph.AddRowCalls);
        Assert.Equal(1, graph.WorksheetRowInserts);

        int insertAt = graph.CallOrder.IndexOf("InsertWorksheetRowBelowTable");
        int addAt = graph.CallOrder.IndexOf("AddRow");
        Assert.True(insertAt >= 0 && addAt >= 0, $"Call order was: {string.Join(" -> ", graph.CallOrder)}");
        Assert.True(
            insertAt < addAt,
            $"The worksheet row must be inserted BEFORE ListRows.Add(); order was: {string.Join(" -> ", graph.CallOrder)}");
    }

    /// <summary>
    /// The row handed back by <c>ListRows.Add()</c> must be written to the body-row
    /// height token, so a body row cannot be left at the padding row's 6pt.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Excel's own inheritance already yields the right height (probe Q3), so this
    /// write is belt-and-braces. It is asserted anyway because the inheritance is a
    /// host behaviour rather than a stated invariant, and the defect it guards was
    /// invisible to every existing test: the integration suite asserted the PADDING
    /// row's height, never the body row's, which is precisely the row that broke.
    /// </para>
    /// <para>
    /// The expected value comes from the token catalogue, never a literal, so a
    /// retuned token cannot leave this test pinning 18pt as if it were the contract.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_appended_body_row_is_written_to_the_body_row_height_token()
    {
        var graph = new Graph();
        var guard = NotProtected();
        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => id);

        Assert.True(outcome.Succeeded);
        (Excel.Range Range, double HeightPt) written = Assert.Single(graph.RowHeightsWritten);
        Assert.Same(graph.RowRange.Object, written.Range);
        Assert.Equal(GanttCatalogues.MetricDefault("GanttRowHeightPt"), written.HeightPt);
    }

    /// <summary>
    /// A POSITIONAL insert must insert a genuine worksheet row INSIDE the table,
    /// because <c>ListRows.Add(position)</c> does not shift the sheet.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the regression test for the reported defect. Measured 2026-10-03
    /// (<c>scripts/probe-positional-insert.ps1</c> Q1): <c>ListRows.Add(2)</c> moved all
    /// three probe shapes by <strong>delta=0</strong> and pushed the row below the
    /// table from 6pt to 15pt. It rearranges rows inside the table and consumes the
    /// padding row, leaving every shape and cell below untouched -- exactly the
    /// reported symptom, where the table's row data moves but the cells and Gantt
    /// shapes do not.
    /// </para>
    /// <para>
    /// The same probe's Q2 measured the fix: a worksheet row inserted inside the table's
    /// range moved all three shapes by delta=18 and auto-expanded the ListObject into
    /// the new row, so no <c>ListRows</c> call is needed at all.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_positional_branch_inserts_a_worksheet_row_inside_the_table()
    {
        var graph = new Graph();

        // Table ACTIVE with the active cell INSIDE the body: table starts at row 1
        // (header), so active row 2 is body index 1 and the target is body index 2,
        // which is worksheet row 1 + 2 = 3.
        _ = graph.Table.SetupGet(t => t.Active).Returns(true);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(2);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);
        _ = graph.TableRows.SetupGet(r => r.Count).Returns(4);

        GanttRowInsertOutcome outcome = graph.Build(NotProtected().Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => GanttRowId.Parse("G-0123456789abcdef0123456789abcdef"));

        Assert.True(outcome.Succeeded);

        // A real worksheet row went in, INSIDE the table's range.
        Assert.Equal(1, graph.WorksheetRowInserts);
        Assert.Equal(3, Assert.Single(graph.InsertedWorksheetRows));
        Assert.Contains("InsertWorksheetRow(3)", graph.CallOrder);

        // And no ListRows.Add of any kind: the table absorbs the inserted row, so the
        // adapter reads the row back instead of asking Excel to create one.
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Equal(2, Assert.Single(graph.ListRowAtCalls));
        Assert.DoesNotContain("AddRow", graph.CallOrder);
        Assert.DoesNotContain("AddRowAtPosition", graph.CallOrder);

        // The body row is normalised to the body height on this branch too.
        Assert.Equal(
            GanttCatalogues.MetricDefault("GanttRowHeightPt"),
            Assert.Single(graph.RowHeightsWritten).HeightPt);
    }

    /// <summary>
    /// A host refusal while writing the row is REPORTED, not thrown, and is reported
    /// honestly as added-but-not-completed.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the positive test for the guard (AGENTS.md validator rule). The insert
    /// path's COM calls were unguarded, so a live workbook produced a
    /// <see cref="System.Runtime.InteropServices.COMException"/> that escaped into the
    /// Ribbon callback and aborted the sequence after the worksheet row already
    /// existed -- leaving the plot at its old extent, which reads as "the bands did not
    /// stretch".
    /// </para>
    /// <para>
    /// The refusal reason is checked as well as the absence of an exception, because a
    /// bare "did not throw" would also pass if the adapter reported
    /// <c>TargetProtected</c> and pretended nothing had been written -- a claim the user
    /// can disprove by looking at the sheet, because the row is really there.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_host_refusal_writing_the_row_is_reported_rather_than_thrown()
    {
        var graph = new Graph();
        var guard = NotProtected();

        // The host refuses the bulk value write, exactly as Excel does on a locked or
        // otherwise hostile range.
        _ = graph.RowRange
            .SetupSet(r => r.Value2 = It.IsAny<object>())
            .Throws(RefusedComException());

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => GanttRowId.Parse("G-0123456789abcdef0123456789abcdef"));

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRowInsertRefusalReason.RowWriteRefused, outcome.Refusal);
        Assert.Null(outcome.BodyIndex);
    }

    /// <summary>
    /// Produces the exception a host actually throws for a refused write, without
    /// constructing the runtime-reserved <c>COMException</c> directly (CA2201).
    /// </summary>
    private static System.Runtime.InteropServices.COMException RefusedComException()
    {
        var exception = (System.Runtime.InteropServices.COMException?)
            System.Runtime.InteropServices.Marshal.GetExceptionForHR(unchecked((int)0x800A03EC));

        return exception ?? throw new InvalidOperationException(
            "The runtime did not produce a COMException for HRESULT 0x800A03EC.");
    }

    private static Mock<IWorksheetProtectionGuard> NotProtected()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        return guard;
    }

    [Fact]
    public void Insert_reuses_the_initial_blank_row_without_appending_a_second_row()
    {
        var graph = new Graph();
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(0);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => id);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.BodyIndex);
        Assert.Equal(0, graph.AddRowCalls);
        var matrix = Assert.IsType<object[,]>(graph.WrittenValue);
        Assert.Equal(id.Value, matrix[0, 0]);
        Assert.Equal("As-Planned Activity", matrix[0, 3]);
        Assert.Equal(0, graph.ClearRangeCalls);
    }

    [Fact]
    public void Insert_clears_the_reused_blank_row_when_type_options_refuses()
    {
        var graph = new Graph();
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(0);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var materialiser = new Mock<ITypeOptionsMaterialiser>();
        _ = materialiser.Setup(m => m.EnsureCurrent())
            .Returns(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch));

        GanttRowInsertOutcome outcome = graph.Build(guard.Object, materialiser.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.Equal(GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TypeOptionsUnavailable), outcome);
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Equal(1, graph.ClearRangeCalls);
        graph.NewRow.Verify(r => r.Delete(), Times.Never);
    }

    /// <summary>
    /// ADR-0035 D3: an active cell INSIDE the table inserts immediately BELOW it; an
    /// active cell anywhere else appends before the padding row.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This INVERTS <c>Every_insert_appends_and_never_uses_the_active_cell</c>, which
    /// asserted the opposite and was correct for the append-only rule ADR-0035 first
    /// landed. Replaced rather than deleted: the property worth keeping is stronger
    /// now, because the position depends on the active cell again AND both branches
    /// are pinned, including the header case that used to insert at position one.
    /// </para>
    /// <para>
    /// The in-table cases select the header and each body row. Note the fixture's
    /// table range starts at row 1, so the body occupies rows 2..4 and the header is
    /// row 1 -- the numbers here are fixture coordinates, not the live sheet's.
    /// </para>
    /// </remarks>
    [Theory]
    [InlineData(true, 1, 1)]   // header selected: the new row goes to the top
    [InlineData(true, 2, 2)]   // first body row: directly below it
    [InlineData(true, 3, 3)]   // middle body row
    [InlineData(true, 4, 0)]   // LAST body row appends rather than inserting below itself
    [InlineData(false, 2, 0)]  // active cell outside the table: append
    public void An_in_table_active_cell_inserts_below_it_and_an_outside_one_appends(
        bool tableActive,
        int activeRow,
        int expectedPosition)
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(tableActive);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(activeRow);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);
        _ = graph.TableRows.SetupGet(r => r.Count).Returns(4);
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.True(outcome.Succeeded);

        // EXACTLY ONE genuine worksheet row goes in on either branch, because only a
        // worksheet row insert moves the sheet. The table starts at row 1, so the
        // worksheet row for body position N is 1 + N.
        Assert.Equal(1, graph.WorksheetRowInserts);
        int insertedRow = expectedPosition == 0
            ? 1 + 3 + 1  // append: one past the last body row (3 body rows)
            : 1 + expectedPosition;
        Assert.Equal(insertedRow, Assert.Single(graph.InsertedWorksheetRows));

        // Only the APPEND still asks Excel to create a list row. The positional branch
        // inserts a worksheet row inside the table's range and the ListObject absorbs
        // it, so no ListRows call is made (probe Q2: ListRows 3 -> 4 with no Add).
        if (expectedPosition == 0)
        {
            Assert.Equal(1, graph.AddRowCalls);
            Assert.Empty(graph.ListRowAtCalls);
        }
        else
        {
            Assert.Equal(0, graph.AddRowCalls);
            Assert.Equal(expectedPosition, Assert.Single(graph.ListRowAtCalls));
        }

        // The row really was written, so neither branch is satisfied by a no-op.
        Assert.NotNull(graph.WrittenValue);
    }

    /// <summary>
    /// BOTH branches insert exactly one worksheet row, so the sheet shifts down once
    /// and the Gantt shapes move with it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This replaces a test that asserted the opposite and was wrong. It claimed a
    /// middle insert must NOT shift the sheet, on the grounds that
    /// <c>ListRows.Add(position)</c> "already inserts a real worksheet row". Measured
    /// 2026-10-03 (<c>scripts/probe-positional-insert.ps1</c> Q1), it does not:
    /// <c>ListRows.Add(2)</c> moved all three probe shapes by <strong>delta=0</strong>
    /// and pushed the row below the table from 6pt to 15pt. That test therefore
    /// encoded the defect as a requirement.
    /// </para>
    /// <para>
    /// Q2 of the same probe measured the corrected behaviour: a worksheet row inserted
    /// inside the table's range moved all three shapes by delta=18 with TopLeftCell
    /// following, and auto-expanded the ListObject into the new row.
    /// </para>
    /// </remarks>
    [Fact]
    public void Both_branches_insert_exactly_one_worksheet_row()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        var append = new Graph();
        _ = append.Table.SetupGet(t => t.Active).Returns(false);
        _ = append.ListRows.SetupGet(r => r.Count).Returns(3);
        _ = append.TableRows.SetupGet(r => r.Count).Returns(4);

        _ = append.Build(guard.Object).Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New);

        // Append: the row goes BELOW the table, so ListRows.Add() is still needed --
        // Q4 measured that a row inserted below the table does not auto-expand it.
        Assert.Equal(1, append.AddRowCalls);
        Assert.Equal(1, append.WorksheetRowInserts);
        Assert.Contains("InsertWorksheetRowBelowTable", append.CallOrder);

        var middle = new Graph();
        _ = middle.Table.SetupGet(t => t.Active).Returns(true);
        _ = middle.ActiveCell.SetupGet(r => r.Row).Returns(3);
        _ = middle.ListRows.SetupGet(r => r.Count).Returns(3);
        _ = middle.TableRows.SetupGet(r => r.Count).Returns(4);

        _ = middle.Build(guard.Object).Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New);

        // Middle: the row goes INSIDE the table, which shifts the sheet and is
        // absorbed by the ListObject, so no ListRows call at all.
        Assert.Equal(0, middle.AddRowCalls);
        Assert.Equal(1, middle.WorksheetRowInserts);
        Assert.DoesNotContain("AddRow", middle.CallOrder);
    }

    /// <summary>
    /// A refused padding-row insert is reported, not swallowed.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the defect behind "the padding row still takes data". The helper used
    /// to swallow every failure and return <c>void</c>, so a refused push-down was
    /// indistinguishable from a successful one and the command reported success
    /// either way -- leaving the table sitting on the chart's bottom margin with
    /// nothing said.
    /// </para>
    /// <para>
    /// The row is still written when the host refuses: the user asked for it and it
    /// is already in the table. What must not happen is silence, so the outcome
    /// carries the fact and the command can say so.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_refused_padding_row_insert_is_reported_rather_than_swallowed()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(false);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);

        TestableInserter inserter = graph.Build(guard.Object);
        inserter.PaddingRowReservedByHost = false;

        GanttRowInsertOutcome outcome = inserter.Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New);

        // The push-down was attempted...
        Assert.Equal(1, graph.WorksheetRowInserts);

        // ...the row was still written, so the insert is not a refusal...
        Assert.True(outcome.Succeeded);

        // ...but the lost margin is stated rather than passed over.
        Assert.False(outcome.PaddingRowReserved);
    }

    /// <summary>
    /// The appended row normalises the reserved rows below the body, with the anchor
    /// row's OWN token (ADR-0038 D1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the positive test for the behaviour the live gate demanded.</b> The
    /// reserved rows were only ever normalised through Refresh, so on the Add-Activity
    /// path they kept Excel's default height: measured live, every row below the table
    /// read <b>15pt</b> after Initialise plus three appends, leaving the chart's bottom
    /// margin at ~30pt of visible sheet instead of the 6pt the layout reserves.
    /// </para>
    /// <para>
    /// Asserted on the VALUES the stub received, not merely that normalisation was
    /// called: an inserter that passed the padding height for both rows would satisfy a
    /// presence check while making the anchor a 5.75pt strip rather than a sub-row
    /// anchor.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_appended_row_normalises_the_reserved_rows_below_the_body()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(false);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);

        TestableInserter inserter = graph.Build(guard.Object);

        GanttRowInsertOutcome outcome = inserter.Insert(
            GanttEntityType.AsPlannedActivity, GanttRowId.New);

        Assert.True(outcome.Succeeded);
        Assert.True(outcome.ReservedRowsNormalised);

        // The two reserved rows are distinguished, not given one shared height.
        Assert.Equal(
            GanttCatalogues.MetricDefault("ChartAnchorRowHeightPt"),
            graph.RowHeightNormaliser.AnchorRowHeightPt);
        Assert.Equal(
            GanttCatalogues.MetricDefault("ChartPaddingRowHeightPt"),
            graph.RowHeightNormaliser.PaddingRowHeightPt);
        Assert.NotEqual(
            graph.RowHeightNormaliser.AnchorRowHeightPt,
            graph.RowHeightNormaliser.PaddingRowHeightPt);

        // And the strip they reserve is still the designed total, so this cannot be
        // satisfied by two rows that happen to sum correctly by accident.
        Assert.Equal(
            GanttCatalogues.MetricDefault("ChartOuterPaddingPt"),
            (graph.RowHeightNormaliser.AnchorRowHeightPt ?? 0d)
                + (graph.RowHeightNormaliser.PaddingRowHeightPt ?? 0d),
            6);
    }

    /// <summary>
    /// A REFUSED normalisation is reported without undoing the insert, because the row
    /// is already in the sheet (ADR-0008, ADR-0038 D1).
    /// </summary>
    /// <remarks>
    /// Reporting the whole insert as failed would be a lie the user can disprove by
    /// looking at the sheet -- the same reasoning as the existing
    /// <c>RowWriteRefused</c> case. The outcome carries the fact so
    /// <c>AddRowCommand</c> can say the margin may be the wrong size and that a Refresh
    /// corrects it.
    /// </remarks>
    [Fact]
    public void A_refused_reserved_row_normalisation_is_reported_and_the_row_keeps_its_data()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(false);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);
        graph.RowHeightNormaliser.Succeeds = false;

        TestableInserter inserter = graph.Build(guard.Object);

        GanttRowInsertOutcome outcome = inserter.Insert(
            GanttEntityType.AsPlannedActivity, GanttRowId.New);

        // The insert stands: the row was written and nothing was rolled back...
        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.BodyIndex);
        Assert.NotNull(graph.WrittenValue);

        // ...and the un-normalised margin is stated rather than passed over.
        Assert.False(outcome.ReservedRowsNormalised);

        // A refusal is not silently "mostly fine": it must not read as the success
        // case, which is what the padding-row tests above guard against for their own
        // field.
        Assert.True(outcome.PaddingRowReserved);
    }

    /// <summary>
    /// A REFUSED positional insert refuses the whole insert, before the adapter reads
    /// back the row that now occupies the target position.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the positive test for the guard, and it is the only path that can
    /// corrupt.</b> The positional branch finds its row with
    /// <c>GetListRowAt(rows, position)</c>, which returns whatever row already sits
    /// there. So when the worksheet-row insert is refused, continuing would read the
    /// USER'S EXISTING ROW and hand the bulk write a scaffold payload aimed at it: the
    /// Id, Type, dates, and style of the new row would overwrite real schedule data,
    /// and the command would report success.
    /// </para>
    /// <para>
    /// Asserted on the absence of the read and the write, not only on the refusal: a
    /// refusal raised <em>after</em> <c>GetListRowAt</c> would leave the data already
    /// overwritten, so the ordering is the behaviour under test.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_refused_positional_insert_refuses_before_reading_or_writing_the_existing_row()
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(true);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(2);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);

        TestableInserter inserter = graph.Build(NotProtected().Object);
        inserter.PaddingRowReservedByHost = false;

        GanttRowInsertOutcome outcome = inserter.Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRowInsertRefusalReason.RowInsertRefused, outcome.Refusal);

        // The dangerous read never happened, so no existing row was ever in scope for
        // a write...
        Assert.Empty(graph.ListRowAtCalls);

        // ...and nothing was written, not even to a row the adapter would have had to
        // guess at.
        Assert.Null(graph.WrittenValue);
        Assert.Empty(graph.RowHeightsWritten);
    }

    /// <summary>
    /// The APPEND branch is unchanged by that guard: a refused push-down still adds a
    /// row, and reports only the lost padding row.
    /// </summary>
    /// <remarks>
    /// The counterpart that stops the guard from being "refuse whenever the insert
    /// fails". <c>AddRow</c> creates a genuine new row below the table, so the row
    /// exists and the user asked for it; the only casualty is the chart's bottom
    /// margin, which <c>PaddingRowReserved</c> already carries (ADR-0035 D3).
    /// </remarks>
    [Fact]
    public void A_refused_append_still_adds_its_row_and_reports_only_the_lost_padding_row()
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(false);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);

        TestableInserter inserter = graph.Build(NotProtected().Object);
        inserter.PaddingRowReservedByHost = false;

        GanttRowInsertOutcome outcome = inserter.Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, graph.AddRowCalls);
        Assert.NotNull(graph.WrittenValue);
        Assert.False(outcome.PaddingRowReserved);
    }

    /// <summary>
    /// A host refusal while CREATING or SIZING the row is reported, not thrown.
    /// </summary>
    /// <remarks>
    /// The three sites — <c>GetListRowAt</c>, <c>AddRow</c>, <c>SetRowHeight</c> — were
    /// unguarded while the bulk <c>WriteRow</c> beside them was guarded, so a refusal
    /// escaped into the Ribbon callback with the worksheet row already inserted and the
    /// sequence stopped partway. Each site is its own case: a guard tested through only
    /// one of them would still pass with the other two removed.
    /// </remarks>
    [Theory]
    [InlineData("AddRow", false)]
    [InlineData("SetRowHeight", false)]
    [InlineData("GetListRowAt", true)]
    [InlineData("SetRowHeight", true)]
    public void A_host_refusal_creating_or_sizing_the_row_is_reported_rather_than_thrown(
        string seam,
        bool positional)
    {
        var graph = new Graph();
        RefuseRowCreationSeam(graph, seam);
        _ = graph.Table.SetupGet(t => t.Active).Returns(positional);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(2);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);

        Exception? escaped = Record.Exception(() => graph.Build(NotProtected().Object)
            .Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New));

        Assert.Null(escaped);
    }

    /// <summary>
    /// The same refusals report <c>RowWriteRefused</c>: not a throw, and not a success.
    /// </summary>
    /// <remarks>
    /// The reason is load-bearing rather than cosmetic. <c>RowWriteRefused</c> is the
    /// member that means "the row IS in the worksheet, so do not say the insert
    /// failed" — and by this point the worksheet row really has been inserted. A
    /// "nothing happened" reason would be a claim the user can disprove by looking at
    /// the sheet.
    /// </remarks>
    [Theory]
    [InlineData("AddRow", false)]
    [InlineData("SetRowHeight", false)]
    [InlineData("GetListRowAt", true)]
    public void A_host_refusal_creating_or_sizing_the_row_is_reported_as_added_but_unfinished(
        string seam,
        bool positional)
    {
        var graph = new Graph();
        RefuseRowCreationSeam(graph, seam);
        _ = graph.Table.SetupGet(t => t.Active).Returns(positional);
        _ = graph.ActiveCell.SetupGet(r => r.Row).Returns(2);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);

        GanttRowInsertOutcome outcome = graph.Build(NotProtected().Object)
            .Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRowInsertRefusalReason.RowWriteRefused, outcome.Refusal);
        Assert.Null(outcome.BodyIndex);
    }

    /// <summary>
    /// A host refusal while normalising the reserved rows leaves the insert standing and
    /// is reported as an un-normalised margin.
    /// </summary>
    /// <remarks>
    /// This call runs LAST, after the row is written, so a thrown refusal here escaped a
    /// completed insert into the Ribbon callback. Reporting the whole insert as failed
    /// would be the same lie as elsewhere — the row is on the sheet — and
    /// <c>ReservedRowsNormalised</c> is exactly the field that says "the margin may be
    /// the wrong size", which a Refresh repairs.
    /// </remarks>
    [Fact]
    public void A_refused_normalisation_leaves_the_insert_successful_with_the_margin_flagged()
    {
        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(false);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);
        graph.RowHeightNormaliser.ThrowsComException = true;

        Exception? escaped = Record.Exception(() => graph.Build(NotProtected().Object)
            .Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New));

        Assert.Null(escaped);

        // And the outcome is the success-with-a-flag one, not a refusal.
        GanttRowInsertOutcome outcome = graph.Build(NotProtected().Object)
            .Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New);
        Assert.True(outcome.Succeeded);
        Assert.False(outcome.ReservedRowsNormalised);

        // The row still holds its scaffold values: nothing was rolled back.
        Assert.NotNull(graph.WrittenValue);
        Assert.True(outcome.PaddingRowReserved);
    }

    /// <summary>
    /// A successful push-down reports the padding row as reserved.
    /// </summary>
    /// <remarks>
    /// The counterpart to the refusal test. Without it the field could default to
    /// false and every passing insert would cry wolf.
    /// </remarks>
    [Fact]
    public void A_successful_padding_row_insert_reports_the_row_as_reserved()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        var graph = new Graph();
        _ = graph.Table.SetupGet(t => t.Active).Returns(false);
        _ = graph.ListRows.SetupGet(r => r.Count).Returns(3);

        GanttRowInsertOutcome outcome =
            graph.Build(guard.Object).Insert(GanttEntityType.AsPlannedActivity, GanttRowId.New);

        Assert.True(outcome.Succeeded);
        Assert.True(outcome.PaddingRowReserved);
    }

    /// <summary>
/// The positional path is reached through the active cell and is distinct from the
/// append branch, which is what makes the rule conditional rather than global.
/// </summary>
    /// <remarks>
    /// The behavioural theory above cannot separate the branches on its own, so the
    /// counters come from the adapter's own seams. This source assertion additionally
    /// pins that the position resolver is actually consulted -- so a future change
    /// cannot quietly hard-code "always append" behind them.
    /// <para>
    /// It also pins that the POSITIONAL branch no longer calls <c>ListRows.Add</c>.
    /// Measured 2026-10-03 (<c>scripts/probe-positional-insert.ps1</c> Q1), a
    /// positional <c>Add</c> moved no shape at all, so its presence here is the defect.
    /// The positional branch inserts a worksheet row inside the table's range instead,
    /// which shifts the sheet and is absorbed by the <c>ListObject</c>.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_position_is_resolved_from_the_active_cell_rather_than_hard_coded()
    {
        string source = File.ReadAllText(LocateInserterSource());

        Assert.Contains("GetInsertionPosition", source, StringComparison.Ordinal);
        Assert.Contains("GetActiveCell", source, StringComparison.Ordinal);
        Assert.Contains("GetTableActive", source, StringComparison.Ordinal);

        // The positional branch inserts a real worksheet row, and reads the row back.
        Assert.Contains("InsertWorksheetRow(", source, StringComparison.Ordinal);
        Assert.Contains("GetListRowAt(rows", source, StringComparison.Ordinal);

        // Only the append still asks Excel to create a list row.
        Assert.Contains("rows.Add(Type.Missing)", source, StringComparison.Ordinal);
        Assert.DoesNotContain("rows.Add(position)", source, StringComparison.Ordinal);
    }

    private static string LocateInserterSource()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            string candidate = Path.Combine(
                directory.FullName,
                "src",
                "GanttCreator.Office",
                "ExcelGanttRowInserter.cs");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not locate src/GanttCreator.Office/ExcelGanttRowInserter.cs from " + AppContext.BaseDirectory);
    }

    [Fact]
    public void Insert_maps_columns_by_header_name_not_position()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;

        // Reverse the table's physical order so the logical scaffold cannot be
        // found by position alone.
        for (var index = 0; index < columns.Count; index++)
        {
            var schemaIndex = columns.Count - index - 1;
            _ = graph.Columns[index].SetupGet(c => c.Name).Returns(columns[schemaIndex].Name);
        }

        // Ask the fake which physical index now carries each header, rather than
        // recomputing it. GetColumnAt is 1-based and resolves to Columns[index - 1],
        // so a second arithmetic copy of that rule is a second place to be wrong
        // by one, and the failure reads as a production defect when it is a
        // mistake in the expectation.
        int PhysicalOf(string logicalName) => Array.FindIndex(graph.Columns, c => c.Object.Name == logicalName);

        var id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");
        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(GanttEntityType.Delineator, () => id);

        Assert.True(outcome.Succeeded);
        var matrix = Assert.IsType<object[,]>(graph.WrittenValue);

        // The physical column for a logical column is the REVERSED index, because
        // the loop above gave physical index i the name of logical
        // (Count - 1 - i). Deriving both from the schema means adding a column
        // moves them together; a pinned literal would have asserted the wrong
        // cell and still passed, which is the failure mode this test exists to
        // rule out.
        // The scaffold writes Id, StyleKey and Type and leaves the rest null --
        // SortOrder included, which has no default. The original assertion of
        // string.Empty at a pinned physical 0 was only true by accident: it read
        // whatever column the reversal happened to put there. What this test is
        // for is that the named values land in the header-named cells.
        Assert.Equal(id.Value, matrix[0, PhysicalOf("Id")]);
        Assert.Equal("Delineator", matrix[0, PhysicalOf("Type")]);
        Assert.Equal("DefaultDelineator", matrix[0, PhysicalOf("StyleKey")]);
    }

    [Theory]
    [InlineData(ProtectionGuardOutcome.SheetProtected, GanttRowInsertRefusalReason.TargetProtected)]
    [InlineData(ProtectionGuardOutcome.WorkbookStructureProtected, GanttRowInsertRefusalReason.TargetProtected)]
    [InlineData(ProtectionGuardOutcome.NoActiveWorkbook, GanttRowInsertRefusalReason.NoActiveWorkbook)]
    public void Insert_refuses_before_any_workbook_mutation(
        ProtectionGuardOutcome guardOutcome,
        GanttRowInsertRefusalReason expected)
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(guardOutcome);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.False(outcome.Succeeded);
        Assert.Equal(expected, outcome.Refusal);
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Null(graph.WrittenValue);
        guard.Verify(g => g.Query(), Times.Once);
        graph.Workbook.VerifyGet(w => w.Sheets, Times.Never);
    }

    [Fact]
    public void Insert_returns_table_missing_when_no_table_is_found()
    {
        var graph = new Graph();
        _ = graph.ListObjects.SetupGet(l => l.Count).Returns(0);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);

        GanttRowInsertOutcome outcome = graph.Build(guard.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.New);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRowInsertRefusalReason.TableMissing, outcome.Refusal);
        Assert.Equal(0, graph.AddRowCalls);
        Assert.Null(graph.WrittenValue);
    }

    [Fact]
    public void Insert_deletes_the_new_row_when_type_options_refuses()
    {
        var graph = new Graph();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var materialiser = new Mock<ITypeOptionsMaterialiser>();
        _ = materialiser.Setup(m => m.EnsureCurrent())
            .Returns(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch));

        GanttRowInsertOutcome outcome = graph.Build(guard.Object, materialiser.Object).Insert(
            GanttEntityType.AsPlannedActivity,
            () => GanttRowId.Parse("G-0123456789abcdef0123456789abcdef"));

        Assert.Equal(GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.TypeOptionsUnavailable), outcome);
        graph.NewRow.Verify(r => r.Delete(), Times.Once);
    }
}
