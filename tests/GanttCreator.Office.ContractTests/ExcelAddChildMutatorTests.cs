using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R5.3 WI-E add-child adapter.
/// </summary>
/// <remarks>
/// The tests drive the adapter through the same fake-COM-graph seam
/// discipline as the R5.3 parent/child mutator tests: every COM
/// parameterised property is an <c>internal virtual</c> seam the testable
/// subclass overrides, so the full guard-first pipeline runs without a host.
/// The Core planner runs for real -- these tests pin the ADAPTER's
/// application of the plan, not the planner's arithmetic.
/// </remarks>
public sealed class ExcelAddChildMutatorTests
{
    private const string ParentId = "G-0123456789abcdef0123456789abcdef";
    private const string ChildId = "G-fedcba9876543210fedcba9876543210";

    /// <summary>
    /// The testable adapter: every COM seam is overridden, so the real
    /// pipeline -- guard, locate, resolve, read, plan, insert, write -- runs.
    /// </summary>
    private sealed class TestableMutator(
        Excel.Application? application,
        IWorksheetProtectionGuard? guard,
        IGanttRowInserter? inserter,
        IOutlineGroupPort? outlineGroups,
        Func<Excel.Sheets, int, object> sheetAt,
        Func<Excel.ListObjects, int, Excel.ListObject> tableAt,
        Func<Excel.ListColumns, int, Excel.ListColumn> columnAt,
        Func<Excel.ListObject, Excel.ListRows> listRowsAt,
        Func<Excel.ListRows, int> listRowCountAt,
        Func<Excel.ListObject, Excel.Range> tableRangeAt,
        Func<Excel.ListObject, Excel.Range?> dataBodyRangeAt,
        Func<Excel.Range, Excel.Range> rangeRowsAt,
        Func<Excel.Range, int, Excel.Range> rangeAt,
        Func<Excel.Range, int, Excel.Range> cellAt,
        Func<Excel.Range, object?> rangeValue2At,
        Action<Excel.Range, object> setRangeValues,
        Func<Excel.Application, Excel.Range?> activeCellAt,
        Func<Excel.ListObject, bool> tableActiveAt,
        Func<Excel.Range, int> rangeRowAt) : ExcelAddChildMutator(application, guard, inserter, outlineGroups)
    {
        private readonly Func<Excel.Sheets, int, object> _sheetAt = sheetAt;

        private readonly Func<Excel.ListObjects, int, Excel.ListObject> _tableAt = tableAt;

        private readonly Func<Excel.ListColumns, int, Excel.ListColumn> _columnAt = columnAt;

        private readonly Func<Excel.ListObject, Excel.ListRows> _listRowsAt = listRowsAt;

        private readonly Func<Excel.ListRows, int> _listRowCountAt = listRowCountAt;

        private readonly Func<Excel.ListObject, Excel.Range> _tableRangeAt = tableRangeAt;

        private readonly Func<Excel.ListObject, Excel.Range?> _dataBodyRangeAt = dataBodyRangeAt;

        private readonly Func<Excel.Range, Excel.Range> _rangeRowsAt = rangeRowsAt;

        private readonly Func<Excel.Range, int, Excel.Range> _rangeAt = rangeAt;

        private readonly Func<Excel.Range, int, Excel.Range> _cellAt = cellAt;

        private readonly Func<Excel.Range, object?> _rangeValue2At = rangeValue2At;

        private readonly Action<Excel.Range, object> _setRangeValues = setRangeValues;

        private readonly Func<Excel.Application, Excel.Range?> _activeCellAt = activeCellAt;

        private readonly Func<Excel.ListObject, bool> _tableActiveAt = tableActiveAt;

        private readonly Func<Excel.Range, int> _rangeRowAt = rangeRowAt;

        internal override bool TryFindTable(
            Excel.Sheets sheets,
            out Excel.Worksheet? worksheet,
            out Excel.ListObject? table)
        {
            worksheet = null;
            table = null;
            for (int index = 1; index <= 1; index++)
            {
                if (_sheetAt(sheets, index) is Excel.Worksheet candidate)
                {
                    worksheet = candidate;
                    table = _tableAt(candidate.ListObjects, index);
                    return true;
                }
            }

            return false;
        }

        internal override bool TryBuildColumnMap(Excel.ListObject table, out int[]? columnMap)
        {
            columnMap = null;
            IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
            int[] map = new int[schema.Count];
            Dictionary<string, int> byName = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int index = 1; index <= schema.Count; index++)
            {
                byName[_columnAt(table.ListColumns, index).Name] = index;
            }

            for (int schemaIndex = 0; schemaIndex < schema.Count; schemaIndex++)
            {
                GanttTableColumn expected = schema[schemaIndex];
                if (!byName.TryGetValue(expected.Name, out int tableIndex))
                {
                    if (expected.IsRequired)
                    {
                        return false;
                    }

                    map[schemaIndex] = -1;
                }
                else
                {
                    map[schemaIndex] = tableIndex;
                }
            }

            columnMap = map;
            return true;
        }

        internal override bool TryReadBody(Excel.ListObject table, out List<object?[]>? body, out int[]? columnMap)
        {
            body = null;
            columnMap = null;
            if (!TryBuildColumnMap(table, out columnMap) || columnMap is null)
            {
                return false;
            }

            Excel.Range? dataBody = _dataBodyRangeAt(table);
            if (dataBody is null)
            {
                return false;
            }

            body = ExcelValue2Matrix.ReadRows(_rangeValue2At(dataBody));
            return true;
        }

        internal override int ResolveActiveBodyIndex(Excel.ListObject table, Excel.ListRows rows)
        {
            Excel.Range? activeCell = _activeCellAt(null!);
            if (activeCell is null || !_tableActiveAt(table))
            {
                return 0;
            }

            int bodyIndex = _rangeRowAt(activeCell) - _rangeRowAt(_tableRangeAt(table));
            return bodyIndex < 1 || bodyIndex > _listRowCountAt(rows) ? 0 : bodyIndex;
        }

        internal override Excel.ListRows GetListRows(Excel.ListObject table) => _listRowsAt(table);

        internal override int GetListColumnCount(Excel.ListColumns columns) => columns.Count;

        internal override Excel.ListColumn GetListColumnAt(Excel.ListColumns columns, int index) => _columnAt(columns, index);

        internal override int GetListRowCount(Excel.ListRows rows) => _listRowCountAt(rows);

        internal override Excel.Range GetTableRange(Excel.ListObject table) => _tableRangeAt(table);

        internal override Excel.Range? GetDataBodyRange(Excel.ListObject table) => _dataBodyRangeAt(table);

        internal override Excel.Range GetRangeRows(Excel.Range range) => _rangeRowsAt(range);

        internal override Excel.Range GetRangeAt(Excel.Range rows, int index) => _rangeAt(rows, index);

        internal override Excel.Range GetCellAt(Excel.Range row, int columnIndex) => _cellAt(row, columnIndex);

        internal override object? GetRangeValue2(Excel.Range range) => _rangeValue2At(range);

        internal override void SetRangeValues(Excel.Range range, object payload) => _setRangeValues(range, payload);

        internal override Excel.Range? GetActiveCell(Excel.Application? application) => _activeCellAt(null!);

        internal override bool GetTableActive(Excel.ListObject table) => _tableActiveAt(table);

        internal override int GetRangeRow(Excel.Range range) => _rangeRowAt(range);
    }

    /// <summary>
    /// Builds the fake-COM graph the adapter drives: a one-row table whose
    /// single body row is an activity, with the active cell on it.
    /// </summary>
    private static (
        TestableMutator Mutator,
        List<(int Row, int Column, object? Value)> WrittenCells,
        Mock<IGanttRowInserter> Inserter) Build(
        string parentType = "As-Built Activity",
        ProtectionGuardOutcome guardOutcome = ProtectionGuardOutcome.NotProtected,
        int childCount = 7)
    {
        List<(int Row, int Column, object? Value)> writtenCells = new();

        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();
        Mock<Excel.Sheets> sheets = new();
        Mock<Excel.Worksheet> worksheet = new();
        Mock<Excel.ListObjects> listObjects = new();
        Mock<Excel.ListObject> table = new();
        Mock<Excel.ListColumns> listColumns = new();
        Mock<Excel.ListColumn>[] columns = new Mock<Excel.ListColumn>[GanttTableSchema.Default.Columns.Count];
        for (int index = 0; index < columns.Length; index++)
        {
            int captured = index;
            columns[index] = new Mock<Excel.ListColumn>();
            columns[index].Setup(c => c.Name).Returns(GanttTableSchema.Default.Columns[captured].Name);
        }

        Mock<Excel.ListRows> listRows = new();
        Mock<Excel.Range> tableRange = new();
        Mock<Excel.Range> dataBodyRange = new();
        Mock<Excel.Range> bodyRows = new();
        Mock<Excel.Range> activeCell = new();

        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.Setup(w => w.Sheets).Returns(sheets.Object);
        _ = worksheet.Setup(w => w.ListObjects).Returns(listObjects.Object);
        _ = table.Setup(t => t.ListColumns).Returns(listColumns.Object);
        _ = table.Setup(t => t.ListRows).Returns(listRows.Object);
        _ = table.Setup(t => t.Range).Returns(tableRange.Object);
        _ = table.Setup(t => t.DataBodyRange).Returns(dataBodyRange.Object);
        _ = listRows.Setup(r => r.Count).Returns(1);
        _ = tableRange.Setup(r => r.Row).Returns(1);
        _ = activeCell.Setup(r => r.Row).Returns(2);
        _ = listColumns.Setup(c => c.Count).Returns(columns.Length);
        _ = listColumns.Setup(c => c[It.IsAny<int>()]).Returns((Excel.ListColumn)columns[0].Object);

        // The body matrix: one row, an activity, with the parent's id.
        // A multidimensional array because Excel's Value2 is a SAFEARRAY.
#pragma warning disable CA1814
        object?[,] matrix = new object?[1, columns.Length];
#pragma warning restore CA1814
        for (int index = 0; index < columns.Length; index++)
        {
            matrix[0, index] = string.Empty;
        }

        int idColumn = GanttTableSchema.Default.Columns
            .Select((column, index) => (column, index))
            .Where(pair => pair.column.Name == "Id")
            .Select(pair => pair.index)
            .First();
        int typeColumn = GanttTableSchema.Default.Columns
            .Select((column, index) => (column, index))
            .Where(pair => pair.column.Name == "Type")
            .Select(pair => pair.index)
            .First();
        int orderColumn = GanttTableSchema.Default.Columns
            .Select((column, index) => (column, index))
            .Where(pair => pair.column.Name == "SiblingOrder")
            .Select(pair => pair.index)
            .First();

        matrix[0, idColumn] = ParentId;
        matrix[0, typeColumn] = parentType;
        matrix[0, orderColumn] = 0;

        _ = dataBodyRange.Setup(r => r.Value2).Returns(matrix);
        _ = bodyRows.Setup(r => r[It.IsAny<int>()]).Returns((Excel.Range)dataBodyRange.Object);
        _ = dataBodyRange.Setup(r => r.Rows).Returns(bodyRows.Object);

        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(guardOutcome);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<Excel.Worksheet>())).Returns(ProtectionGuardOutcome.NotProtected);

        Mock<IGanttRowInserter> inserter = new();
        _ = inserter
            .Setup(i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()))
            .Returns(GanttRowInsertOutcome.Ok(2));

        TestableMutator mutator = new(
            application.Object,
            guard.Object,
            inserter.Object,
            null,
            (sheetsArg, index) => worksheet.Object,
            (listObjectsArg, index) => table.Object,
            (columnsArg, index) => columns[index - 1].Object,
            tableArg => listRows.Object,
            rowsArg => 1,
            tableArg => tableRange.Object,
            tableArg => dataBodyRange.Object,
            rangeArg => bodyRows.Object,
            (rangeArg, index) => dataBodyRange.Object,
            (rangeArg, index) => dataBodyRange.Object,
            rangeArg => matrix,
            (range, payload) =>
            {
                _ = range;
                writtenCells.Add((0, 0, payload));
            },
            applicationArg => activeCell.Object,
            tableArg => true,
            rangeArg => rangeArg == tableRange.Object ? 1 : 2);

        _ = childCount;
        return (mutator, writtenCells, inserter);
    }

    [Fact]
    public void A_protected_sheet_is_refused_before_anything_is_read()
    {
        (TestableMutator mutator, List<(int Row, int Column, object? Value)> written, Mock<IGanttRowInserter> inserter) =
            Build(guardOutcome: ProtectionGuardOutcome.SheetProtected);

        AddChildMutationOutcome outcome = mutator.AddChild(
            GanttEntityType.AsBuiltActivity,
            () => GanttRowId.Parse(ChildId));

        Assert.False(outcome.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(written);
        inserter.Verify(i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()), Times.Never);
    }

    [Fact]
    public void A_parent_that_may_not_own_children_is_refused_before_the_row_is_inserted()
    {
        (TestableMutator mutator, _, Mock<IGanttRowInserter> inserter) =
            Build(parentType: "As-Built Milestone");

        AddChildMutationOutcome outcome = mutator.AddChild(
            GanttEntityType.AsBuiltActivity,
            () => GanttRowId.Parse(ChildId));

        Assert.False(outcome.Succeeded);
        Assert.Equal(AddChildRefusalReason.ParentCannotOwnChildren, outcome.PlanRefusal);
        inserter.Verify(i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()), Times.Never);
    }

    [Fact]
    public void A_type_that_may_not_be_a_child_is_refused_before_the_row_is_inserted()
    {
        (TestableMutator mutator, _, Mock<IGanttRowInserter> inserter) = Build();

        AddChildMutationOutcome outcome = mutator.AddChild(
            GanttEntityType.Spacer,
            () => GanttRowId.Parse(ChildId));

        Assert.False(outcome.Succeeded);
        Assert.Equal(AddChildRefusalReason.TypeCannotBeChild, outcome.PlanRefusal);
        inserter.Verify(i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()), Times.Never);
    }

    [Fact]
    public void A_successful_add_inserts_the_row_and_writes_the_parent_and_sibling_order()
    {
        (TestableMutator mutator, List<(int Row, int Column, object? Value)> written, Mock<IGanttRowInserter> inserter) =
            Build();

        AddChildMutationOutcome outcome = mutator.AddChild(
            GanttEntityType.AsBuiltActivity,
            () => GanttRowId.Parse(ChildId));

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.SelectedRowNumber);
        Assert.Equal(2, outcome.InsertedBodyIndex);
        inserter.Verify(i => i.Insert(GanttEntityType.AsBuiltActivity, It.IsAny<Func<GanttRowId>>()), Times.Once);

        // The two hierarchy cells: the parent id and the sibling order.
        Assert.Equal(2, written.Count);
        Assert.Contains(written, write => write.Item3 as string == ParentId);
        Assert.Contains(written, write => write.Item3 is 0);
    }

    [Fact]
    public void A_refused_insert_is_reported_and_writes_nothing()
    {
        (TestableMutator mutator, List<(int Row, int Column, object? Value)> written, Mock<IGanttRowInserter> inserter) =
            Build();
        _ = inserter
            .Setup(i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()))
            .Returns(GanttRowInsertOutcome.Refused(GanttRowInsertRefusalReason.RowInsertRefused));

        AddChildMutationOutcome outcome = mutator.AddChild(
            GanttEntityType.AsBuiltActivity,
            () => GanttRowId.Parse(ChildId));

        Assert.False(outcome.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.InsertRefused, outcome.Refusal);
        Assert.Empty(written);
    }

    /// <summary>
    /// Every refusal the adapter owns, driven through one parameterised graph so a
    /// new arm cannot ship unexercised. The knobs exist because the shared
    /// <c>Build</c> helper above is deliberately a single happy-path graph; making
    /// it carry ten optional behaviours would obscure the tests that use it.
    /// </summary>
    [Fact]
    public void A_null_id_factory_is_a_programming_error()
    {
        Variant variant = Variant.Build();

        ArgumentNullException thrown = Assert.Throws<ArgumentNullException>(
            () => variant.Mutator.AddChild(GanttEntityType.AsBuiltActivity, null!));

        Assert.Equal("nextId", thrown.ParamName);
        variant.Inserter.Verify(
            i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()), Times.Never);
    }

    [Theory]
    [InlineData(ProtectionGuardOutcome.SheetProtected)]
    [InlineData(ProtectionGuardOutcome.WorkbookStructureProtected)]
    public void Any_active_sheet_refusal_is_reported_as_target_protected(ProtectionGuardOutcome outcome)
    {
        Variant variant = Variant.Build(guardOutcome: outcome);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.TargetProtected, result.Refusal);
        Assert.Empty(variant.WrittenCells);
    }

    [Fact]
    public void An_active_sheet_guard_without_a_workbook_is_reported_as_no_active_workbook()
    {
        Variant variant = Variant.Build(guardOutcome: ProtectionGuardOutcome.NoActiveWorkbook);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.NoActiveWorkbook, result.Refusal);
    }

    [Fact]
    public void An_absent_application_is_refused_as_no_active_workbook()
    {
        Variant variant = Variant.Build(applicationPresent: false);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.NoActiveWorkbook, result.Refusal);
        Assert.Empty(variant.WrittenCells);
    }

    [Fact]
    public void A_workbook_without_the_data_table_is_refused()
    {
        Variant variant = Variant.Build(tablePresent: false);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.TableMissing, result.Refusal);
    }

    [Fact]
    public void A_table_missing_a_required_column_is_refused()
    {
        Variant variant = Variant.Build(completeColumns: false);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.TableMissing, result.Refusal);
    }

    [Fact]
    public void A_protected_target_sheet_is_refused_as_target_protected()
    {
        Variant variant = Variant.Build(targetOutcome: ProtectionGuardOutcome.SheetProtected);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.TargetProtected, result.Refusal);
        variant.Inserter.Verify(
            i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()), Times.Never);
    }

    [Fact]
    public void A_target_guard_without_a_workbook_is_refused_as_no_active_workbook()
    {
        Variant variant = Variant.Build(targetOutcome: ProtectionGuardOutcome.NoActiveWorkbook);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.NoActiveWorkbook, result.Refusal);
    }

    [Fact]
    public void A_selection_on_the_header_row_is_refused_as_outside_the_table()
    {
        Variant variant = Variant.Build(activeRow: 1);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.SelectionOutsideTable, result.Refusal);
    }

    [Fact]
    public void A_selection_below_the_last_body_row_is_refused_as_outside_the_table()
    {
        // bodyIndex resolves to 2 from the active cell, but the body has one row.
        Variant variant = Variant.Build(activeRow: 3, listRowCount: 3);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.SelectionOutsideTable, result.Refusal);
    }

    [Fact]
    public void A_table_body_that_cannot_be_read_is_refused_as_table_missing()
    {
        Variant variant = Variant.Build(bodyPresent: false);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.TableMissing, result.Refusal);
    }

    /// <summary>
    /// A selected row whose Type cell cannot be parsed is refused for exactly that:
    /// the cell is unreadable. It is not reported as a row that cannot own children.
    /// </summary>
    /// <remarks>
    /// The two statements have different remedies. "This row cannot have children"
    /// tells the user to pick a different row, when the row was fine and its Type cell
    /// was not; it also hides a data defect behind a legitimate refusal. The refusal is
    /// the mutator's own (<see cref="AddChildMutationRefusalReason.SelectedRowTypeUnreadable"/>)
    /// rather than a planner one, because the planner never learns that the cell was
    /// unreadable — a projected event can only carry the enum's default.
    /// </remarks>
    [Fact]
    public void A_selected_row_whose_type_cell_is_unreadable_is_refused_as_exactly_that()
    {
        Variant variant = Variant.Build(rowType: "Not A Type");

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.SelectedRowTypeUnreadable, result.Refusal);
        Assert.Null(result.PlanRefusal);
        Assert.Empty(variant.WrittenCells);
        variant.Inserter.Verify(
            i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()), Times.Never);
    }

    [Fact]
    public void A_missing_row_inserter_is_refused_as_insert_refused()
    {
        Variant variant = Variant.Build(inserterPresent: false);

        AddChildMutationOutcome result = variant.Add();

        Assert.False(result.Succeeded);
        Assert.Equal(AddChildMutationRefusalReason.InsertRefused, result.Refusal);
        Assert.Empty(variant.WrittenCells);
    }

    /// <summary>The parameterisable graph the refusal tests drive.</summary>
    private sealed record Variant(
        TestableMutator Mutator,
        List<(int Row, int Column, object? Value)> WrittenCells,
        Mock<IGanttRowInserter> Inserter)
    {
        /// <summary>Runs one add using the child id the tests expect.</summary>
        /// <returns>The outcome.</returns>
        public AddChildMutationOutcome Add() =>
            Mutator.AddChild(GanttEntityType.AsBuiltActivity, () => GanttRowId.Parse(ChildId));

        /// <summary>Builds the graph with one knob per refusal arm.</summary>
        /// <param name="guardOutcome">The active-sheet guard outcome.</param>
        /// <param name="targetOutcome">The target-sheet guard outcome.</param>
        /// <param name="applicationPresent">Whether an application object was supplied.</param>
        /// <param name="tablePresent">Whether the sheet reports the data table.</param>
        /// <param name="completeColumns">Whether every schema column is present.</param>
        /// <param name="bodyPresent">Whether the table reports a body.</param>
        /// <param name="inserterPresent">Whether a row inserter was supplied.</param>
        /// <param name="activeRow">The active cell's worksheet row.</param>
        /// <param name="listRowCount">The table's list-row count.</param>
        /// <param name="rowType">The body row's Type text.</param>
        /// <returns>The variant.</returns>
        public static Variant Build(
            ProtectionGuardOutcome guardOutcome = ProtectionGuardOutcome.NotProtected,
            ProtectionGuardOutcome targetOutcome = ProtectionGuardOutcome.NotProtected,
            bool applicationPresent = true,
            bool tablePresent = true,
            bool completeColumns = true,
            bool bodyPresent = true,
            bool inserterPresent = true,
            int activeRow = 2,
            int listRowCount = 1,
            string rowType = "As-Built Activity")
        {
            List<(int Row, int Column, object? Value)> written = [];
            Mock<Excel.Application> application = new();
            Mock<Excel.Workbook> workbook = new();
            Mock<Excel.Sheets> sheets = new();
            Mock<Excel.Worksheet> worksheet = new();
            Mock<Excel.ListObjects> listObjects = new();
            Mock<Excel.ListObject> table = new();
            Mock<Excel.ListColumns> listColumns = new();
            Mock<Excel.ListRows> listRows = new();
            Mock<Excel.Range> tableRange = new();
            Mock<Excel.Range> dataBodyRange = new();
            Mock<Excel.Range> bodyRows = new();
            Mock<Excel.Range> activeCell = new();
            int columnCount = GanttTableSchema.Default.Columns.Count;
            Mock<Excel.ListColumn>[] columns = new Mock<Excel.ListColumn>[columnCount];
            for (int index = 0; index < columnCount; index++)
            {
                int captured = index;
                columns[index] = new Mock<Excel.ListColumn>();
                _ = columns[index]
                    .Setup(c => c.Name)
                    .Returns(completeColumns
                        ? GanttTableSchema.Default.Columns[captured].Name
                        : "NotAColumn");
            }

            _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
            _ = workbook.Setup(w => w.Sheets).Returns(sheets.Object);
            _ = worksheet.Setup(w => w.ListObjects).Returns(listObjects.Object);
            _ = table.Setup(t => t.ListColumns).Returns(listColumns.Object);
            _ = table.Setup(t => t.ListRows).Returns(listRows.Object);
            _ = table.Setup(t => t.Range).Returns(tableRange.Object);
            _ = table.Setup(t => t.DataBodyRange).Returns(dataBodyRange.Object);
            _ = listRows.Setup(r => r.Count).Returns(listRowCount);
            _ = tableRange.Setup(r => r.Row).Returns(1);
            _ = activeCell.Setup(r => r.Row).Returns(activeRow);
            _ = listColumns.Setup(c => c.Count).Returns(columnCount);

#pragma warning disable CA1814 // COM interop requires a rectangular SAFEARRAY.
            object?[,] matrix = new object?[1, columnCount];
#pragma warning restore CA1814
            for (int index = 0; index < columnCount; index++)
            {
                matrix[0, index] = string.Empty;
            }

            int Column(string name) => GanttTableSchema.Default.Columns
                .Select((column, index) => (column, index))
                .First(pair => pair.column.Name == name).index;

            matrix[0, Column("Id")] = ParentId;
            matrix[0, Column("Type")] = rowType;
            matrix[0, Column("SiblingOrder")] = 0;

            _ = dataBodyRange.Setup(r => r.Value2).Returns(matrix);
            _ = bodyRows.Setup(r => r[It.IsAny<int>()]).Returns((Excel.Range)dataBodyRange.Object);
            _ = dataBodyRange.Setup(r => r.Rows).Returns(bodyRows.Object);

            Mock<IWorksheetProtectionGuard> guard = new();
            _ = guard.Setup(g => g.Query()).Returns(guardOutcome);
            _ = guard.Setup(g => g.QueryTarget(It.IsAny<Excel.Worksheet>())).Returns(targetOutcome);

            Mock<IGanttRowInserter> inserter = new();
            _ = inserter
                .Setup(i => i.Insert(It.IsAny<GanttEntityType>(), It.IsAny<Func<GanttRowId>>()))
                .Returns(GanttRowInsertOutcome.Ok(2));

            TestableMutator mutator = new(
                applicationPresent ? application.Object : null,
                guard.Object,
                inserterPresent ? inserter.Object : null,
                null,
                (sheetsArg, index) => tablePresent ? worksheet.Object : new object(),
                (listObjectsArg, index) => table.Object,
                (columnsArg, index) => columns[index - 1].Object,
                tableArg => listRows.Object,
                rowsArg => listRowCount,
                tableArg => tableRange.Object,
                tableArg => bodyPresent ? dataBodyRange.Object : null,
                rangeArg => bodyRows.Object,
                (rangeArg, index) => dataBodyRange.Object,
                (rangeArg, index) => dataBodyRange.Object,
                rangeArg => matrix,
                (range, payload) => written.Add((0, 0, payload)),
                applicationArg => activeCell.Object,
                tableArg => true,
                rangeArg => rangeArg == tableRange.Object ? 1 : activeRow);

            return new Variant(mutator, written, inserter);
        }
    }
}
