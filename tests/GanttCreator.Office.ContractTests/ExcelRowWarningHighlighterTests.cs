using System.Globalization;
using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelRowWarningHighlighter"/>.
/// </summary>
/// <remarks>
/// The Excel object graph is mocked and only the two COM indexer seams are
/// substituted (<see cref="ExcelRowWarningHighlighter.GetRowRange"/> and
/// <see cref="ExcelRowWarningHighlighter.GetCellRange"/>), because a
/// parameterised COM property cannot appear in a Moq expression tree. Every
/// decision in the adapter therefore runs on the test host: the protection guard
/// ordering, the setup refusals, the body bounds, the tracking set, the reconcile
/// diff and the best-effort persistence. The real indexer behaviour stays with the
/// tagged live-Office integration test.
/// </remarks>
public sealed class ExcelRowWarningHighlighterTests
{
    /// <summary>Pastel red/pink = RGB(255, 200, 200) as OLE_COLOR.</summary>
    private const int _pastelRed = 0x00C8C8FF;

    /// <summary>Pastel amber/yellow = RGB(255, 235, 170) as OLE_COLOR.</summary>
    private const int _pastelAmber = 0x00AAEBFF;

    /// <summary>The OLE colour Excel uses to remove a fill.</summary>
    private const int _cleared = (int)Excel.XlColorIndex.xlColorIndexNone;

    [Fact]
    public void A_row_fill_writes_the_error_colour_once_and_reports_the_column_count()
    {
        Harness harness = new();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(2, GanttValidationSeverity.Error);

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.Equal(4, outcome.NotesWritten);
        Assert.Equal([_pastelRed], harness.RowColours(2));
    }

    [Fact]
    public void A_warning_row_fill_uses_the_pastel_amber_colour()
    {
        Harness harness = new();

        _ = harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Warning);

        Assert.Equal([_pastelAmber], harness.RowColours(1));
    }

    [Fact]
    public void A_severity_outside_the_known_set_clears_the_row_instead_of_filling_it()
    {
        Harness harness = new();
        _ = harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        _ = harness.Highlighter.HighlightRow(1, (GanttValidationSeverity)99);

        Assert.Equal([_pastelRed, _cleared], harness.RowColours(1));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(4)]
    [InlineData(99)]
    public void A_row_outside_the_table_body_is_ignored_without_a_write(int rowNumber)
    {
        Harness harness = new();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(rowNumber, GanttValidationSeverity.Error);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Empty(harness.RowColours(rowNumber));
    }

    [Fact]
    public void A_missing_application_is_refused_before_the_workbook_is_read()
    {
        Harness harness = new(application: false);

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.NoActiveWorkbook, outcome.Refusal);
        harness.Application.VerifyGet(a => a.ActiveWorkbook, Times.Never);
        Assert.Empty(harness.RowColours(1));
    }

    [Fact]
    public void Active_sheet_protection_refuses_before_the_workbook_is_read()
    {
        Harness harness = new();
        harness.ProtectActiveSheet();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TargetProtected, outcome.Refusal);
        harness.Application.VerifyGet(a => a.ActiveWorkbook, Times.Never);
    }

    [Fact]
    public void An_absent_active_workbook_is_refused()
    {
        Harness harness = new();
        harness.ForgetActiveWorkbook();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightCell(1, 1, GanttValidationSeverity.Warning);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void A_workbook_without_the_data_table_is_refused()
    {
        Harness harness = new();
        harness.RenameTable("tblSomethingElse");

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void A_chart_sheet_before_the_data_sheet_is_skipped()
    {
        Harness harness = new();
        harness.SheetIsWorksheet = false;

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void A_workbook_whose_sheet_count_throws_is_refused_as_table_missing()
    {
        Harness harness = new();
        harness.MakeSheetCountThrow();

        GanttValidationReportOutcome outcome = harness.Highlighter.ClearHighlight(1);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void A_protected_target_sheet_is_refused_after_the_table_is_found()
    {
        Harness harness = new();
        harness.ProtectTargetSheet();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(harness.RowColours(1));
    }

    [Fact]
    public void A_target_guard_without_a_workbook_maps_to_the_no_active_workbook_refusal()
    {
        Harness harness = new();
        harness.TargetHasNoWorkbook();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void A_table_without_a_body_treats_every_write_as_a_no_op()
    {
        Harness harness = new();
        harness.RemoveTableBody();

        GanttValidationReportOutcome row =
            harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);
        GanttValidationReportOutcome cell =
            harness.Highlighter.HighlightCell(1, 1, GanttValidationSeverity.Error);
        GanttValidationReportOutcome clear = harness.Highlighter.ClearHighlight(1);
        GanttValidationReportOutcome reconcile = harness.Highlighter.ReconcileHighlights(
            [new HighlightTarget(1, null, GanttValidationSeverity.Error)]);

        Assert.Equal(0, row.NotesWritten);
        Assert.Equal(0, cell.NotesWritten);
        Assert.Equal(0, clear.NotesWritten);
        Assert.True(reconcile.Succeeded);
        Assert.Equal(0, reconcile.NotesWritten);
        Assert.Empty(harness.RowColours(1));
        Assert.Empty(harness.CellColours(1, 1));
    }

    [Fact]
    public void A_cell_fill_writes_the_pastel_colour_and_reports_one_cell()
    {
        Harness harness = new();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightCell(2, 3, GanttValidationSeverity.Warning);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.NotesWritten);
        Assert.Equal([_pastelAmber], harness.CellColours(2, 3));
        Assert.Empty(harness.RowColours(2));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(4, 1)]
    [InlineData(1, 5)]
    public void A_cell_outside_the_table_body_is_ignored_without_a_write(int rowNumber, int column)
    {
        Harness harness = new();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightCell(rowNumber, column, GanttValidationSeverity.Error);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Empty(harness.CellColours(rowNumber, column));
    }

    [Fact]
    public void A_cell_range_that_cannot_be_reached_is_not_counted_or_tracked()
    {
        Harness harness = new();
        _ = harness.MissingCells.Add((2, 2));

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightCell(2, 2, GanttValidationSeverity.Error);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);

        // Still not tracked: a later reconcile retries the same unreachable cell,
        // because the sheet state is unknown rather than known-cleared.
        GanttValidationReportOutcome reconcile = harness.Highlighter.ReconcileHighlights(
            [new HighlightTarget(2, 2, GanttValidationSeverity.Error)]);

        Assert.Equal(0, reconcile.NotesWritten);
    }

    [Fact]
    public void Clearing_a_single_cell_writes_none_and_reports_one_cell()
    {
        Harness harness = new();
        _ = harness.Highlighter.HighlightCell(1, 2, GanttValidationSeverity.Error);

        GanttValidationReportOutcome outcome = harness.Highlighter.ClearHighlight(1, 2);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.NotesWritten);
        Assert.Equal([_pastelRed, _cleared], harness.CellColours(1, 2));
    }

    [Fact]
    public void Clearing_a_single_cell_outside_the_body_reports_nothing()
    {
        Harness harness = new();

        GanttValidationReportOutcome outcome = harness.Highlighter.ClearHighlight(9, 9);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Empty(harness.CellColours(9, 9));
    }

    [Fact]
    public void Clearing_a_single_cell_that_cannot_be_reached_reports_nothing()
    {
        Harness harness = new();
        _ = harness.MissingCells.Add((1, 1));

        GanttValidationReportOutcome outcome = harness.Highlighter.ClearHighlight(1, 1);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
    }

    [Fact]
    public void Clearing_a_row_writes_none_across_the_row_and_reports_the_column_count()
    {
        Harness harness = new();
        _ = harness.Highlighter.HighlightRow(2, GanttValidationSeverity.Error);

        GanttValidationReportOutcome outcome = harness.Highlighter.ClearHighlight(2);

        Assert.True(outcome.Succeeded);
        Assert.Equal(4, outcome.NotesWritten);
        Assert.Equal([_pastelRed, _cleared], harness.RowColours(2));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(4)]
    public void Clearing_a_row_outside_the_body_reports_nothing(int rowNumber)
    {
        Harness harness = new();

        GanttValidationReportOutcome outcome = harness.Highlighter.ClearHighlight(rowNumber);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Empty(harness.RowColours(rowNumber));
    }

    [Fact]
    public void Reconcile_refuses_a_null_target_list()
    {
        Harness harness = new();

        _ = Assert.Throws<ArgumentNullException>(
            () => harness.Highlighter.ReconcileHighlights(null!));
    }

    [Fact]
    public void Reconcile_refuses_when_the_active_sheet_is_protected()
    {
        Harness harness = new();
        harness.ProtectActiveSheet();

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TargetProtected, outcome.Refusal);
    }

    [Fact]
    public void Reconcile_refuses_when_the_workbook_has_no_data_table()
    {
        Harness harness = new();
        harness.RenameTable("tblSomethingElse");

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void Reconcile_clears_a_row_that_is_no_longer_a_target()
    {
        Harness harness = new();
        _ = harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(4, outcome.NotesWritten);
        Assert.Equal([_pastelRed, _cleared], harness.RowColours(1));
    }

    [Fact]
    public void Reconcile_clears_a_single_cell_that_is_no_longer_a_target()
    {
        Harness harness = new();
        _ = harness.Highlighter.HighlightCell(1, 2, GanttValidationSeverity.Error);

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([]);

        Assert.Equal(1, outcome.NotesWritten);
        Assert.Equal([_pastelRed, _cleared], harness.CellColours(1, 2));
    }

    [Fact]
    public void Reconcile_leaves_rows_outside_the_supplied_scope_untouched()
    {
        Harness harness = new();
        _ = harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([], [2]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Equal([_pastelRed], harness.RowColours(1));
    }

    [Fact]
    public void Reconcile_fills_a_new_row_target_and_a_new_cell_target()
    {
        Harness harness = new();

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights(
        [
            new HighlightTarget(1, null, GanttValidationSeverity.Error),
            new HighlightTarget(3, 2, GanttValidationSeverity.Warning),
        ]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(5, outcome.NotesWritten);
        Assert.Equal([_pastelRed], harness.RowColours(1));
        Assert.Equal([_pastelAmber], harness.CellColours(3, 2));
    }

    [Fact]
    public void Reconcile_skips_a_target_outside_the_body_bounds()
    {
        Harness harness = new();

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights(
        [
            new HighlightTarget(0, null, GanttValidationSeverity.Error),
            new HighlightTarget(9, null, GanttValidationSeverity.Error),
            new HighlightTarget(1, 9, GanttValidationSeverity.Error),
        ]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
    }

    [Fact]
    public void Reconcile_does_not_rewrite_a_target_it_already_highlights()
    {
        Harness harness = new();
        _ = harness.Highlighter.ReconcileHighlights(
            [new HighlightTarget(1, null, GanttValidationSeverity.Error)]);

        GanttValidationReportOutcome second = harness.Highlighter.ReconcileHighlights(
            [new HighlightTarget(1, null, GanttValidationSeverity.Error)]);

        Assert.True(second.Succeeded);
        Assert.Equal(0, second.NotesWritten);
        Assert.Equal([_pastelRed], harness.RowColours(1));
    }

    [Fact]
    public void Reconcile_clears_a_persisted_row_and_cell_from_the_workbook_name()
    {
        Harness harness = new();
        harness.PersistedHighlights("=\"1|;2|3\"");

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(5, outcome.NotesWritten);
        Assert.Equal([_cleared], harness.RowColours(1));
        Assert.Equal([_cleared], harness.CellColours(2, 3));
    }

    [Fact]
    public void A_malformed_persisted_entry_is_skipped_without_blocking_the_rest()
    {
        Harness harness = new();
        harness.PersistedHighlights("=\"bogus;notanumber|x;2|3\"");

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.NotesWritten);
        Assert.Equal([_cleared], harness.CellColours(2, 3));
        Assert.Empty(harness.RowColours(1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("=\"\"")]
    public void An_empty_persisted_value_loads_no_highlights(string refersTo)
    {
        Harness harness = new();
        harness.PersistedHighlights(refersTo);

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
    }

    [Fact]
    public void An_unreadable_persisted_name_does_not_block_highlighting()
    {
        Harness harness = new();
        harness.PersistedNameThrows();

        GanttValidationReportOutcome outcome =
            harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);

        Assert.True(outcome.Succeeded);
        Assert.Equal([_pastelRed], harness.RowColours(1));
    }

    [Fact]
    public void A_persisted_cell_outside_the_body_is_left_alone()
    {
        Harness harness = new();
        harness.PersistedHighlights("=\"1|9\"");

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights([]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Empty(harness.CellColours(1, 9));
    }

    [Fact]
    public void A_persisted_cell_highlight_reloads_as_a_cell_not_a_whole_row()
    {
        Harness harness = new();
        harness.PersistedHighlights("=\"2|3\"");

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights(
            [new HighlightTarget(2, 3, GanttValidationSeverity.Warning)]);

        // Already tracked, so nothing is written - and crucially no whole-row
        // clear or fill happens on row 2. Reading the column from the wrong field
        // used to turn this into a four-cell row write.
        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Empty(harness.RowColours(2));
    }

    [Fact]
    public void The_tracked_highlights_are_persisted_back_to_the_workbook_name()
    {
        Harness harness = new();
        harness.NameExists();
        _ = harness.Highlighter.HighlightRow(1, GanttValidationSeverity.Error);
        _ = harness.Highlighter.HighlightCell(2, 3, GanttValidationSeverity.Warning);

        _ = harness.Highlighter.ReconcileHighlights(
        [
            new HighlightTarget(1, null, GanttValidationSeverity.Error),
            new HighlightTarget(2, 3, GanttValidationSeverity.Warning),
        ]);

        Assert.Equal("=\"1|;2|3\"", harness.Name.Object.RefersTo);
    }

    [Fact]
    public void A_persistence_failure_does_not_fail_the_reconcile()
    {
        Harness harness = new();
        harness.PersistedNameThrows();

        GanttValidationReportOutcome outcome = harness.Highlighter.ReconcileHighlights(
            [new HighlightTarget(1, null, GanttValidationSeverity.Error)]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(4, outcome.NotesWritten);
    }

    /// <summary>
    /// The adapter with its COM indexer seams substituted. Every indexer is a
    /// parameterised COM property, which cannot be expressed in a Moq setup, so the
    /// production seams exist for exactly this substitution.
    /// </summary>
    private sealed class TestableHighlighter(
        Harness harness,
        object? application,
        IWorksheetProtectionGuard guard)
        : ExcelRowWarningHighlighter(application, guard)
    {
        internal override object GetSheetAt(Excel.Sheets sheets, int index) =>
            harness.SheetAt(index);

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
            harness.TableAt(index);

        internal override Excel.Range GetRowRange(Excel.Range body, int rowNumber) =>
            harness.RowRange(body, rowNumber);

        internal override Excel.Range? GetCellRange(Excel.Range? body, int rowNumber, int columnIndex) =>
            harness.CellRange(rowNumber, columnIndex);
    }

    /// <summary>
    /// The mocked Excel graph plus the colour sinks the seam substitution writes to.
    /// </summary>
    private sealed class Harness
    {
        private readonly Dictionary<int, List<int>> _rowWrites = [];
        private readonly Dictionary<(int Row, int Column), List<int>> _cellWrites = [];
        private readonly Dictionary<int, Excel.Range> _rowRanges = [];
        private readonly Dictionary<(int Row, int Column), Excel.Range?> _cellRanges = [];

        public Harness(bool application = true, int rowCount = 3, int columnCount = 4)
        {
            Mock<Excel.Range> bodyRows = new();
            Mock<Excel.Range> bodyColumns = new();
            _ = bodyRows.SetupGet(r => r.Count).Returns(rowCount);
            _ = bodyColumns.SetupGet(r => r.Count).Returns(columnCount);
            _ = Body.SetupGet(r => r.Rows).Returns(bodyRows.Object);
            _ = Body.SetupGet(r => r.Columns).Returns(bodyColumns.Object);
            _ = Table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = Table.SetupGet(t => t.DataBodyRange).Returns(Body.Object);
            _ = ListObjects.SetupGet(l => l.Count).Returns(1);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = Worksheet.SetupGet(w => w.Parent).Returns(Workbook.Object);
            _ = Sheets.SetupGet(s => s.Count).Returns(1);
            _ = Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            _ = Workbook.SetupGet(w => w.Names).Returns(Names.Object);
            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            _ = Guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
            _ = Guard.Setup(g => g.QueryTarget(It.IsAny<object?>()))
                .Returns(ProtectionGuardOutcome.NotProtected);

            Name = new Mock<Excel.Name>();
            _ = Name.SetupProperty(n => n.RefersTo);
            Highlighter = new TestableHighlighter(
                this,
                application ? Application.Object : null,
                Guard.Object);
        }

        /// <summary>The mocked application.</summary>
        public Mock<Excel.Application> Application { get; } = new();

        /// <summary>The mocked active workbook.</summary>
        public Mock<Excel.Workbook> Workbook { get; } = new();

        /// <summary>The mocked sheet collection.</summary>
        public Mock<Excel.Sheets> Sheets { get; } = new();

        /// <summary>The mocked data worksheet.</summary>
        public Mock<Excel.Worksheet> Worksheet { get; } = new();

        /// <summary>The mocked table collection on the worksheet.</summary>
        public Mock<Excel.ListObjects> ListObjects { get; } = new();

        /// <summary>The mocked data table.</summary>
        public Mock<Excel.ListObject> Table { get; } = new();

        /// <summary>The mocked table body.</summary>
        public Mock<Excel.Range> Body { get; } = new();

        /// <summary>The mocked workbook name collection.</summary>
        public Mock<Excel.Names> Names { get; } = new();

        /// <summary>The mocked persisted-highlight name.</summary>
        public Mock<Excel.Name> Name { get; }

        /// <summary>The substituted protection guard.</summary>
        public Mock<IWorksheetProtectionGuard> Guard { get; } = new();

        /// <summary>The adapter under test.</summary>
        public TestableHighlighter Highlighter { get; }

        /// <summary>Cell addresses whose range cannot be reached.</summary>
        public HashSet<(int Row, int Column)> MissingCells { get; } = [];

        /// <summary>The colour values written to one body row.</summary>
        /// <param name="rowNumber">The one-based body row.</param>
        /// <returns>The writes, in order.</returns>
        public List<int> RowColours(int rowNumber)
        {
            if (!_rowWrites.TryGetValue(rowNumber, out List<int>? writes))
            {
                writes = [];
                _rowWrites[rowNumber] = writes;
            }

            return writes;
        }

        /// <summary>The colour values written to one body cell.</summary>
        /// <param name="rowNumber">The one-based body row.</param>
        /// <param name="column">The one-based table column.</param>
        /// <returns>The writes, in order.</returns>
        public List<int> CellColours(int rowNumber, int column)
        {
            if (!_cellWrites.TryGetValue((rowNumber, column), out List<int>? writes))
            {
                writes = [];
                _cellWrites[(rowNumber, column)] = writes;
            }

            return writes;
        }

        /// <summary>Whether the first sheet is a worksheet rather than a chart sheet.</summary>
        public bool SheetIsWorksheet { get; set; } = true;

        /// <summary>Substitutes the sheet-index seam.</summary>
        /// <param name="index">The one-based sheet index.</param>
        /// <returns>The worksheet, or an opaque non-worksheet entry.</returns>
        public object SheetAt(int index) =>
            SheetIsWorksheet || index != 1 ? Worksheet.Object : new object();

        /// <summary>Substitutes the table-index seam.</summary>
        /// <param name="index">The one-based table index.</param>
        /// <returns>The data table.</returns>
        public Excel.ListObject TableAt(int index) => index == 1
            ? Table.Object
            : throw new ArgumentOutOfRangeException(nameof(index), "Only one table is configured.");

        /// <summary>Substitutes the row-range seam.</summary>
        /// <param name="body">The table body the adapter resolved.</param>
        /// <param name="rowNumber">The one-based body row.</param>
        /// <returns>The colour sink for that row.</returns>
        public Excel.Range RowRange(Excel.Range body, int rowNumber)
        {
            Assert.Same(Body.Object, body);
            if (!_rowRanges.TryGetValue(rowNumber, out Excel.Range? range))
            {
                range = ColourSink(RowColours(rowNumber));
                _rowRanges[rowNumber] = range;
            }

            return range;
        }

        /// <summary>Substitutes the single-cell seam.</summary>
        /// <param name="rowNumber">The one-based body row.</param>
        /// <param name="column">The one-based table column.</param>
        /// <returns>The colour sink, or <see langword="null"/> when unreachable.</returns>
        public Excel.Range? CellRange(int rowNumber, int column)
        {
            if (!_cellRanges.TryGetValue((rowNumber, column), out Excel.Range? range))
            {
                range = MissingCells.Contains((rowNumber, column))
                    ? null
                    : ColourSink(CellColours(rowNumber, column));
                _cellRanges[(rowNumber, column)] = range;
            }

            return range;
        }

        /// <summary>Makes the workbook's active-sheet query report protection.</summary>
        public void ProtectActiveSheet() => _ = Guard
            .Setup(g => g.Query())
            .Returns(ProtectionGuardOutcome.SheetProtected);

        /// <summary>Makes the target-sheet query report protection.</summary>
        public void ProtectTargetSheet() => _ = Guard
            .Setup(g => g.QueryTarget(It.IsAny<object?>()))
            .Returns(ProtectionGuardOutcome.WorkbookStructureProtected);

        /// <summary>Makes the target-sheet query report an absent workbook.</summary>
        public void TargetHasNoWorkbook() => _ = Guard
            .Setup(g => g.QueryTarget(It.IsAny<object?>()))
            .Returns(ProtectionGuardOutcome.NoActiveWorkbook);

        /// <summary>Removes the workbook from the application.</summary>
        public void ForgetActiveWorkbook() => _ = Application
            .SetupGet(a => a.ActiveWorkbook)
            .Returns((Excel.Workbook)null!);

        /// <summary>Renames the only table so the lookup cannot match.</summary>
        /// <param name="name">The replacement name.</param>
        public void RenameTable(string name) => _ = Table.SetupGet(t => t.Name).Returns(name);

        /// <summary>Removes the table body, as an empty table reports none.</summary>
        public void RemoveTableBody() => _ = Table
            .SetupGet(t => t.DataBodyRange)
            .Returns((Excel.Range)null!);

        /// <summary>Makes the sheet count read throw, as a dead COM proxy would.</summary>
        public void MakeSheetCountThrow() => _ = Sheets
            .SetupGet(s => s.Count)
            .Throws(new InvalidOperationException("COM proxy is dead."));

        /// <summary>Makes the workbook name resolve to the persisted-highlight name.</summary>
        public void NameExists() => _ = Names
            .Setup(n => n.Item(It.IsAny<object>()))
            .Returns(Name.Object);

        /// <summary>Reports a persisted highlight list through the workbook name.</summary>
        /// <param name="refersTo">The stored <c>RefersTo</c> payload.</param>
        public void PersistedHighlights(string refersTo)
        {
            NameExists();
            _ = Name.SetupGet(n => n.RefersTo).Returns(refersTo);
        }

        /// <summary>Makes reading the workbook name throw, as a missing name would.</summary>
        public void PersistedNameThrows() => _ = Names
            .Setup(n => n.Item(It.IsAny<object>()))
            .Throws(new InvalidOperationException("The name is not defined."));

        private static Excel.Range ColourSink(List<int> writes)
        {
            Mock<Excel.Interior> interior = new();
            _ = interior
                .SetupSet(i => i.Color = It.IsAny<object>())
                .Callback<object>(value => writes.Add(Convert.ToInt32(value, CultureInfo.InvariantCulture)));
            Mock<Excel.Range> range = new();
            _ = range.SetupGet(r => r.Interior).Returns(interior.Object);
            return range.Object;
        }
    }
}
