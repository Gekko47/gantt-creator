using System.Collections;
using System.Globalization;
using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelWarningHighlighter"/>.
/// </summary>
/// <remarks>
/// The Excel graph is mocked and the three COM parameterised-property seams
/// (<c>GetSheetAt</c>, both <c>GetTableAt</c> overloads, <c>GetCellRange</c>) are
/// substituted, so the protection ordering, the table and column-map refusals, the
/// warning-only filter, the per-cell highlight and the stale-fill sweep all run on
/// the test host. The sweep's COM enumerator is stubbed the same way the duration
/// writer's table collection is.
/// </remarks>
public sealed class ExcelWarningHighlighterTests
{
    private static readonly uint _warningFill = GanttCatalogues.WarningFill.ARGB;

    [Fact]
    public void A_missing_application_is_refused_before_any_workbook_read()
    {
        Fixture fixture = new() { ApplicationPresent = false };

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.NoActiveWorkbook, outcome.Refusal);
        fixture.VerifyNoWorkbookRead();
    }

    [Fact]
    public void Active_sheet_protection_refuses_before_the_workbook_is_read()
    {
        Fixture fixture = new();
        fixture.ProtectActiveSheet();

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TargetProtected, outcome.Refusal);
        fixture.VerifyNoWorkbookRead();
    }

    [Fact]
    public void An_absent_active_workbook_is_refused_as_no_active_workbook()
    {
        Fixture fixture = new();
        fixture.ProtectActiveSheet(ProtectionGuardOutcome.NoActiveWorkbook);

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void A_workbook_without_the_data_table_is_refused()
    {
        Fixture fixture = new();
        fixture.RemoveTables();

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void A_protected_target_sheet_is_refused_as_target_protected()
    {
        Fixture fixture = new();
        fixture.ProtectTarget();

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TargetProtected, outcome.Refusal);
    }

    [Fact]
    public void A_target_guard_without_a_workbook_is_refused_as_no_active_workbook()
    {
        Fixture fixture = new();
        fixture.ProtectTarget(ProtectionGuardOutcome.NoActiveWorkbook);

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void A_table_missing_a_required_column_is_refused_as_table_missing()
    {
        GanttTableColumn required = GanttTableSchema.Default.Columns.First(column => column.IsRequired);
        Fixture fixture = new(without: required.Name);

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
    }

    [Fact]
    public void A_warning_on_a_mapped_column_fills_that_cell_once()
    {
        Fixture fixture = new();

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 2));

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.NotesWritten);
        Assert.Equal(_warningFill, fixture.FillOf(2, "Start"));
    }

    [Fact]
    public void An_error_issue_is_left_to_the_refusal_path()
    {
        Fixture fixture = new();

        GanttValidationReportOutcome outcome = fixture.Highlight(
            new GanttValidationIssue(1, "Start", "InvalidStart", GanttValidationSeverity.Error, "bad"));

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Null(fixture.FillOf(1, "Start"));
    }

    [Fact]
    public void A_warning_on_a_field_the_schema_does_not_define_is_skipped()
    {
        Fixture fixture = new();

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("NotAColumn", 1));

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
    }

    [Fact]
    public void A_warning_on_a_column_the_table_omits_is_skipped()
    {
        string optional = GanttTableSchema.Default.Columns.Last(column => !column.IsRequired).Name;
        Fixture fixture = new(without: optional);

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning(optional, 1));

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
    }

    [Fact]
    public void A_warning_whose_cell_cannot_be_reached_is_skipped()
    {
        Fixture fixture = new();
        fixture.MissingCells.Add((1, "Start"));

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
    }

    [Fact]
    public void A_cell_whose_fill_write_fails_is_not_counted()
    {
        Fixture fixture = new();
        fixture.FailingCells.Add((1, "Start"));

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
    }

    [Fact]
    public void An_empty_warning_list_still_clears_stale_fills()
    {
        Fixture fixture = new();
        fixture.BodyCells.Add(("Start", 1));

        GanttValidationReportOutcome outcome = fixture.Highlight();

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Equal([(1, "Start")], fixture.Cleared);
    }

    [Fact]
    public void The_sweep_clears_only_cells_carrying_the_warning_fill()
    {
        Fixture fixture = new();
        fixture.BodyCells.Add(("Start", 1));
        fixture.BodyCells.Add(("Finish", 1));
        fixture.OtherFillCells.Add((1, "Finish"));

        _ = fixture.Highlight();

        Assert.Equal([(1, "Start")], fixture.Cleared);
    }

    [Fact]
    public void A_body_enumerator_that_fails_does_not_block_highlighting()
    {
        Fixture fixture = new();
        fixture.BodyCellEnumerationFails = true;

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.NotesWritten);
        Assert.Equal(_warningFill, fixture.FillOf(1, "Start"));
    }

    [Fact]
    public void A_table_without_a_body_still_reports_success()
    {
        Fixture fixture = new();
        fixture.RemoveBody();

        GanttValidationReportOutcome outcome = fixture.Highlight(Warning("Start", 1));

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Null(fixture.FillOf(1, "Start"));
    }

    private static GanttValidationIssue Warning(string field, int row) =>
        new(row, field, "MissingValue", GanttValidationSeverity.Warning, "warning");

    /// <summary>The adapter with its COM indexer seams substituted.</summary>
    private sealed class TestableHighlighter(
        Fixture fixture,
        object? application,
        IWorksheetProtectionGuard guard)
        : ExcelWarningHighlighter(application, guard)
    {
        internal override object GetSheetAt(Excel.Sheets sheets, int index) => fixture.WorksheetObject;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
            fixture.TableObject;

        internal override Excel.ListColumn GetTableAt(Excel.ListColumns columns, int index) =>
            fixture.ColumnAt(index);

        internal override Excel.Range? GetCellRange(Excel.Range? body, int rowNumber, int columnIndex) =>
            fixture.CellRange(rowNumber, columnIndex);
    }

    /// <summary>The mocked Excel graph, the body-cell sweep and the write recorder.</summary>
    private sealed class Fixture
    {
        private readonly string[] _columns;
        private readonly Dictionary<(int Row, int Column), Excel.Range> _cells = [];

        public Fixture(string? without = null)
        {
            _columns = [.. GanttTableSchema.Default.Columns
                .Select(column => column.Name)
                .Where(name => !string.Equals(name, without, StringComparison.Ordinal))];

            _ = Columns.SetupGet(c => c.Count).Returns(_columns.Length);
            _ = Table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = Table.SetupGet(t => t.ListColumns).Returns(Columns.Object);
            _ = ListObjects.SetupGet(l => l.Count).Returns(1);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = Sheets.SetupGet(s => s.Count).Returns(1);
            _ = Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            _ = Guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
            _ = Guard.Setup(g => g.QueryTarget(It.IsAny<object?>()))
                .Returns(ProtectionGuardOutcome.NotProtected);
        }

        /// <summary>Cells whose fill write throws.</summary>
        public HashSet<(int Row, string Column)> FailingCells { get; } = [];

        /// <summary>Cells whose range cannot be reached.</summary>
        public HashSet<(int Row, string Column)> MissingCells { get; } = [];

        /// <summary>Cells the body sweep enumerates.</summary>
        public List<(string Column, int Row)> BodyCells { get; } = [];

        /// <summary>Body cells that carry a fill other than the warning one.</summary>
        public HashSet<(int Row, string Column)> OtherFillCells { get; } = [];

        /// <summary>The cells the sweep cleared, in visit order.</summary>
        public List<(int Row, string Column)> Cleared { get; } = [];

        /// <summary>Whether enumerating the body cells throws, as a dead proxy would.</summary>
        public bool BodyCellEnumerationFails { get; set; }

        /// <summary>Whether the host supplied an Excel application object.</summary>
        public bool ApplicationPresent { get; set; } = true;

        /// <summary>The mocked worksheet.</summary>
        public Excel.Worksheet WorksheetObject => Worksheet.Object;

        /// <summary>The mocked data table.</summary>
        public Excel.ListObject TableObject => Table.Object;

        private Mock<Excel.Application> Application { get; } = new();

        private Mock<Excel.Workbook> Workbook { get; } = new();

        private Mock<Excel.Sheets> Sheets { get; } = new();

        private Mock<Excel.Worksheet> Worksheet { get; } = new();

        private Mock<Excel.ListObjects> ListObjects { get; } = new();

        private Mock<Excel.ListObject> Table { get; } = new();

        private Mock<Excel.ListColumns> Columns { get; } = new();

        private Mock<IWorksheetProtectionGuard> Guard { get; } = new();

        private bool _bodyPresent = true;

        /// <summary>Runs the highlighter.</summary>
        /// <param name="issues">The validation issues.</param>
        /// <returns>The outcome.</returns>
        public GanttValidationReportOutcome Highlight(params GanttValidationIssue[] issues)
        {
            WireBody();
            return new TestableHighlighter(
                this,
                ApplicationPresent ? Application.Object : null,
                Guard.Object).HighlightWarnings(issues);
        }

        /// <summary>The fills applied to cells, keyed by body row and column name.</summary>
        public Dictionary<(int Row, string Column), uint> Fills { get; } = [];

        /// <summary>The fill applied to a cell, or <see langword="null"/> when untouched.</summary>
        /// <param name="row">The one-based body row.</param>
        /// <param name="column">The column name.</param>
        /// <returns>The applied colour.</returns>
        public uint? FillOf(int row, string column) =>
            Fills.TryGetValue((row, column), out uint fill) ? fill : null;

        /// <summary>Asserts the workbook was never read.</summary>
        public void VerifyNoWorkbookRead() =>
            Application.VerifyGet(a => a.ActiveWorkbook, Times.Never);

        /// <summary>Substitutes the column-index seam.</summary>
        /// <param name="index">The one-based column index.</param>
        /// <returns>The mocked list column.</returns>
        public Excel.ListColumn ColumnAt(int index)
        {
            Mock<Excel.ListColumn> column = new();
            _ = column.SetupGet(c => c.Name).Returns(_columns[index - 1]);
            return column.Object;
        }

        /// <summary>Substitutes the single-cell seam.</summary>
        /// <param name="row">The one-based body row.</param>
        /// <param name="column">The one-based column.</param>
        /// <returns>The mocked cell, or <see langword="null"/> when unreachable.</returns>
        public Excel.Range? CellRange(int row, int column)
        {
            string name = _columns[column - 1];
            if (MissingCells.Contains((row, name)))
            {
                return null;
            }

            if (!_cells.TryGetValue((row, column), out Excel.Range? cell))
            {
                cell = BuildCell(row, name);
                _cells[(row, column)] = cell;
            }

            return cell;
        }

        /// <summary>Makes the worksheet report no tables.</summary>
        public void RemoveTables() => _ = ListObjects.SetupGet(l => l.Count).Returns(0);

        /// <summary>Makes the table report no body.</summary>
        public void RemoveBody() => _bodyPresent = false;

        /// <summary>Makes the active-sheet query report the supplied outcome.</summary>
        /// <param name="outcome">The guard outcome.</param>
        public void ProtectActiveSheet(ProtectionGuardOutcome outcome = ProtectionGuardOutcome.SheetProtected) =>
            _ = Guard.Setup(g => g.Query()).Returns(outcome);

        /// <summary>Makes the target-sheet query report the supplied outcome.</summary>
        /// <param name="outcome">The guard outcome.</param>
        public void ProtectTarget(ProtectionGuardOutcome outcome = ProtectionGuardOutcome.SheetProtected) =>
            _ = Guard.Setup(g => g.QueryTarget(It.IsAny<object?>())).Returns(outcome);

        private Excel.Range BuildCell(int row, string name)
        {
            Mock<Excel.Interior> interior = new();
            if (FailingCells.Contains((row, name)))
            {
                _ = interior.SetupSet(i => i.Color = It.IsAny<object>())
                    .Throws(new InvalidOperationException("COM proxy is dead."));
            }
            else
            {
                _ = interior.SetupSet(i => i.Color = It.IsAny<object>())
                    .Callback<object>(value =>
                        Fills[(row, name)] = Convert.ToUInt32(value, CultureInfo.InvariantCulture));
            }

            Mock<Excel.Range> mock = new();
            _ = mock.SetupGet(r => r.Interior).Returns(interior.Object);
            return mock.Object;
        }

        private void WireBody()
        {
            if (!_bodyPresent)
            {
                _ = Table.SetupGet(t => t.DataBodyRange).Returns((Excel.Range)null!);
                return;
            }

            Mock<Excel.Range> cellsCollection = new();
            List<Excel.Range> cells = [];
            foreach ((string column, int row) in BodyCells)
            {
                Mock<Excel.Interior> interior = new();
                _ = interior.SetupGet(i => i.Color).Returns(
                    (object)(OtherFillCells.Contains((row, column)) ? 0x00112233u : _warningFill));
                _ = interior.SetupSet(i => i.ColorIndex = It.IsAny<object>())
                    .Callback<object>(_ => Cleared.Add((row, column)));
                Mock<Excel.Range> cell = new();
                _ = cell.SetupGet(r => r.Interior).Returns(interior.Object);
                cells.Add(cell.Object);
            }

            if (BodyCellEnumerationFails)
            {
                _ = cellsCollection.Setup(r => r.GetEnumerator())
                    .Throws(new InvalidOperationException("COM proxy is dead."));
            }
            else
            {
                _ = cellsCollection.Setup(r => r.GetEnumerator())
                    .Returns(() => ((IEnumerable)cells).GetEnumerator());
            }

            Mock<Excel.Range> body = new();
            _ = body.SetupGet(b => b.Cells).Returns(cellsCollection.Object);
            _ = Table.SetupGet(t => t.DataBodyRange).Returns(body.Object);
        }
    }
}
