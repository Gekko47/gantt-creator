using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelGanttValidationReporter"/> — the
/// automated half of the R2.6 gate (docs/03-ROADMAP.md): the fake-worksheet
/// mutation set, the deterministic error order, and note ownership. No live
/// Excel is required; the COM seams are substituted exactly like
/// <see cref="GanttTableReaderTests"/>.
/// </summary>
/// <remarks>
/// AGENTS.md validator rule: every refusal path in <c>Report</c> is exercised by
/// a positive test that constructs the bad input and asserts the typed refusal.
/// </remarks>
public class ValidationReporterTests
{
    private static List<string> SchemaHeaders()
    {
        var headers = new List<string>();
        foreach (GanttTableColumn column in GanttTableSchema.Default.Columns)
        {
            headers.Add(column.Name);
        }

        return headers;
    }

    /// <summary>One body cell, its captured note writes, and its note state.</summary>
    private sealed class CellGraph
    {
        public CellGraph(int row, int column)
        {
            Row = row;
            Column = column;
            _ = Cell.Setup(c => c.AddComment(It.IsAny<object>()))
                .Callback<object>(text => NotesWritten.Add((string)text));
        }

        public int Row { get; }

        public int Column { get; }

        public Mock<Excel.Range> Cell { get; } = new();

        /// <summary>Texts passed to <c>Range.AddComment</c> (a fresh note).</summary>
        public List<string> NotesWritten { get; } = [];

        public Mock<Excel.Comment>? ExistingComment { get; private set; }

        /// <summary>
        /// Texts written through <c>Comment.Text(value)</c>. A read
        /// (<c>Text()</c>) passes <c>Type.Missing</c> and is not recorded.
        /// </summary>
        public IReadOnlyList<string> NoteTextWrites => ExistingComment is null
            ? []
            : [.. ExistingComment.Invocations
                .Where(i => i.Method.Name == "Text" && i.Arguments[0] is string)
                .Select(i => (string)i.Arguments[0]!)];

        /// <summary>How many times the pre-existing note was deleted.</summary>
        public int NoteDeleteCount => ExistingComment is null
            ? 0
            : ExistingComment.Invocations.Count(i => i.Method.Name == "Delete");

        /// <summary>Puts a pre-existing note on the cell that the add-in does not own.</summary>
        public void WithUserNote(string text) => WithNote(text);

        /// <summary>Puts a pre-existing note on the cell that carries the add-in marker.</summary>
        public void WithOwnedNote(string text) =>
            WithNote(ValidationReportComposer.NoteSentinelPrefix + " " + text);

        /// <summary>
        /// Puts a note on the cell. Reads return the original text; writes through
        /// <c>Text(value)</c> are observed through <see cref="NoteTextWrites"/>,
        /// mirroring the real one-note-per-cell API where a write replaces the
        /// whole note text.
        /// </summary>
        private void WithNote(string text)
        {
            ExistingComment = new Mock<Excel.Comment>();
            _ = ExistingComment
                .Setup(c => c.Text(It.IsAny<object>(), It.IsAny<object>(), It.IsAny<object>()))
                .Returns(text);
            _ = Cell.SetupGet(c => c.Comment).Returns(ExistingComment.Object);
        }

        /// <summary>Every property setter invoked on this cell (must stay empty).</summary>
        public IReadOnlyList<string> SetterInvocations =>
            [.. Cell.Invocations.Where(i => i.Method.Name.StartsWith("set_", StringComparison.Ordinal))
                .Select(i => i.Method.Name)];
    }

    /// <summary>The live-shaped COM graph a successful report walks.</summary>
    private sealed class ReporterGraph
    {
        private readonly List<Mock<Excel.ListColumn>> _columns = [];

        public ReporterGraph(IReadOnlyList<string> headers)
        {
            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            _ = Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            _ = Sheets.SetupGet(s => s.Count).Returns(1);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = ListObjects.SetupGet(l => l.Count).Returns(1);
            _ = Table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = Table.SetupGet(t => t.ListColumns).Returns(Columns.Object);
            _ = Table.SetupGet(t => t.DataBodyRange).Returns(Body.Object);
            _ = Columns.SetupGet(c => c.Count).Returns(headers.Count);
            foreach (string header in headers)
            {
                var column = new Mock<Excel.ListColumn>();
                _ = column.SetupGet(c => c.Name).Returns(header);
                _columns.Add(column);
            }
        }

        public Mock<Excel.Application> Application { get; } = new();

        public Mock<Excel.Workbook> Workbook { get; } = new();

        public Mock<Excel.Sheets> Sheets { get; } = new();

        public Mock<Excel.Worksheet> Worksheet { get; } = new();

        public Mock<Excel.ListObjects> ListObjects { get; } = new();

        public Mock<Excel.ListObject> Table { get; } = new();

        public Mock<Excel.ListColumns> Columns { get; } = new();

        public Mock<Excel.Range> Body { get; } = new();

        public Dictionary<(int Row, int Column), CellGraph> Cells { get; } = [];

        /// <summary>
        /// Cells that <c>SpecialCells(xlCellTypeComments)</c> reports, in the
        /// order the removal pass should visit them.
        /// </summary>
        public List<CellGraph> CommentedCells { get; } = [];

        /// <summary>
        /// Makes the body report the supplied cells as the ones carrying notes,
        /// so the stale-section removal path runs against fakes.
        /// </summary>
        public void WithCommentedCells(params CellGraph[] cells)
        {
            CommentedCells.Clear();
            CommentedCells.AddRange(cells);
            _ = Body
                .Setup(r => r.SpecialCells(It.IsAny<Excel.XlCellType>(), It.IsAny<object>()))
                .Returns(Body.Object); // the same mock stands in for the commented range
        }

        /// <summary>Builds the reporter under test wired to this graph.</summary>
        public TestableReporter Build() => new TestableReporter(
            Application.Object,
            sheetAt: (_, _) => Worksheet.Object,
            tableAt: (_, _) => Table.Object,
            columnAt: (_, index) => _columns[index - 1].Object,
            cellAt: (_, row, column) => CellAt(row, column).Cell.Object,
            enumerate: _ => CommentedCells.Select(cell => cell.Cell.Object));

        public CellGraph Cell(int row, int column) => CellAt(row, column);

        public Excel.ListColumn BuildColumn(int index) => _columns[index - 1].Object;

        public bool WasCellRequested(int row, int column) => Cells.ContainsKey((row, column));

        /// <summary>Every cell the reporter asked for, as (row, column) pairs.</summary>
        public IReadOnlyList<(int Row, int Column)> RequestedCells => [.. Cells.Keys];

        private CellGraph CellAt(int row, int column)
        {
            if (!Cells.TryGetValue((row, column), out CellGraph? cell))
            {
                cell = new CellGraph(row, column);
                Cells[(row, column)] = cell;
            }

            return cell;
        }
    }

    /// <summary>
    /// The reporter with its three COM indexer seams replaced by delegates, so a
    /// fake worksheet is exercised without a live Excel host.
    /// </summary>
    private sealed class TestableReporter : ExcelGanttValidationReporter
    {
        public TestableReporter(
            object? application,
            IWorksheetProtectionGuard? protectionGuard = null,
            Func<Excel.Sheets, int, object>? sheetAt = null,
            Func<Excel.ListObjects, int, Excel.ListObject>? tableAt = null,
            Func<Excel.ListColumns, int, Excel.ListColumn>? columnAt = null,
            Func<Excel.Range?, int, int, Excel.Range>? cellAt = null,
            Func<Excel.Range, IEnumerable<Excel.Range>>? enumerate = null)
            : base(application, protectionGuard ?? new AllowProtectionGuard())
        {
            SheetAt = sheetAt ?? ((_, _) => throw new InvalidOperationException("sheet seam not expected"));
            TableAt = tableAt ?? ((_, _) => throw new InvalidOperationException("table seam not expected"));
            ColumnAt = columnAt ?? ((_, _) => throw new InvalidOperationException("column seam not expected"));
            CellAt = cellAt ?? ((_, _, _) => throw new InvalidOperationException("cell seam not expected"));
            Enumerate = enumerate ?? (_ => throw new InvalidOperationException("enumerate seam not expected"));
        }

        public Func<Excel.Sheets, int, object> SheetAt { get; }

        public Func<Excel.ListObjects, int, Excel.ListObject> TableAt { get; }

        public Func<Excel.ListColumns, int, Excel.ListColumn> ColumnAt { get; }

        public Func<Excel.Range?, int, int, Excel.Range> CellAt { get; }

        public Func<Excel.Range, IEnumerable<Excel.Range>> Enumerate { get; }

        internal override object GetSheetAt(Excel.Sheets sheets, int index) => SheetAt(sheets, index);

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index)
            => TableAt(listObjects, index);

        internal override Excel.ListColumn GetTableAt(Excel.ListColumns columns, int index)
            => ColumnAt(columns, index);

        internal override Excel.Range GetCellRange(Excel.Range? body, int rowNumber, int columnIndex)
            => CellAt(body, rowNumber, columnIndex);

        internal override IEnumerable<Excel.Range> EnumerateCells(Excel.Range range) => Enumerate(range);
    }

    private sealed class AllowProtectionGuard : IWorksheetProtectionGuard
    {
        public ProtectionGuardOutcome Query() => ProtectionGuardOutcome.NotProtected;

        public ProtectionGuardOutcome QueryTarget(object? target) => ProtectionGuardOutcome.NotProtected;
    }

    private static GanttValidationIssue Issue(
        int row,
        string field,
        GanttValidationSeverity severity,
        string message) => new(row, field, "GC9001", severity, message);

    private static ReporterGraph Graph() => new(SchemaHeaders());

    /// <summary>The one-based schema column index of a header (for anchoring assertions).</summary>
    private static int ColumnOf(string header)
    {
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        for (int index = 0; index < columns.Count; index++)
        {
            if (string.Equals(columns[index].Name, header, StringComparison.Ordinal))
            {
                return index + 1;
            }
        }

        throw new InvalidOperationException($"No schema column named '{header}'.");
    }

    [Fact]
    public void Report_writes_one_note_on_the_offending_cell_and_touches_no_other_cell()
    {
        // Gate A (fake-worksheet mutation): two findings on two different rows
        // and columns must produce exactly two note writes, on exactly those two
        // cells, and the reporter must never even ask for another cell.
        ReporterGraph graph = Graph();
        CellGraph startCell = graph.Cell(1, ColumnOf("Start"));
        CellGraph typeCell = graph.Cell(2, ColumnOf("Type"));

        GanttValidationReportOutcome outcome = graph.Build().Report(
        [
            Issue(1, "Start", GanttValidationSeverity.Error, "Start is not a date."),
            Issue(2, "Type", GanttValidationSeverity.Error, "Type 'Phase' is unknown."),
        ]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.NotesWritten);
        Assert.Null(outcome.Refusal);

        Assert.StartsWith(
            ValidationReportComposer.NoteSentinelPrefix,
            Assert.Single(startCell.NotesWritten),
            StringComparison.Ordinal);
        Assert.Contains(
            "Start is not a date.", Assert.Single(startCell.NotesWritten), StringComparison.Ordinal);
        Assert.Contains(
            "Type 'Phase' is unknown.", Assert.Single(typeCell.NotesWritten), StringComparison.Ordinal);
        Assert.Contains("column 'Type'", Assert.Single(typeCell.NotesWritten), StringComparison.Ordinal);

        Assert.Equal(2, graph.RequestedCells.Count);
        Assert.True(graph.WasCellRequested(1, ColumnOf("Start")));
        Assert.True(graph.WasCellRequested(2, ColumnOf("Type")));
    }

    [Fact]
    public void Report_collapses_two_findings_on_one_cell_into_one_note()
    {
        ReporterGraph graph = Graph();
        CellGraph cell = graph.Cell(3, ColumnOf("Finish"));

        GanttValidationReportOutcome outcome = graph.Build().Report(
        [
            Issue(3, "Finish", GanttValidationSeverity.Error, "Finish is before Start."),
            Issue(3, "Finish", GanttValidationSeverity.Warning, "Finish is very far in the future."),
        ]);

        Assert.Equal(1, outcome.NotesWritten);
        string note = Assert.Single(cell.NotesWritten);
        Assert.Contains("Finish is before Start.", note, StringComparison.Ordinal);
        Assert.Contains("Finish is very far in the future.", note, StringComparison.Ordinal);
    }

    [Fact]
    public void Report_keeps_the_supplied_error_order_within_a_note()
    {
        // Gate B (error order): the note lines must preserve the
        // Severity -> Row -> Field -> Code order the validator produced.
        ReporterGraph graph = Graph();
        CellGraph cell = graph.Cell(1, ColumnOf("Visible"));

        _ = graph.Build().Report(
        [
            Issue(1, "Visible", GanttValidationSeverity.Error, "E-first."),
            Issue(1, "Visible", GanttValidationSeverity.Error, "E-second."),
            Issue(1, "Visible", GanttValidationSeverity.Warning, "W-third."),
        ]);

        string note = Assert.Single(cell.NotesWritten);
        var first = note.IndexOf("E-first.", StringComparison.Ordinal);
        var second = note.IndexOf("E-second.", StringComparison.Ordinal);
        var third = note.IndexOf("W-third.", StringComparison.Ordinal);
        Assert.True(first >= 0 && second > first && third > second, $"Order violated: {note}");
    }

    [Fact]
    public void Report_write_order_follows_the_supplied_issue_order()
    {
        ReporterGraph graph = Graph();
        CellGraph laterRow = graph.Cell(5, ColumnOf("Id"));
        CellGraph earlierRow = graph.Cell(1, ColumnOf("Id"));

        _ = graph.Build().Report(
        [
            Issue(1, "Id", GanttValidationSeverity.Error, "Row one."),
            Issue(5, "Id", GanttValidationSeverity.Error, "Row five."),
        ]);

        Assert.Contains("Row one.", Assert.Single(earlierRow.NotesWritten), StringComparison.Ordinal);
        Assert.Contains("Row five.", Assert.Single(laterRow.NotesWritten), StringComparison.Ordinal);
    }

    [Fact]
    public void Report_writes_no_note_but_still_scans_for_stale_sections_when_every_row_is_valid()
    {
        // A row that has become valid must stop showing its old note, so the
        // removal scan runs even when there is nothing new to write.
        ReporterGraph graph = Graph();

        GanttValidationReportOutcome outcome = graph.Build().Report([]);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.NotesWritten);
        Assert.Empty(graph.RequestedCells);
        graph.Body.Verify(
            r => r.SpecialCells(It.IsAny<Excel.XlCellType>(), It.IsAny<object>()), Times.Once);
    }

    [Fact]
    public void Report_preserves_a_user_note_by_replacing_only_the_add_in_section()
    {
        // Excel allows one classic note per cell and Range.AddComment fails when
        // a note already exists, so the report is written through Comment.Text
        // and the user's text survives verbatim as a prefix. The user's note is
        // never deleted and no second note is created.
        ReporterGraph graph = Graph();
        CellGraph cell = graph.Cell(1, ColumnOf("Start"));
        cell.WithUserNote("Reminder I typed myself.");

        GanttValidationReportOutcome outcome = graph.Build().Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is not a date.")]);

        Assert.Equal(1, outcome.NotesWritten);
        Assert.Empty(cell.NotesWritten);
        Assert.Equal(0, cell.NoteDeleteCount);
        var written = Assert.Single(cell.NoteTextWrites);
        Assert.StartsWith("Reminder I typed myself.", written, StringComparison.Ordinal);
        Assert.Contains("Start is not a date.", written, StringComparison.Ordinal);
    }

    [Fact]
    public void Report_replaces_a_stale_owned_note_on_the_offending_cell()
    {
        ReporterGraph graph = Graph();
        CellGraph cell = graph.Cell(1, ColumnOf("Start"));
        cell.WithOwnedNote("row 1, column 'Start': Error: stale finding");

        GanttValidationReportOutcome outcome = graph.Build().Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is not a date.")]);

        Assert.Equal(1, outcome.NotesWritten);
        Assert.Empty(cell.NotesWritten);
        var written = Assert.Single(cell.NoteTextWrites);
        Assert.Contains("Start is not a date.", written, StringComparison.Ordinal);
        Assert.DoesNotContain("stale finding", written, StringComparison.Ordinal);
    }

    [Fact]
    public void Report_removes_a_purely_owned_stale_note_and_keeps_a_user_prefix()
    {
        // The removal pass runs before writing: a cell whose note is purely the
        // add-in's loses the note entirely, a cell with user text keeps that text,
        // and a note the add-in never wrote is ignored.
        ReporterGraph graph = Graph();
        CellGraph ownedOnly = graph.Cell(1, ColumnOf("Id"));
        ownedOnly.WithOwnedNote("row 1, column 'Id': Error: fixed since");
        CellGraph userPlusOwned = graph.Cell(2, ColumnOf("Id"));
        userPlusOwned.WithUserNote(
            "Keep me."
            + Environment.NewLine
            + ValidationReportComposer.NoteSentinelPrefix
            + " row 1, column 'Id': Error: old section");
        CellGraph userOnly = graph.Cell(3, ColumnOf("Id"));
        userOnly.WithUserNote("Mine.");
        graph.WithCommentedCells(ownedOnly, userPlusOwned, userOnly);

        GanttValidationReportOutcome outcome = graph.Build().Report([]);

        Assert.Equal(0, outcome.NotesWritten);
        Assert.Equal(1, ownedOnly.NoteDeleteCount);
        Assert.Equal(0, userPlusOwned.NoteDeleteCount);
        Assert.Equal("Keep me.", Assert.Single(userPlusOwned.NoteTextWrites));
        Assert.Empty(userOnly.NoteTextWrites);
        Assert.Equal(0, userOnly.NoteDeleteCount);
    }

    [Fact]
    public void Report_scans_for_stale_sections_once_per_run()
    {
        ReporterGraph graph = Graph();

        _ = graph.Build().Report([Issue(1, "Start", GanttValidationSeverity.Error, "Start is invalid.")]);
        _ = graph.Build().Report([]);

        graph.Body.Verify(
            r => r.SpecialCells(It.IsAny<Excel.XlCellType>(), It.IsAny<object>()), Times.Exactly(2));
    }

    [Fact]
    public void Report_never_sets_any_cell_or_range_property()
    {
        // The R2.6 invariant: validation annotates, it never mutates a value or a
        // format. Any property setter on any reached COM object fails this test.
        ReporterGraph graph = Graph();
        _ = graph.Cell(1, ColumnOf("Start"));
        _ = graph.Cell(2, ColumnOf("Type"));

        _ = graph.Build().Report(
        [
            Issue(1, "Start", GanttValidationSeverity.Error, "Start is not a date."),
            Issue(2, "Type", GanttValidationSeverity.Error, "Type is unknown."),
        ]);

        Assert.Empty(SetterInvocations(graph.Body));
        Assert.Empty(SetterInvocations(graph.Table));
        Assert.Empty(SetterInvocations(graph.Worksheet));
        foreach (CellGraph cell in graph.Cells.Values)
        {
            Assert.Empty(cell.SetterInvocations);
        }

        graph.Body.VerifySet(r => r.Value2 = It.IsAny<object>(), Times.Never);
    }

    [Fact]
    public void Report_counts_distinct_cells_not_issues()
    {
        ReporterGraph graph = Graph();
        _ = graph.Cell(1, ColumnOf("Start"));
        _ = graph.Cell(2, ColumnOf("Finish"));

        GanttValidationReportOutcome outcome = graph.Build().Report(
        [
            Issue(1, "Start", GanttValidationSeverity.Error, "One."),
            Issue(1, "Start", GanttValidationSeverity.Warning, "Two."),
            Issue(2, "Finish", GanttValidationSeverity.Error, "Three."),
        ]);

        Assert.Equal(2, outcome.NotesWritten);
    }

    [Fact]
    public void Report_refuses_target_protected_before_touching_any_note()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        ReporterGraph graph = Graph();
        CellGraph cell = graph.Cell(1, ColumnOf("Start"));
        cell.WithOwnedNote("row 1, column 'Start': Error: stale");
        TestableReporter reporter = new TestableReporter(
            graph.Application.Object,
            guard.Object,
            sheetAt: (_, _) => graph.Worksheet.Object,
            tableAt: (_, _) => graph.Table.Object,
            columnAt: (_, index) => graph.BuildColumn(index),
            cellAt: (_, row, column) => graph.Cell(row, column).Cell.Object,
            enumerate: _ => graph.CommentedCells.Select(c => c.Cell.Object));

        GanttValidationReportOutcome outcome = reporter.Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is invalid.")]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Equal(0, outcome.NotesWritten);
        guard.Verify(g => g.Query(), Times.Once);
        Assert.Empty(cell.NotesWritten);
        Assert.Empty(cell.NoteTextWrites);
        Assert.Equal(0, cell.NoteDeleteCount);
        graph.Body.Verify(
            r => r.SpecialCells(It.IsAny<Excel.XlCellType>(), It.IsAny<object>()),
            Times.Never);
    }

    [Fact]
    public void Report_refuses_workbook_structure_protected_before_touching_any_note()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.WorkbookStructureProtected);
        ReporterGraph graph = Graph();
        TestableReporter reporter = new TestableReporter(
            graph.Application.Object,
            guard.Object,
            sheetAt: (_, _) => graph.Worksheet.Object,
            tableAt: (_, _) => graph.Table.Object,
            columnAt: (_, index) => graph.BuildColumn(index),
            cellAt: (_, row, column) => graph.Cell(row, column).Cell.Object,
            enumerate: _ => graph.CommentedCells.Select(c => c.Cell.Object));

        GanttValidationReportOutcome outcome = reporter.Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is invalid.")]);

        Assert.Equal(GanttValidationReportRefusalReason.TargetProtected, outcome.Refusal);
        guard.Verify(g => g.Query(), Times.Once);
        graph.Body.Verify(
            r => r.SpecialCells(It.IsAny<Excel.XlCellType>(), It.IsAny<object>()),
            Times.Never);
    }

    [Fact]
    public void Report_refuses_with_no_active_workbook_when_the_host_supplied_no_application()
    {
        GanttValidationReportOutcome outcome = new TestableReporter(null).Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is invalid.")]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Equal(0, outcome.NotesWritten);
    }

    [Fact]
    public void Report_refuses_with_no_active_workbook_when_no_workbook_is_open()
    {
        var application = new Mock<Excel.Application>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        GanttValidationReportOutcome outcome = new TestableReporter(application.Object).Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is invalid.")]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    [Fact]
    public void Report_refuses_when_the_table_is_absent_and_touches_no_cell()
    {
        ReporterGraph graph = Graph();
        _ = graph.Sheets.SetupGet(s => s.Count).Returns(0);

        GanttValidationReportOutcome outcome = graph.Build().Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is invalid.")]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(graph.RequestedCells);
        graph.Body.Verify(
            r => r.SpecialCells(It.IsAny<Excel.XlCellType>(), It.IsAny<object>()), Times.Never);
    }

    [Fact]
    public void Report_refuses_when_a_required_column_is_missing_and_touches_no_cell()
    {
        List<string> headers = SchemaHeaders();
        _ = headers.Remove("Id");
        var graph = new ReporterGraph(headers);

        GanttValidationReportOutcome outcome = graph.Build().Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is invalid.")]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(graph.RequestedCells);
    }

    [Fact]
    public void Report_refuses_when_the_table_is_named_differently()
    {
        ReporterGraph graph = Graph();
        _ = graph.Table.SetupGet(t => t.Name).Returns("tblSomethingElse");

        GanttValidationReportOutcome outcome = graph.Build().Report(
            [Issue(1, "Start", GanttValidationSeverity.Error, "Start is invalid.")]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttValidationReportRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(graph.RequestedCells);
    }

    private static IReadOnlyList<string> SetterInvocations(Mock mock) =>
        [.. mock.Invocations
            .Where(i => i.Method.Name.StartsWith("set_", StringComparison.Ordinal))
            .Select(i => i.Method.Name)];
}
