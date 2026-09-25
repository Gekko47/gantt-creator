using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the identity repairer's ParentId rewrite pass. The audit
/// finding was that the blank-parent and duplicated-parent guards sat in a loop
/// that performed no writes, so the loop that does write could rewrite a blank
/// ParentId cell and could rewrite a ParentId whose target Id is duplicated —
/// the case ADR-0011 requires to be left unchanged for GanttRowValidator to
/// report as ParentAmbiguous.
/// </summary>
public sealed class IdentityRepairParentRewriteTests
{
    private static GanttRowId Id(char suffix) =>
        GanttRowId.Parse("G-0123456789abcdef0123456789abcde" + suffix);

    private sealed class TestableRepairer(
        object? application,
        IWorksheetProtectionGuard guard,
        Array bodyValues,
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        Action<int, int, object> recordWrite)
        : ExcelGanttRowIdentityRepairer(application, guard)
    {
        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) => worksheet;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) => table;

        internal override object? GetBodyValues(Excel.Range body) => bodyValues;

        internal override Excel.Range GetCell(Excel.Range body, int row, int column)
        {
            var cell = new Mock<Excel.Range>();
            _ = cell.SetupSet(c => c.Value2 = It.IsAny<object>())
                .Callback<object>(value => recordWrite(row, column, value));
            return cell.Object;
        }
    }

    private sealed class Graph
    {
        public Dictionary<(int Row, int Column), object?> Written { get; } = [];

        private readonly Mock<Excel.Range> _body = new();
        private readonly Mock<Excel.Application> _application = new();
        private readonly Mock<Excel.Worksheet> _worksheet = new();
        private readonly Mock<Excel.ListObject> _table = new();

        public Graph()
        {
            var workbook = new Mock<Excel.Workbook>();
            var sheets = new Mock<Excel.Sheets>();
            var listObjects = new Mock<Excel.ListObjects>();
            var listColumns = new Mock<Excel.ListColumns>();
            var idColumn = new Mock<Excel.ListColumn>();
            var parentColumn = new Mock<Excel.ListColumn>();

            _ = _application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
            _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
            _ = sheets.SetupGet(s => s.Count).Returns(1);
            _ = sheets.Setup(s => s[1]).Returns(_worksheet.Object);
            _ = _worksheet.SetupGet(w => w.ListObjects).Returns(listObjects.Object);
            _ = listObjects.SetupGet(l => l.Count).Returns(1);
            _ = listObjects.Setup(l => l[1]).Returns(_table.Object);
            _ = _table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = _table.SetupGet(t => t.DataBodyRange).Returns(_body.Object);
            _ = _table.SetupGet(t => t.ListColumns).Returns(listColumns.Object);
            _ = idColumn.SetupGet(c => c.Name).Returns("Id");
            _ = parentColumn.SetupGet(c => c.Name).Returns("ParentId");
            _ = listColumns.SetupGet(c => c.Count).Returns(2);
            _ = listColumns.Setup(c => c[1]).Returns(idColumn.Object);
            _ = listColumns.Setup(c => c[2]).Returns(parentColumn.Object);
        }

        public TestableRepairer Build(Array bodyValues, IWorksheetProtectionGuard guard) =>
            new(
                _application.Object,
                guard,
                bodyValues,
                _worksheet.Object,
                _table.Object,
                (row, column, value) => Written[(row, column)] = value);
    }

    private static IWorksheetProtectionGuard NotProtected()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
        return guard.Object;
    }

    private static Array BodyMatrix(params (string Id, string? ParentId)[] rows)
    {
        var matrix = Array.CreateInstance(typeof(object), [rows.Length, 2], [1, 1]);
        for (var index = 0; index < rows.Length; index++)
        {
            matrix.SetValue(rows[index].Id, index + 1, 1);
            matrix.SetValue(rows[index].ParentId, index + 1, 2);
        }

        return matrix;
    }

    [Fact]
    public void A_blank_parent_id_cell_is_never_rewritten_or_counted()
    {
        // Row 1's blank Id is repaired to a new Id; the blank ParentId in the same
        // row must stay blank rather than pick up that new Id.
        var graph = new Graph();
        TestableRepairer repairer = graph.Build(
            BodyMatrix(("", null), (Id('1').Value, null)),
            NotProtected());

        GanttRowIdentityRepairOutcome outcome = repairer.Repair();

        Assert.Null(outcome.Refusal);
        // Only the blank Id cell is repaired.
        Assert.Equal(1, outcome.RepairedCount);
        Assert.True(graph.Written.ContainsKey((1, 1)));
        Assert.DoesNotContain((1, 2), graph.Written.Keys);
        Assert.DoesNotContain((2, 2), graph.Written.Keys);
    }

    [Fact]
    public void A_parent_id_whose_target_is_duplicated_is_left_unchanged()
    {
        // ADR-0011: an ambiguous duplicate-parent reference is left alone and
        // reported by GanttRowValidator as ParentAmbiguous, so it must not be
        // rewritten. Row 1's malformed Id is canonical *and* replaced, so only
        // the duplicate-count guard stops row 3's ParentId from being pointed at
        // row 1's new Id.
        var graph = new Graph();
        TestableRepairer repairer = graph.Build(
            BodyMatrix(
                ("BAD", null),
                ("BAD", null),
                (Id('2').Value, "BAD")),
            NotProtected());

        GanttRowIdentityRepairOutcome outcome = repairer.Repair();

        Assert.Null(outcome.Refusal);
        // Only the non-canonical duplicate's Id cell is repaired; the canonical
        // first row is malformed too, so it is repaired as well, and the
        // ambiguous ParentId is left alone.
        Assert.Equal(2, outcome.RepairedCount);
        Assert.True(graph.Written.ContainsKey((1, 1)));
        Assert.True(graph.Written.ContainsKey((2, 1)));
        Assert.DoesNotContain((3, 2), graph.Written.Keys);
    }

    [Fact]
    public void An_unambiguous_parent_reference_is_rewritten_to_the_replacement()
    {
        // The positive control for the two guards above: a single, unique
        // malformed parent Id is remapped to its row's replacement and counted.
        var graph = new Graph();
        TestableRepairer repairer = graph.Build(
            BodyMatrix(
                ("BAD", null),
                (Id('2').Value, "BAD")),
            NotProtected());

        GanttRowIdentityRepairOutcome outcome = repairer.Repair();

        Assert.Null(outcome.Refusal);
        Assert.Equal(2, outcome.RepairedCount);
        Assert.True(GanttRowId.TryParse(
            Convert.ToString(graph.Written[(1, 1)], CultureInfo.InvariantCulture),
            out GanttRowId? original));
        Assert.True(GanttRowId.TryParse(
            Convert.ToString(graph.Written[(2, 2)], CultureInfo.InvariantCulture),
            out GanttRowId? rewritten));
        Assert.Equal(original!.Value, rewritten!.Value);
    }
}
