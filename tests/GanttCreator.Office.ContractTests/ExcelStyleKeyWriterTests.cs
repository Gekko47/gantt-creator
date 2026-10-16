using System.Runtime.InteropServices;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// <see cref="ExcelStyleKeyWriter"/> writes the engine-owned <c>StyleKey</c>
/// column (column K) in one bulk pass: read once, patch the planned rows, write
/// back once. Its COM touch points are internal virtual test seams, which this
/// test class subclasses instead of touching Excel.
/// </summary>
public sealed class ExcelStyleKeyWriterTests
{
    /// <summary>
    /// Test harness over <see cref="ExcelStyleKeyWriter"/>: every COM call is
    /// overridden, so no Excel host is required.
    /// </summary>
    private sealed class Harness : ExcelStyleKeyWriter
    {
        public Harness(
            object? application,
            IWorksheetProtectionGuard protectionGuard)
            : base(application, protectionGuard)
        {
        }

        public Excel.Range? Column { get; set; }
        public Array? Values { get; set; }
        public bool FindTable { get; set; } = true;
        public Worksheet? Worksheet { get; set; }
        public ListObject? Table { get; set; }
        public int SetRangeValuesCalls { get; private set; }
        public int GetRangeValuesCalls { get; private set; }
        public object? LastPayload { get; private set; }
        public bool SetRangeValuesThrows { get; set; }

        internal override bool TryGetStyleKeyColumnRange(
            ListObject table,
            out Excel.Range? column)
        {
            column = Column;
            return Column is not null;
        }

        internal override object? GetRangeValues(Excel.Range range)
        {
            GetRangeValuesCalls++;
            return Values;
        }

        internal override void SetRangeValues(Excel.Range range, object payload)
        {
            SetRangeValuesCalls++;
            LastPayload = payload;
            if (SetRangeValuesThrows)
            {
                Marshal.ThrowExceptionForHR(unchecked((int)0x80004005));
            }
        }

        internal override bool TryFindTable(
            Sheets sheets,
            out Worksheet? worksheet,
            out ListObject? table)
        {
            worksheet = Worksheet;
            table = Table;
            return FindTable;
        }
    }

    private static Harness WithWorkbook(
        bool findTable = true,
        Excel.Range? column = null,
        IWorksheetProtectionGuard? guard = null)
    {
        Mock<Application> application = new(MockBehavior.Loose);
        application.SetupGet(a => a.ActiveWorkbook).Returns(new Mock<Workbook>(MockBehavior.Loose).Object);

        return new Harness(application.Object, guard ?? NotProtectedGuard())
        {
            Column = column,
            FindTable = findTable,
            Worksheet = new Mock<Worksheet>(MockBehavior.Loose).Object,
            Table = new Mock<ListObject>(MockBehavior.Loose).Object,
        };
    }

    private static Excel.Range RangeMock() => new Mock<Excel.Range>(MockBehavior.Loose).Object;

    /// <summary>
    /// A one-column, N-row range body as Excel returns it: a rank-2 array, never
    /// jagged.
    /// </summary>
    private static Array CreateColumn(params object?[] values)
    {
        var column = Array.CreateInstance(typeof(object), values.Length, 1);
        for (var row = 0; row < values.Length; row++)
        {
            column.SetValue(values[row], row, 0);
        }

        return column;
    }

    private static IWorksheetProtectionGuard NotProtectedGuard()
    {
        Mock<IWorksheetProtectionGuard> guard = new(MockBehavior.Loose);
        guard.Setup(g => g.QueryTarget(It.IsAny<object?>())).Returns(ProtectionGuardOutcome.NotProtected);
        return guard.Object;
    }

    private static IWorksheetProtectionGuard ProtectedGuard()
    {
        Mock<IWorksheetProtectionGuard> guard = new(MockBehavior.Loose);
        guard.Setup(g => g.QueryTarget(It.IsAny<object?>())).Returns(ProtectionGuardOutcome.SheetProtected);
        return guard.Object;
    }

    /// <summary>
    /// An empty plan performs no write at all, so a correctly written sheet is not
    /// marked dirty on every Refresh.
    /// </summary>
    [Fact]
    public void An_empty_plan_writes_nothing_and_succeeds()
    {
        Harness writer = WithWorkbook();
        StyleKeyWritePlan plan = new([]);

        StyleKeyWriteOutcome outcome = writer.Write(plan);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.CellsWritten);
        Assert.Equal(0, outcome.WriteCount);
        Assert.Equal(0, writer.SetRangeValuesCalls);
        Assert.Equal(0, writer.GetRangeValuesCalls);
    }

    /// <summary>
    /// The planned rows are patched into the column's current values and written
    /// back in one ranged pass; every other row keeps its existing value.
    /// </summary>
    [Fact]
    public void The_planned_cells_are_written_in_one_bulk_pass()
    {
        Harness writer = WithWorkbook(column: RangeMock());
        writer.Values = CreateColumn("AsBuiltActivity", null);
        StyleKeyWritePlan plan = new([new StyleKeyWrite(2, "AsPlannedActivity")]);

        StyleKeyWriteOutcome outcome = writer.Write(plan);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.CellsWritten);
        Assert.Equal(1, outcome.WriteCount);
        Assert.Equal(1, writer.SetRangeValuesCalls);
        Assert.Equal(1, writer.GetRangeValuesCalls);
        Array payload = Assert.IsAssignableFrom<Array>(writer.LastPayload);
        Assert.Equal(2, payload.GetLength(0));
        Assert.Equal(1, payload.GetLength(1));
        Assert.Equal("AsBuiltActivity", payload.GetValue(0, 0));
        Assert.Equal("AsPlannedActivity", payload.GetValue(1, 0));
    }

    /// <summary>
    /// A write planned past the end of the column is refused before anything is
    /// written back.
    /// </summary>
    [Fact]
    public void A_write_past_the_end_of_the_column_is_refused()
    {
        Harness writer = WithWorkbook(column: RangeMock());
        writer.Values = CreateColumn("AsBuiltActivity", null);
        StyleKeyWritePlan plan = new([new StyleKeyWrite(5, "AsPlannedActivity")]);

        StyleKeyWriteOutcome outcome = writer.Write(plan);

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyWriteRefusalReason.WriteFailed, outcome.Refusal);
        Assert.Equal(0, outcome.CellsWritten);
        Assert.Equal(0, writer.SetRangeValuesCalls);
    }

    /// <summary>
    /// A COM failure writing the range is reported as a write refusal, not as an
    /// unhandled exception.
    /// </summary>
    [Fact]
    public void A_COM_write_failure_is_refused_as_write_failed()
    {
        Harness writer = WithWorkbook(column: RangeMock());
        writer.Values = CreateColumn("AsBuiltActivity");
        writer.SetRangeValuesThrows = true;
        StyleKeyWritePlan plan = new([new StyleKeyWrite(1, "AsPlannedActivity")]);

        StyleKeyWriteOutcome outcome = writer.Write(plan);

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyWriteRefusalReason.WriteFailed, outcome.Refusal);
        Assert.Equal(1, writer.SetRangeValuesCalls);
    }

    /// <summary>
    /// With no active workbook the writer refuses before touching the sheet.
    /// </summary>
    [Fact]
    public void No_active_workbook_is_refused()
    {
        Harness writer = new(null, NotProtectedGuard());
        StyleKeyWritePlan plan = new([new StyleKeyWrite(1, "AsPlannedActivity")]);

        StyleKeyWriteOutcome outcome = writer.Write(plan);

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyWriteRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Equal(0, outcome.CellsWritten);
    }

    /// <summary>
    /// A missing <c>tblGanttData</c> table is reported as TableMissing.
    /// </summary>
    [Fact]
    public void A_missing_table_is_refused()
    {
        Harness writer = WithWorkbook(findTable: false);
        StyleKeyWritePlan plan = new([new StyleKeyWrite(1, "AsPlannedActivity")]);

        StyleKeyWriteOutcome outcome = writer.Write(plan);

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyWriteRefusalReason.TableMissing, outcome.Refusal);
    }

    /// <summary>
    /// A protected target is refused before any cell is written.
    /// </summary>
    [Fact]
    public void A_protected_target_is_refused()
    {
        // The protection guard is consulted after the column exists, so a column
        // has to be present for the refusal to come from the guard.
        Harness writer = WithWorkbook(column: RangeMock(), guard: ProtectedGuard());
        StyleKeyWritePlan plan = new([new StyleKeyWrite(1, "AsPlannedActivity")]);

        StyleKeyWriteOutcome outcome = writer.Write(plan);

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyWriteRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Equal(0, writer.SetRangeValuesCalls);
    }

    /// <summary>
    /// A sheet without a StyleKey column cannot be written and reports TableMissing.
    /// </summary>
    [Fact]
    public void A_missing_style_key_column_is_refused()
    {
        Harness writer = WithWorkbook(column: null);
        writer.Values = CreateColumn("AsBuiltActivity");
        StyleKeyWritePlan plan = new([new StyleKeyWrite(1, "AsPlannedActivity")]);

        StyleKeyWriteOutcome outcome = writer.Write(plan);

        Assert.False(outcome.Succeeded);
        Assert.Equal(StyleKeyWriteRefusalReason.TableMissing, outcome.Refusal);
        Assert.Equal(0, writer.SetRangeValuesCalls);
    }
}

