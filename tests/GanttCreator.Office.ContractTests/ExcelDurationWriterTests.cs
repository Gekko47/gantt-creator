using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelDurationWriter"/> (R4.7F D5): the
/// <c>Duration</c> column is written in ONE bulk pass, an empty plan writes
/// nothing at all, and every refusal path fires before any COM write.
/// </summary>
public sealed class ExcelDurationWriterTests
{
    /// <summary>
    /// Overrides the COM seams so the test can supply a column payload and count
    /// the bulk assignments without a live Excel host. The write counter is what
    /// proves D5: a per-cell implementation would raise it once per cell.
    /// </summary>
    private sealed class TestableWriter(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        object?[] currentColumn,
        List<object?> assignments)
        : ExcelDurationWriter(application, guard)
    {
        internal int ReadCount { get; private set; }

        internal int WriteCount => assignments.Count;

        internal override bool TryFindTable(
            Excel.Sheets sheets,
            out Excel.Worksheet? resolvedWorksheet,
            out Excel.ListObject? resolvedTable)
        {
            resolvedWorksheet = worksheet;
            resolvedTable = table;
            return true;
        }

        internal override bool TryGetDurationColumnRange(Excel.ListObject source, out Excel.Range? column)
        {
            column = new Mock<Excel.Range>().Object;
            return true;
        }

        internal override object? GetRangeValues(Excel.Range range)
        {
            ReadCount++;

            // One column, one value per body row, shaped as Excel returns it.
            return ToColumnPayload(currentColumn);
        }

        /// <summary>
        /// Builds the SAFEARRAY-shaped payload Excel returns for a one-column range.
        /// A jagged array cannot express it, so CA1814 is suppressed here for the
        /// same reason it is suppressed in the adapter.
        /// </summary>
        /// <param name="values">One value per body row.</param>
        /// <returns>The <c>object[,]</c> payload.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage(
            "Performance",
            "CA1814:MultiDimArrays",
            Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage(
            "Performance",
            "CA1859:UseConcreteTypesWhenPossible",
            Justification = "The payload is returned through an object-typed seam, so the declared return type must stay object.")]
        private static object ToColumnPayload(object?[] values)
        {
            var payload = new object[values.Length, 1];
            for (var row = 0; row < values.Length; row++)
            {
                payload[row, 0] = values[row]!;
            }

            return payload;
        }

        internal override void SetRangeValues(Excel.Range range, object payload)
        {
            assignments.Add(payload);
        }
    }

    private static (
        Mock<Excel.Application> Application,
        Mock<Excel.Worksheet> Worksheet,
        Mock<Excel.ListObject> Table) Graph()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var worksheet = new Mock<Excel.Worksheet>();
        var table = new Mock<Excel.ListObject>();

        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        return (application, worksheet, table);
    }

    private static IWorksheetProtectionGuard Guard(ProtectionGuardOutcome outcome)
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object?>())).Returns(outcome);
        return guard.Object;
    }

    private static TestableWriter Writer(
        ProtectionGuardOutcome protection,
        object?[] currentColumn,
        List<object?> assignments)
    {
        (Mock<Excel.Application> application, Mock<Excel.Worksheet> worksheet, Mock<Excel.ListObject> table) = Graph();
        return new TestableWriter(
            application.Object,
            Guard(protection),
            worksheet.Object,
            table.Object,
            currentColumn,
            assignments);
    }

    private static DurationWritePlan Plan(params DurationWrite[] writes) => new(writes, []);

    /// <summary>
    /// A writer whose column payload is supplied verbatim, so a test can present the
    /// bare scalar Excel returns for a one-cell range rather than the array shape
    /// every other test uses.
    /// </summary>
    private sealed class ScalarPayloadWriter(
        object? application,
        IWorksheetProtectionGuard guard,
        object? payload,
        List<object?> assignments)
        : ExcelDurationWriter(application, guard)
    {
        internal int WriteCount => assignments.Count;

        internal override bool TryFindTable(
            Excel.Sheets sheets,
            out Excel.Worksheet? resolvedWorksheet,
            out Excel.ListObject? resolvedTable)
        {
            resolvedWorksheet = new Mock<Excel.Worksheet>().Object;
            resolvedTable = new Mock<Excel.ListObject>().Object;
            return true;
        }

        internal override bool TryGetDurationColumnRange(Excel.ListObject source, out Excel.Range? column)
        {
            column = new Mock<Excel.Range>().Object;
            return true;
        }

        internal override object? GetRangeValues(Excel.Range range) => payload;

        internal override void SetRangeValues(Excel.Range range, object written) => assignments.Add(written);
    }

    private static object?[] PayloadOf(List<object?> assignments)
        {
            var payload = Assert.IsType<object[,]>(Assert.Single(assignments));
            return [.. payload.Cast<object?>()];
        }

    [Fact]
    public void A_plan_is_written_in_one_bulk_assignment_not_one_per_cell()
    {
        // D5's whole point. Three changed rows must produce ONE ranged write; a
        // per-cell loop would produce three and dominate Refresh on a large table.
        List<object?> assignments = [];
        TestableWriter writer = Writer(
            ProtectionGuardOutcome.NotProtected,
            [null, null, null],
            assignments);

        DurationWriteOutcome outcome = writer.Write(Plan(
            new DurationWrite(2, "5"),
            new DurationWrite(3, "12"),
            new DurationWrite(4, "-")));

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, writer.WriteCount);
        Assert.Equal(3, outcome.CellsWritten);
    }

    [Fact]
    public void An_empty_plan_performs_no_read_and_no_write()
    {
        // The workbook must not be marked dirty by a Refresh of a correct sheet.
        List<object?> assignments = [];
        TestableWriter writer = Writer(
            ProtectionGuardOutcome.NotProtected,
            ["5", "5", "5"],
            assignments);

        DurationWriteOutcome outcome = writer.Write(Plan());

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, writer.WriteCount);
        Assert.Equal(0, writer.ReadCount);
        Assert.Equal(0, outcome.CellsWritten);
    }

    [Fact]
    public void The_bulk_payload_patches_only_the_planned_rows()
    {
        List<object?> assignments = [];
        TestableWriter writer = Writer(
            ProtectionGuardOutcome.NotProtected,
            ["keep-me", "5", "also-keep"],
            assignments);

        _ = writer.Write(Plan(new DurationWrite(2, "9")));

        object?[] written = PayloadOf(assignments);
        Assert.Equal("keep-me", written[0]);
        Assert.Equal("9", written[1]);
        Assert.Equal("also-keep", written[2]);
    }

    [Fact]
    public void A_cleared_duration_is_written_as_a_blank_cell_not_an_empty_string()
    {
        // Excel stores a blank as null. Writing "" would leave a zero-length string
        // the reader would report as present-but-empty, so the stale value would
        // read as still set.
        List<object?> assignments = [];
        TestableWriter writer = Writer(
            ProtectionGuardOutcome.NotProtected,
            ["12"],
            assignments);

        _ = writer.Write(Plan(new DurationWrite(1, string.Empty)));

        Assert.Null(PayloadOf(assignments)[0]);
    }

    [Fact]
    public void A_protected_target_refuses_with_no_write()
    {
        List<object?> assignments = [];
        TestableWriter writer = Writer(ProtectionGuardOutcome.SheetProtected, [null], assignments);

        DurationWriteOutcome outcome = writer.Write(Plan(new DurationWrite(1, "5")));

        Assert.False(outcome.Succeeded);
        Assert.Equal(DurationWriteRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Equal(0, writer.WriteCount);
    }

    [Fact]
    public void A_missing_workbook_refuses_with_no_write()
    {
        List<object?> assignments = [];
        TestableWriter writer = Writer(ProtectionGuardOutcome.NoActiveWorkbook, [null], assignments);

        DurationWriteOutcome outcome = writer.Write(Plan(new DurationWrite(1, "5")));

        Assert.False(outcome.Succeeded);
        Assert.Equal(DurationWriteRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Equal(0, writer.WriteCount);
    }

    [Fact]
    public void A_single_cell_column_reported_as_a_scalar_is_written_not_refused()
    {
        // A one-row, one-column `Value2` range returns a bare scalar rather than a
        // SAFEARRAY, and `ReadRows` only handles the array shape -- so this payload
        // produced no rows and the adapter refused with `WriteFailed`. A table with
        // exactly one body row is an ordinary workbook, not a broken one, and it has a
        // `Duration` value to write. Positive test for the scalar branch.
        List<object?> assignments = [];
        (Mock<Excel.Application> application, Mock<Excel.Worksheet> worksheet, Mock<Excel.ListObject> table) = Graph();
        var writer = new ScalarPayloadWriter(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            "7",
            assignments);

        DurationWriteOutcome outcome = writer.Write(Plan(new DurationWrite(1, "9")));

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.Equal(1, writer.WriteCount);
        Assert.Equal(["9"], PayloadOf(assignments));
    }

    [Fact]
    public void Write_refuses_a_null_plan_rather_than_writing_nothing_silently()
    {
        List<object?> assignments = [];
        TestableWriter writer = Writer(ProtectionGuardOutcome.NotProtected, [null], assignments);

        Assert.Throws<ArgumentNullException>(() => writer.Write(null!));
        Assert.Equal(0, writer.WriteCount);
    }
}