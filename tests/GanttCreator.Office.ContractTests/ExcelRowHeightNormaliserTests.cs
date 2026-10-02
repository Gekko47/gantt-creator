using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelRowHeightNormaliser"/> (R4.7D,
/// ADR-0026 D3/D4): a dragged managed row is restored, a normalised sheet writes
/// nothing, and every refusal path fires with no write.
/// </summary>
public sealed class ExcelRowHeightNormaliserTests
{
    private const double ManagedPt = 18;
    private const double SplitterPt = 24;
    private const double SpacerPt = 6;

    /// <summary>
    /// Overrides the COM seams so the test can supply row heights, the per-row
    /// <c>Type</c> text, and record writes without a live Excel host.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="table">The resolved table.</param>
    /// <param name="heights">One measured height per body row.</param>
    /// <param name="written">The rows written, recorded by the test double.</param>
    /// <param name="types">
    /// One Type display name per body row. When absent, no Type column resolves and
    /// every row is treated as managed -- which is what the pre-existing tests rely on.
    /// </param>
    private sealed class TestableNormaliser(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.Worksheet worksheet,
        Excel.ListObject table,
        IReadOnlyList<double> heights,
        List<int> written,
        IReadOnlyList<string>? types = null)
        : ExcelRowHeightNormaliser(application, guard)
    {
        internal override bool TryFindTable(
            Excel.Sheets sheets,
            out Excel.Worksheet? resolvedWorksheet,
            out Excel.ListObject? resolvedTable)
        {
            resolvedWorksheet = worksheet;
            resolvedTable = table;
            return true;
        }

        internal override Excel.Range? GetTableBody(Excel.ListObject source) => SourceBody;

        internal override int GetBodyRowCount(Excel.Range body) => heights.Count;

        internal override Excel.Range? GetBodyRowAt(Excel.Range body, int index)
        {
            var row = new Mock<Excel.Range>();
            _ = row.SetupGet(r => r.RowHeight).Returns(heights[index - 1]);
            _ = row.SetupSet(r => r.RowHeight = It.IsAny<object>())
                .Callback<object>(_ => written.Add(index));
            return row.Object;
        }

        internal override bool TryGetTypeColumn(Excel.ListObject source, out Excel.Range? typeColumn)
        {
            if (types is null)
            {
                typeColumn = null;
                return false;
            }

            typeColumn = new Mock<Excel.Range>().Object;
            return true;
        }

        internal override string? ReadTypeCellText(Excel.Range typeColumn, int index) =>
            types is null || index < 1 || index > types.Count ? null : types[index - 1];

        private Excel.Range SourceBody { get; } = new Mock<Excel.Range>().Object;
    }

    private static (Mock<Excel.Application> Application, Mock<Excel.Worksheet> Worksheet, Mock<Excel.ListObject> Table) Graph()
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

    /// <summary>
    /// A dragged managed row is restored to the token. This is the whole purpose of
    /// the adapter: the chart reads the measured row, so a stray drag would
    /// desynchronise one lane from the row a user is reading.
    /// </summary>
    [Fact]
    public void A_dragged_managed_row_is_restored()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [45, ManagedPt, 30],
            written);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(2, outcome.RowsWritten);
        Assert.Equal([1, 3], written);
    }

    /// <summary>
    /// A correctly normalised sheet writes nothing, so a Refresh does not mark the
    /// workbook dirty on every run.
    /// </summary>
    [Fact]
    public void An_already_normalised_sheet_writes_nothing()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [ManagedPt, ManagedPt, ManagedPt],
            written);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.RowsWritten);
        Assert.Empty(written);
    }

    /// <summary>
    /// A protected target is refused before any write. This is the ADR-0008 D4
    /// guard-first rule, and the write list is asserted empty so the test proves the
    /// refusal happened before the COM write rather than after it.
    /// </summary>
    [Fact]
    public void A_protected_target_is_refused_without_writing()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.SheetProtected),
            worksheet.Object,
            table.Object,
            [45, 45],
            written);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(written);
    }

    /// <summary>
    /// A guard reporting no active workbook is distinguished from a protected sheet,
    /// so the two failures are not collapsed into one message.
    /// </summary>
    [Fact]
    public void A_missing_workbook_from_the_guard_is_refused_as_such()
    {
        var (application, worksheet, table) = Graph();
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NoActiveWorkbook),
            worksheet.Object,
            table.Object,
            [45],
            []);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>
    /// A null application object degrades to the no-active-workbook refusal with no
    /// mutation, rather than throwing on the way to a COM call.
    /// </summary>
    [Fact]
    public void A_null_application_is_refused()
    {
        var normaliser = new TestableNormaliser(
            null,
            Guard(ProtectionGuardOutcome.NotProtected),
            new Mock<Excel.Worksheet>().Object,
            new Mock<Excel.ListObject>().Object,
            [45],
            []);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>
    /// A non-positive target is refused rather than written, because a nonsense row
    /// height is exactly the misalignment this adapter exists to remove.
    /// </summary>
    [Fact]
    public void A_non_positive_target_is_refused_without_writing()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [45],
            written);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(0, SplitterPt, SpacerPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(RowHeightNormalisationRefusalReason.InvalidMeasurement, outcome.Refusal);
        Assert.Empty(written);
    }

    /// <summary>
    /// A <c>Splitter</c> row is restored to <c>SplitterPt</c>, never to the managed
    /// height.
    /// </summary>
    /// <remarks>
    /// The measurement pass used to build every <c>MeasuredRowHeight</c> with the
    /// default <see cref="MeasuredRowKind.Managed"/>, ignoring the row's own
    /// <c>Type</c>. So a structural row was rewritten to the managed height by the very
    /// Refresh meant to restore the sheet: the section header and the blank separator
    /// could never be the right height, and the lane for such a row could not line up
    /// with the worksheet row it sits in. This is the positive test for reading the
    /// Type column -- deleting that read fails it.
    /// </remarks>
    [Fact]
    public void A_splitter_row_is_restored_to_the_splitter_token_not_the_managed_height()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            // Row 1 is dragged; row 2 is a splitter already at its own token.
            [45, SplitterPt],
            written,
            types: ["As-Planned Activity", "Splitter"]);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.True(outcome.Succeeded);
        // Only the dragged managed row is written. If the splitter had been treated as
        // managed it would have been rewritten to 18pt -- a second write that this
        // assertion would catch.
        Assert.Equal(1, outcome.RowsWritten);
        Assert.Equal([1], written);
    }

    /// <summary>
    /// The mirror case for <c>Spacer</c>, and the pair that shows the kind is read from
    /// the Type rather than assumed: both rows sit at their own tokens and neither is
    /// written.
    /// </summary>
    [Fact]
    public void A_spacer_row_is_restored_to_the_spacer_token_not_the_managed_height()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [SpacerPt],
            written,
            types: ["Spacer"]);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.RowsWritten);
        Assert.Empty(written);
    }

    /// <summary>
    /// A dragged structural row is restored to its OWN token rather than left alone.
    /// The no-op case above proves the splitter is not over-written; this proves it is
    /// still recognised when it does need a write, which is the case the missing Type
    /// read broke in production.
    /// </summary>
    [Fact]
    public void A_dragged_splitter_row_is_restored_to_the_splitter_token()
    {
        var (application, worksheet, table) = Graph();
        List<int> written = [];
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [45],
            written,
            types: ["Splitter"]);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1, outcome.RowsWritten);
        Assert.Equal([1], written);
    }

    /// <summary>
    /// An empty body succeeds with zero writes. A table with no rows is a legitimate
    /// state, not a failure.
    /// </summary>
    [Fact]
    public void An_empty_body_succeeds_with_no_writes()
    {
        var (application, worksheet, table) = Graph();
        var normaliser = new TestableNormaliser(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected),
            worksheet.Object,
            table.Object,
            [],
            []);

        RowHeightNormalisationOutcome outcome = normaliser.Normalise(ManagedPt, SplitterPt, SpacerPt);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.RowsWritten);
    }
}
