using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the row-height normaliser's interaction with
/// the outline groups (R4.7D D5): a Refresh runs the outline pass
/// and then the height pass, and a height write un-hides an
/// outline-collapsed row, so the pass must restore what it un-hid.
/// Tagged <c>[Trait("Category","OfficeIntegration")]</c> so
/// <c>verify-quick.ps1</c> and <c>verify.ps1</c> exclude it; run
/// via <c>pwsh ./scripts/verify-office.ps1</c>.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this suite exists.</b> The reported defect: Refresh
/// expands every collapsed group. The outline pass preserves
/// collapse (pinned by <c>OutlineGroupIntegrationTests</c>), but
/// the height pass that follows it in the same Refresh writes
/// <c>RowHeight</c> to every body row off the height token -- and
/// a live host un-hides a row when its height is written, so every
/// collapsed child row came back visible. A body row the
/// initialiser adds sits at the host default, not the token, so
/// the children are off-token by construction and the defect fires
/// without any pre-drag.
/// </para>
/// </remarks>
public class RowHeightNormaliserIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static GanttRowId NewId() => GanttRowId.New();

    private static GanttEvent Event(int rowNumber, GanttRowId id, GanttRowId? parentId = null) =>
        new(
            rowNumber,
            id,
            NewId(),
            0,
            GanttEntityType.AsPlannedActivity,
            "Build frame",
            new DateOnly(2026, 1, 1),
            new DateOnly(2026, 12, 31),
            parentId,
            "AsPlannedActivity",
            null,
            null,
            null,
            true,
            null
        );

    /// <summary>
    /// Creates an initialised Gantt sheet with <paramref name="bodyRows"/> data rows.
    /// </summary>
    private static (Excel.Worksheet Sheet, Excel.ListObject Table) InitialisedGantt(
        OfficeFixture fixture,
        OfficeFixture.ComScope scope,
        int bodyRows
    )
    {
        fixture.CreateWorkbook();
        var initialiser = new ExcelWorkbookInitialiser(fixture.Excel);
        WorkbookInitialiseOutcome initialised = initialiser.Initialise();
        Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

        Excel.Sheets sheets = scope.Track(fixture.Excel.ActiveWorkbook!.Sheets);
        Excel.Worksheet sheet = scope.Track((Excel.Worksheet)sheets[GanttWorkbookContract.GanttSheetLabel]);
        Excel.ListObjects objects = scope.Track(sheet.ListObjects);
        Excel.ListObject table = scope.Track(objects[GanttTableSchema.TableName]);

        // The initialiser creates the table from a header row alone, so a body has to
        // be added before the writer has any rows to outline.
        Excel.ListRows rows = scope.Track(table.ListRows);
        for (var i = 0; i < bodyRows; i++)
        {
            scope.Track(rows.Add());
        }

        return (sheet, table);
    }

    /// <summary>
    /// The reported defect: Refresh expands every collapsed group.
    /// The group is collapsed first -- the user's sequence -- and
    /// the height pass of the same Refresh then runs against the
    /// collapsed children. Their rows are off the height token
    /// (the host default, not <c>GanttRowHeightPt</c>), so the
    /// normaliser has a real write to make; the children must end
    /// the pass both restored to the token and still hidden.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_collapsed_group_stays_collapsed_through_height_normalisation()
    {
        var fixture = new OfficeFixture();
        var scope = new OfficeFixture.ComScope();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            (Excel.Worksheet sheet, Excel.ListObject table) = InitialisedGantt(fixture, scope, bodyRows: 4);

            GanttRowId parent = NewId();
            List<GanttEvent> hierarchy = [Event(1, parent), Event(2, NewId(), parent), Event(3, NewId(), parent), Event(4, NewId())];

            var writer = new ExcelOutlineGroupWriter(fixture.Excel);
            OutlineGroupOutcome outline = writer.Apply(hierarchy);
            Assert.True(outline.Succeeded, $"Outline refresh refused: {outline.Refusal}");

            int firstBodyRow = scope.Track(table.DataBodyRange!).Row;
            double token = GanttCatalogues.MetricDefault("GanttRowHeightPt");

            // The added rows sit at the host default, not the token,
            // so the height pass has a write to make -- the
            // precondition the defect needs.
            Excel.Range childAtDefault = scope.Track(sheet.Rows[firstBodyRow + 1]);
            Assert.NotEqual(token, Convert.ToDouble(childAtDefault.RowHeight, CultureInfo.InvariantCulture));

            Excel.Outline excelOutline = scope.Track(sheet.Outline);
            excelOutline.ShowLevels(1, 1);
            try
            {
                Excel.Range child = scope.Track(sheet.Rows[firstBodyRow + 1]);
                Assert.True(child.Hidden, "Child row did not hide when the group was collapsed.");

                // The height pass of the same Refresh, run while the
                // group is collapsed.
                var normaliser = new ExcelRowHeightNormaliser(fixture.Excel);
                RowHeightNormalisationOutcome heights = normaliser.Normalise(
                    GanttCatalogues.MetricDefault("GanttRowHeightPt"),
                    GanttCatalogues.MetricDefault("SplitterHeightPt"),
                    GanttCatalogues.MetricDefault("SpacerHeightPt"),
                    GanttCatalogues.MetricDefault("PeriodBandHeightPt"),
                    GanttCatalogues.MetricDefault("YearBandHeightPt"),
                    GanttCatalogues.MetricDefault("ChartPaddingRowHeightPt"),
                    GanttCatalogues.MetricDefault("ChartAnchorRowHeightPt"));
                Assert.True(heights.Succeeded, $"Height normalisation refused: {heights.Refusal}");

                // Every body row sits at the host default, so all four
                // were in the write set.
                Assert.True(heights.RowsWritten >= 4, $"Only {heights.RowsWritten} rows written.");

                // The defect: the height write must not un-hide the
                // collapsed children. A hidden row reports its height
                // as 0 on the host, so the restored height is asserted
                // after the group is expanded again, below.
                Excel.Range childAfter = scope.Track(sheet.Rows[firstBodyRow + 1]);
                Assert.True(
                    childAfter.Hidden,
                    "Collapsed group re-expanded by the height normalisation.");

                _output.WriteLine(
                    $"Excel build {fixture.Excel.Version} (PID {fixture.ProcessId}); "
                        + $"collapsed children survived height normalisation ({heights.RowsWritten} rows written)");
            }
            finally
            {
                excelOutline.ShowLevels(2, 2);
            }

            // The height write landed: expanded, the children sit at
            // the token again.
            Excel.Range expandedChild = scope.Track(sheet.Rows[firstBodyRow + 1]);
            Assert.Equal(token, Convert.ToDouble(expandedChild.RowHeight, CultureInfo.InvariantCulture));
        }
        finally
        {
            scope.Dispose();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
