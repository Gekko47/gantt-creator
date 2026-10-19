using GanttCreator.Core;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>Live-Excel acceptance gate for the R5.3 parent/child commands.</summary>
/// <remarks>
/// <para>
/// This is the gate R5.3 recorded as outstanding: the work item requires "a live
/// exercise on the real worksheet" and it had not been run, which is how a
/// refused COM write reached users as
/// <see cref="ParentChildMutationRefusalReason.WriteFailed"/> on both verbs.
/// No mock can catch it — <c>ExcelParentChildMutator.GetCellAt</c> is an
/// <c>internal virtual</c> seam the contract harness overrides, so the
/// production line never ran there. Only real Excel resolves the accessor.
/// </para>
/// <para>
/// The source-shape counterpart that pins the accessor lives in
/// <c>GanttCreator.Architecture.Tests.ParentChildCellAccessorTests</c>, which
/// fails if this file is deleted.
/// </para>
/// </remarks>
public class ParentChildMutationIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static GanttRowId Id(char suffix) =>
        GanttRowId.Parse("G-0123456789abcdef0123456789abcde" + suffix);

    /// <summary>
    /// A promote clears the moved row's <c>ParentId</c> and resequences the run
    /// member above the moved window, with no refusal.
    /// </summary>
    /// <remarks>
    /// The resequenced cell above the window is written through the single-cell
    /// path — the exact call the refused accessor sat on — so a promote that
    /// moves a <em>middle</em> child is the shape that failed in production.
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Promote_clears_the_parent_and_resequences_the_run_above_the_window()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath), "RegisterXLL failed.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            Excel.Worksheet sheet =
                (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];

            // parent, then three children: promoting a MIDDLE child forces the
            // resequenced-cell write above the moved window.
            var inserter = new ExcelGanttRowInserter(fixture.Excel);
            foreach (char suffix in new[] { 'a', 'b', 'c', 'd' })
            {
                Assert.True(
                    inserter.Insert(GanttEntityType.AsPlannedActivity, () => Id(suffix)).Succeeded);
            }

            for (var child = 2; child <= 4; child++)
            {
                ParentOf(table, child, Id('a').Value);
                OrderOf(table, child, child - 2);
            }

            table.ListRows[3].Range.Select();

            ParentChildMutationOutcome outcome = new ExcelParentChildMutator(fixture.Excel).Promote();

            _output.WriteLine(
                $"PROMOTE succeeded={outcome.Succeeded} refusal={outcome.Refusal} " +
                $"planRefusal={outcome.PlanRefusal} cells={outcome.CellsWritten}");

            Assert.True(
                outcome.Succeeded,
                $"Promote refused: refusal={outcome.Refusal}, plan={outcome.PlanRefusal}. " +
                "A refused write is the defect the live gate exists to catch.");
            Assert.Null(outcome.Refusal);
            Assert.Null(outcome.PlanRefusal);

            // The promote moves the selected row's CONTENT to the bottom of its
            // sibling run (rows 2..4), so afterwards:
            //   row 2 = b (resequenced 0), row 3 = d (resequenced 1),
            //   row 4 = c, the promoted row, with its ParentId cleared
            //   and its SiblingOrder its position in the top-level run:
            //   a is 0 (row 1), c is 1 (row 4) (R5.3 D2, amended 2026-10-11).
            Assert.Equal(Id('b').Value, IdText(table, 2));
            Assert.Equal(Id('d').Value, IdText(table, 3));
            Assert.Equal(Id('c').Value, IdText(table, 4));

            // The promoted row left its parent.
            Assert.Null(ParentText(table, 4));
            // The remaining run resequenced densely from the top.
            Assert.Equal(0d, OrderValue(table, 2));
            Assert.Equal(1d, OrderValue(table, 3));
            // The promoted row joined the top-level run at its new
            // position, after the only other top-level row (a).
            Assert.Equal(1d, OrderValue(table, 4));

            // The cells written ABOVE the moved window are the exact call the
            // refused accessor sat on. 2 rows x 16 columns plus the two
            // cells above the window (b's run cell and a's top-level run
            // cell) is the proof they landed; a refusal returns 0 cells.
            Assert.Equal((2 * GanttTableSchema.Default.Columns.Count) + 2, outcome.CellsWritten);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// A demote writes the new <c>ParentId</c> and the resequenced
    /// <c>SiblingOrder</c> cells, with no refusal.
    /// </summary>
    /// <remarks>
    /// Demote writes <em>only</em> planned cells, so every write goes through the
    /// single-cell path. A refusal here means the cell accessor is broken.
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Demote_assigns_the_nearest_eligible_parent_and_resequences()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath), "RegisterXLL failed.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            Excel.Worksheet sheet =
                (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];

            var inserter = new ExcelGanttRowInserter(fixture.Excel);
            Assert.True(inserter.Insert(GanttEntityType.AsPlannedActivity, () => Id('a')).Succeeded);
            Assert.True(inserter.Insert(GanttEntityType.AsPlannedActivity, () => Id('b')).Succeeded);
            OrderOf(table, 1, 0);
            OrderOf(table, 2, 1);

            table.ListRows[2].Range.Select();

            ParentChildMutationOutcome outcome = new ExcelParentChildMutator(fixture.Excel).Demote();

            _output.WriteLine(
                $"DEMOTE succeeded={outcome.Succeeded} refusal={outcome.Refusal} " +
                $"planRefusal={outcome.PlanRefusal} cells={outcome.CellsWritten}");

            Assert.True(
                outcome.Succeeded,
                $"Demote refused: refusal={outcome.Refusal}, plan={outcome.PlanRefusal}. " +
                "A refused write is the defect the live gate exists to catch.");
            Assert.Null(outcome.Refusal);
            Assert.Null(outcome.PlanRefusal);
            Assert.True(outcome.CellsWritten > 0, "No cell was written.");

            // Row 2 became row 1's child and took the tail of its new run.
            Assert.Equal(Id('a').Value, ParentText(table, 2));
            Assert.Equal(0d, OrderValue(table, 2));
            Assert.Equal(0d, OrderValue(table, 1));
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    private static void ParentOf(Excel.ListObject table, int bodyRow, string? value) =>
        table.ListColumns["ParentId"].DataBodyRange.Cells[bodyRow, 1].Value2 = value;

    private static void OrderOf(Excel.ListObject table, int bodyRow, int value) =>
        table.ListColumns["SiblingOrder"].DataBodyRange.Cells[bodyRow, 1].Value2 = value;

    private static string? IdText(Excel.ListObject table, int bodyRow) =>
        table.ListColumns["Id"].DataBodyRange.Cells[bodyRow, 1].Value2 as string;

    private static string? ParentText(Excel.ListObject table, int bodyRow) =>
        table.ListColumns["ParentId"].DataBodyRange.Cells[bodyRow, 1].Value2 as string;

    private static double OrderValue(Excel.ListObject table, int bodyRow) =>
        Convert.ToDouble(
            table.ListColumns["SiblingOrder"].DataBodyRange.Cells[bodyRow, 1].Value2,
            System.Globalization.CultureInfo.InvariantCulture);
}
