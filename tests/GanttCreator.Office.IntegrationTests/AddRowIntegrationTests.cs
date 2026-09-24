using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Office;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>Live-Excel acceptance gate for the R2.8 add-row commands.</summary>
public class AddRowIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    private static GanttRowId FixedId(char suffix) =>
        GanttRowId.Parse("G-0123456789abcdef0123456789abcde" + suffix);

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Insert_appends_all_three_command_rows_without_rendering()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            ExcelWorkbookInitialiser initialiser = new(fixture.Excel);
            WorkbookInitialiseOutcome initialised = initialiser.Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];
            int shapesBefore = sheet.Shapes.Count;
            var inserter = new ExcelGanttRowInserter(fixture.Excel);

            GanttRowInsertOutcome activity = inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('1'));
            GanttRowInsertOutcome milestone = inserter.Insert(
                GanttEntityType.AsPlannedMilestone,
                () => FixedId('2'));
            GanttRowInsertOutcome delineator = inserter.Insert(
                GanttEntityType.Delineator,
                () => FixedId('3'));

            Assert.True(activity.Succeeded);
            Assert.True(milestone.Succeeded);
            Assert.True(delineator.Succeeded);
            Assert.Equal(1, activity.BodyIndex);
            Assert.Equal(2, milestone.BodyIndex);
            Assert.Equal(3, delineator.BodyIndex);
            Assert.Equal(3, table.ListRows.Count);
            Assert.Equal(shapesBefore, sheet.Shapes.Count);

            AssertRow(table.ListRows[1], "As-Planned Activity", "AsPlannedActivity", FixedId('1').Value);
            AssertRow(table.ListRows[2], "As-Planned Milestone", "AsPlannedMilestone", FixedId('2').Value);
            AssertRow(table.ListRows[3], "Delineator", "DefaultDelineator", FixedId('3').Value);

            _output.WriteLine("R2.8 appended activity, milestone, and delineator rows; no shapes changed.");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Insert_refuses_a_protected_sheet_without_adding_a_row()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];
            int rowsBefore = table.ListRows.Count;
            sheet.Protect(Type.Missing, true, Type.Missing, true, true, false, false);

            GanttRowInsertOutcome outcome = new ExcelGanttRowInserter(fixture.Excel).Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('4'));

            Assert.False(outcome.Succeeded);
            Assert.Equal(GanttRowInsertRefusalReason.TargetProtected, outcome.Refusal);
            Assert.Equal(rowsBefore, table.ListRows.Count);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static void AssertRow(Excel.ListRow row, string type, string styleKey, string id)
    {
        Excel.Range range = row.Range;
        Assert.Equal(id, Convert.ToString(range.Cells[1, 1].Value2, System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(type, Convert.ToString(range.Cells[1, 4].Value2, System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(styleKey, Convert.ToString(range.Cells[1, 9].Value2, System.Globalization.CultureInfo.InvariantCulture));
    }
}
