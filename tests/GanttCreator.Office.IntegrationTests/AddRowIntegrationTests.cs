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
            Assert.True(activity.BodyIndex > 0);
            Assert.True(milestone.BodyIndex > activity.BodyIndex);
            Assert.True(delineator.BodyIndex > milestone.BodyIndex);
            Assert.True(table.ListRows.Count >= delineator.BodyIndex);
            Assert.Equal(shapesBefore, sheet.Shapes.Count);
            Assert.Contains(
                workbook.Names.Cast<Excel.Name>(),
                name => string.Equals(name.Name, GanttWorkbookContract.TypeOptionsDefinedName, StringComparison.OrdinalIgnoreCase));
            Excel.Name typeOptions = workbook.Names.Item(GanttWorkbookContract.TypeOptionsDefinedName);
            Assert.Contains(GanttWorkbookContract.ConfigSheetName, typeOptions.RefersTo, StringComparison.Ordinal);
            Assert.Contains("$B$2:$B$17", typeOptions.RefersTo, StringComparison.Ordinal);
            Excel.ListColumn typeColumn = table.ListColumns["Type"];
            Assert.NotNull(typeColumn.DataBodyRange);
            Excel.Validation validation = typeColumn.DataBodyRange.Validation;
            Assert.Equal((int)Excel.XlDVType.xlValidateList, (int)validation.Type);
            Assert.Equal($"={GanttWorkbookContract.TypeOptionsDefinedName}", validation.Formula1);
            Assert.True(validation.InCellDropdown);
            _output.WriteLine("R2.9 TypeOptions name and Type-column validation are present after add-row fallback.");

            AssertRow(table.ListRows[activity.BodyIndex], "As-Planned Activity", "AsPlannedActivity", FixedId('1').Value);
            AssertRow(table.ListRows[milestone.BodyIndex], "As-Planned Milestone", "AsPlannedMilestone", FixedId('2').Value);
            AssertRow(table.ListRows[delineator.BodyIndex], "Delineator", "DefaultDelineator", FixedId('3').Value);

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
