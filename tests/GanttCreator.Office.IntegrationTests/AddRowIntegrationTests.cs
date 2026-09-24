using GanttCreator.Core;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

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

            var sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            var table = sheet.ListObjects[GanttTableSchema.TableName];
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
            Assert.Equal((int)Excel.XlDVType.xlValidateList, validation.Type);
            Assert.Equal($"={GanttWorkbookContract.TypeOptionsDefinedName}", validation.Formula1);
            Assert.True(validation.InCellDropdown);
            _output.WriteLine("R2.9 TypeOptions name and Type-column validation are present after add-row fallback.");

            AssertRow(table.ListRows[activity.BodyIndex], "As-Planned Activity", "AsPlannedActivity", FixedId('1').Value);
            AssertRow(table.ListRows[milestone.BodyIndex], "As-Planned Milestone", "AsPlannedMilestone", FixedId('2').Value);
            Assert.Equal(3, table.DataBodyRange.Rows.Count);
            Assert.Equal(4, table.Range.Rows.Count);
            AssertRow(table.ListRows[delineator.BodyIndex], "Delineator", "DefaultDelineator", FixedId('3').Value);

            _output.WriteLine("R2.8 appended activity, milestone, and delineator rows; no trailing blank row or shapes changed.");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Insert_first_activity_reuses_the_initial_blank_row()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            var sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];

            Assert.Equal(0, table.ListRows.Count);
            Assert.Equal(2, table.Range.Rows.Count);
            Assert.Null(table.DataBodyRange);

            GanttRowInsertOutcome outcome = new ExcelGanttRowInserter(fixture.Excel).Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('0'));

            Assert.True(outcome.Succeeded, $"Insert refused: {outcome.Refusal}");
            Assert.Equal(1, outcome.BodyIndex);
            Assert.Equal(1, table.ListRows.Count);
            Assert.Equal(2, table.Range.Rows.Count);
            Assert.NotNull(table.DataBodyRange);
            Assert.Equal(1, table.DataBodyRange.Rows.Count);
            Assert.Equal(14, table.DataBodyRange.Columns.Count);
            AssertRow(table.ListRows[1], "As-Planned Activity", "AsPlannedActivity", FixedId('0').Value);
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

            var sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            var table = sheet.ListObjects[GanttTableSchema.TableName];
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
