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
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];
            var shapesBefore = sheet.Shapes.Count;
            var inserter = new ExcelGanttRowInserter(fixture.Excel);

            GanttRowInsertOutcome activity = inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('1'));
            var activityIndex = RequiredBodyIndex(activity);
            table.ListRows[activityIndex].Range.Select();
            GanttRowInsertOutcome milestone = inserter.Insert(
                GanttEntityType.AsPlannedMilestone,
                () => FixedId('2'));
            var milestoneIndex = RequiredBodyIndex(milestone);
            table.ListRows[milestoneIndex].Range.Select();
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
    public async Task Insert_active_row_places_new_row_below_it_and_shifts_following_rows()
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
            var inserter = new ExcelGanttRowInserter(fixture.Excel);

            Assert.True(inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('1')).Succeeded);
            table.ListRows[1].Range.Select();
            Assert.True(inserter.Insert(
                GanttEntityType.AsPlannedMilestone,
                () => FixedId('2')).Succeeded);
            table.ListRows[2].Range.Select();
            Assert.True(inserter.Insert(
                GanttEntityType.Delineator,
                () => FixedId('3')).Succeeded);

            table.ListRows[2].Range.Select();
            GanttRowInsertOutcome inserted = inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('4'));

            Assert.True(inserted.Succeeded, $"Insert refused: {inserted.Refusal}");
            Assert.Equal(2, inserted.BodyIndex);
            Assert.Equal(4, table.DataBodyRange.Rows.Count);
            Assert.Equal(5, table.Range.Rows.Count);
            AssertId(table.ListRows[1], FixedId('1').Value);
            AssertId(table.ListRows[2], FixedId('4').Value);
            AssertId(table.ListRows[3], FixedId('2').Value);
            AssertId(table.ListRows[4], FixedId('3').Value);
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
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];
            var rowsBefore = table.ListRows.Count;
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

    private static int RequiredBodyIndex(GanttRowInsertOutcome outcome)
    {
        Assert.True(outcome.Succeeded, $"Insert refused: {outcome.Refusal}");
        _ = Assert.NotNull(outcome.BodyIndex);
        return outcome.BodyIndex.Value;
    }

    private static void AssertId(Excel.ListRow row, string id)
    {
        Assert.Equal(id, Convert.ToString(
            row.Range.Cells[1, 1].Value2,
            System.Globalization.CultureInfo.InvariantCulture));
    }

    private static void AssertRow(Excel.ListRow row, string type, string styleKey, string id)
    {
        Excel.Range range = row.Range;
        Assert.Equal(id, Convert.ToString(range.Cells[1, 1].Value2, System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(type, Convert.ToString(range.Cells[1, 4].Value2, System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(styleKey, Convert.ToString(range.Cells[1, 9].Value2, System.Globalization.CultureInfo.InvariantCulture));
    }
}
