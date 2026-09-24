using System.Globalization;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;
using GanttCreator.Office;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>Live-Excel acceptance tests for the R2.10 safe-repair workflow.</summary>
public class ConfigRepairIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    private static string ValidId(char suffix) =>
        $"G-0123456789abcdef0123456789abcde{suffix}";

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Repair_restores_visibility_plot_anchor_and_typeoptions()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath), $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            Excel.Worksheet gantt = FindSheet(workbook, GanttWorkbookContract.GanttSheetLabel);
            Excel.Worksheet config = FindSheet(workbook, GanttWorkbookContract.ConfigSheetName);
            int sheetCountBefore = workbook.Sheets.Count;

            config.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            gantt.Names.Item(GanttWorkbookContract.PlotAnchorDefinedName).Delete();
            workbook.Names.Item(GanttWorkbookContract.TypeOptionsDefinedName).Delete();

            ExcelConfigIntegrityChecker checker = new(fixture.Excel);
            ConfigIntegrityCheckOutcome check = checker.Check();
            ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build(check.Findings);
            _output.WriteLine(
                $"Findings before repair: {string.Join(", ", plan.Findings.Select(finding => finding.Kind))}");

            Assert.Contains(plan.Findings, finding => finding.Kind == ConfigIntegrityFindingKind.WrongVisibility);
            Assert.Contains(plan.Findings, finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement);
            Assert.Contains(plan.Findings, finding => finding.Kind == ConfigIntegrityFindingKind.TypeOptionsMissing);
            Assert.Empty(plan.RefusedFindings);

            ConfigRepairOutcome outcome = new ExcelConfigRepairer(fixture.Excel).Repair(plan, true);

            Assert.Null(outcome.Refusal);
            Assert.Equal(Excel.XlSheetVisibility.xlSheetVeryHidden, config.Visible);
            Assert.Equal(
                $"='{GanttWorkbookContract.GanttSheetLabel}'!$O$1",
                gantt.Names.Item(GanttWorkbookContract.PlotAnchorDefinedName).RefersTo,
                StringComparer.Ordinal);
            Assert.Equal(
                GanttWorkbookContract.TypeOptionsDefinedName,
                workbook.Names.Item(GanttWorkbookContract.TypeOptionsDefinedName).Name,
                StringComparer.OrdinalIgnoreCase);
            Assert.Equal(sheetCountBefore, workbook.Sheets.Count);
            Assert.Equal(1, CountVisibleWorksheets(workbook));
            Assert.Equal(1, CountVeryHiddenWorksheets(workbook));
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Repair_confirmed_identity_damage_changes_only_duplicate_id_cells()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath), $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            Excel.Worksheet gantt = FindSheet(workbook, GanttWorkbookContract.GanttSheetLabel);
            Excel.ListObject table = gantt.ListObjects[GanttTableSchema.TableName];
            Excel.ListRow first = table.ListRows.Add();
            Excel.ListRow second = table.ListRows.Add();
            WriteActivityRow(first, ValidId('1'), "First user row");
            WriteActivityRow(second, ValidId('1'), "Second user row");
            string firstDescription = Convert.ToString(first.Range.Cells[1, 5].Value2, CultureInfo.InvariantCulture) ?? string.Empty;
            string secondDescription = Convert.ToString(second.Range.Cells[1, 5].Value2, CultureInfo.InvariantCulture) ?? string.Empty;
            int sheetCountBefore = workbook.Sheets.Count;

            ConfigIntegrityCheckOutcome check = new ExcelConfigIntegrityChecker(fixture.Excel).Check();
            ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build(check.Findings);
            _output.WriteLine(
                $"Identity findings before repair: {string.Join(", ", plan.Findings.Select(finding => finding.Kind))}");

            Assert.Contains(plan.Findings, finding => finding.Kind == ConfigIntegrityFindingKind.IdentityDamage);
            Assert.DoesNotContain(plan.RefusedFindings, finding => finding.Kind == ConfigIntegrityFindingKind.IdentityDamage);

            ConfigRepairOutcome declined = new ExcelConfigRepairer(fixture.Excel).Repair(plan, false);
            Assert.Null(declined.Refusal);
            Assert.Equal(2, declined.ConfirmationSkippedCount);
            Assert.Equal(
                ValidId('1'),
                Convert.ToString(second.Range.Cells[1, 1].Value2, CultureInfo.InvariantCulture),
                StringComparer.Ordinal);

            ConfigRepairOutcome repaired = new ExcelConfigRepairer(fixture.Excel).Repair(plan, true);
            Assert.Null(repaired.Refusal);
            string repairedFirst = Convert.ToString(first.Range.Cells[1, 1].Value2, CultureInfo.InvariantCulture) ?? string.Empty;
            string repairedSecond = Convert.ToString(second.Range.Cells[1, 1].Value2, CultureInfo.InvariantCulture) ?? string.Empty;
            Assert.Equal(ValidId('1'), repairedFirst, StringComparer.Ordinal);
            Assert.True(GanttRowId.TryParse(repairedSecond, out _));
            Assert.NotEqual(repairedFirst, repairedSecond, StringComparer.Ordinal);
            Assert.Equal(
                firstDescription,
                Convert.ToString(first.Range.Cells[1, 5].Value2, CultureInfo.InvariantCulture),
                StringComparer.Ordinal);
            Assert.Equal(
                secondDescription,
                Convert.ToString(second.Range.Cells[1, 5].Value2, CultureInfo.InvariantCulture),
                StringComparer.Ordinal);
            Assert.Equal(sheetCountBefore, workbook.Sheets.Count);
            Assert.Equal(1, CountVisibleWorksheets(workbook));
            Assert.Equal(1, CountVeryHiddenWorksheets(workbook));
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static void WriteActivityRow(Excel.ListRow row, string id, string description)
    {
        Excel.Range range = row.Range;
        range.Cells[1, 1].Value2 = id;
        range.Cells[1, 2].Value2 = ValidId('a');
        range.Cells[1, 3].Value2 = 0;
        range.Cells[1, 4].Value2 = "As-Planned Activity";
        range.Cells[1, 5].Value2 = description;
        range.Cells[1, 6].Value2 = new DateTime(2026, 9, 1).ToOADate();
        range.Cells[1, 7].Value2 = new DateTime(2026, 9, 2).ToOADate();
        range.Cells[1, 13].Value2 = true;
    }

    private static Excel.Worksheet FindSheet(Excel.Workbook workbook, string name) =>
        workbook.Sheets
            .Cast<Excel.Worksheet>()
            .First(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase));

    private static int CountVisibleWorksheets(Excel.Workbook workbook) =>
        workbook.Sheets.Cast<Excel.Worksheet>().Count(sheet => sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible);

    private static int CountVeryHiddenWorksheets(Excel.Workbook workbook) =>
        workbook.Sheets.Cast<Excel.Worksheet>().Count(sheet => sheet.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden);
}
