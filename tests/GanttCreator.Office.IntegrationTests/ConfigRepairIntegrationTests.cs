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

            // ADR-0055: the loaded pack's background refresh would protect the sheet and write
            // over the repair this test asserts.
            OfficeFixture.DisableChartProtection(workbook);

            Excel.Worksheet gantt = FindSheet(workbook, GanttWorkbookContract.GanttSheetLabel);
            Excel.Worksheet config = FindSheet(workbook, GanttWorkbookContract.ConfigSheetName);
            Excel.ListObject ganttTable = gantt.ListObjects[GanttTableSchema.TableName];
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
            // BOTH coordinates come from the layout authority, as in
            // InitialiseSheetIntegrationTests: a literal column letter went stale when
            // R4.7A added SiblingOrder, a literal ROW went stale when ADR-0030 added the
            // reserved row, and adding the live column count to worksheet column 1 went
            // stale when R5.18k moved the table to column B -- the live gate then failed
            // naming $R$3 against the product's correct $S$3.
            Excel.Range anchorCell =
                gantt.Cells[GanttSheetLayout.HeaderRowIndex, GanttSheetLayout.PlotAnchorColumnIndex];
            Assert.Equal(
                $"='{GanttWorkbookContract.GanttSheetLabel}'!{anchorCell.Address[External: false]}",
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
            OfficeFixture.DisableChartProtection(workbook);

            Excel.Worksheet gantt = FindSheet(workbook, GanttWorkbookContract.GanttSheetLabel);
            Excel.ListObject table = gantt.ListObjects[GanttTableSchema.TableName];
            Excel.ListRow first = table.ListRows.Add();
            Excel.ListRow second = table.ListRows.Add();
            WriteActivityRow(first, table, ValidId('1'), "First user row");
            WriteActivityRow(second, table, ValidId('1'), "Second user row");
            string firstDescription = ReadCell(first, table, "Description");
            string secondDescription = ReadCell(second, table, "Description");
            int sheetCountBefore = workbook.Sheets.Count;

            ConfigIntegrityCheckOutcome check = new ExcelConfigIntegrityChecker(fixture.Excel).Check();
            ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build(check.Findings);
            _output.WriteLine(
                $"Identity findings before repair: {string.Join(", ", plan.Findings.Select(finding => finding.Kind))}");

            Assert.Contains(plan.Findings, finding => finding.Kind == ConfigIntegrityFindingKind.IdentityDamage);
            Assert.DoesNotContain(plan.RefusedFindings, finding => finding.Kind == ConfigIntegrityFindingKind.IdentityDamage);

            ConfigRepairOutcome declined = new ExcelConfigRepairer(fixture.Excel).Repair(plan, false);
            Assert.Null(declined.Refusal);

            // ConfirmationSkippedCount counts CONFIRM-REQUIRED FINDINGS, not the
            // cells that would be rewritten: ExcelConfigRepairer returns
            // plan.ConfirmRequiredFindings.Count when the caller declines. This
            // fixture produces two findings — a PlotAnchorDisagreement and an
            // IdentityDamage — and ADR-0011's safe-repair boundary classifies only
            // IdentityDamage as ConfirmRequired; PlotAnchorDisagreement is
            // AutoRepair because it is a deterministic, safe repair. So the plan
            // has exactly one confirm-required finding. The single rewritten cell
            // asserted below is a different quantity and is unaffected by this.
            _ = Assert.Single(plan.ConfirmRequiredFindings);
            Assert.Equal(1, declined.ConfirmationSkippedCount);
            Assert.Equal(ValidId('1'), ReadCell(second, table, "Id"), StringComparer.Ordinal);

            ConfigRepairOutcome repaired = new ExcelConfigRepairer(fixture.Excel).Repair(plan, true);
            Assert.Null(repaired.Refusal);
            string repairedFirst = ReadCell(first, table, "Id");
            string repairedSecond = ReadCell(second, table, "Id");
            Assert.Equal(ValidId('1'), repairedFirst, StringComparer.Ordinal);
            Assert.True(GanttRowId.TryParse(repairedSecond, out _));
            Assert.NotEqual(repairedFirst, repairedSecond, StringComparer.Ordinal);
            Assert.Equal(firstDescription, ReadCell(first, table, "Description"), StringComparer.Ordinal);
            Assert.Equal(secondDescription, ReadCell(second, table, "Description"), StringComparer.Ordinal);
            Assert.Equal(sheetCountBefore, workbook.Sheets.Count);
            Assert.Equal(1, CountVisibleWorksheets(workbook));
            Assert.Equal(1, CountVeryHiddenWorksheets(workbook));
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// Writes one valid activity row into the live table, addressed by COLUMN NAME.
    /// </summary>
    /// <param name="row">The list row to write.</param>
    /// <param name="table">The table the row belongs to, used to resolve each header's index.</param>
    /// <param name="id">The row identity to write (deliberately duplicated by the fixture).</param>
    /// <param name="description">The description text.</param>
    /// <remarks>
    /// <b>Why by name.</b> This helper used to write one-based positions
    /// (<c>Cells[1, 4]</c> was <c>Type</c> before R5.18k). The reorder that led the table
    /// with its visible columns left every positional literal pointing at a different
    /// column, so the fixture wrote an id into <c>Type</c> and a boolean into
    /// <c>DateLabelPosition</c>; the integrity checker then produced a second,
    /// unrelated identity finding and the repair assertions measured the wrong cells.
    /// The schema is the authority for order, so the name is the only stable address.
    /// </remarks>
    private static void WriteActivityRow(
        Excel.ListRow row,
        Excel.ListObject table,
        string id,
        string description)
    {
        WriteCell(row, table, "Id", id);
        WriteCell(row, table, "LaneId", ValidId('a'));
        WriteCell(row, table, "StackIndex", 0);
        WriteCell(row, table, "Type", "As-Planned Activity");
        WriteCell(row, table, "Description", description);
        WriteCell(row, table, "Start", new DateTime(2026, 9, 1).ToOADate());
        WriteCell(row, table, "Finish", new DateTime(2026, 9, 2).ToOADate());
        WriteCell(row, table, "Visible", true);
    }

    private static void WriteCell(Excel.ListRow row, Excel.ListObject table, string columnName, object value)
    {
        Excel.Range range = row.Range;
        range.Cells[1, ColumnIndex(table, columnName)].Value2 = value;
    }

    private static string ReadCell(Excel.ListRow row, Excel.ListObject table, string columnName)
    {
        Excel.Range range = row.Range;
        return Convert.ToString(
            range.Cells[1, ColumnIndex(table, columnName)].Value2,
            CultureInfo.InvariantCulture) ?? string.Empty;
    }

    /// <summary>Resolves a table column's one-based index from its header name.</summary>
    private static int ColumnIndex(Excel.ListObject table, string columnName)
    {
        int schemaIndex = GanttTableSchema.Default.Columns
            .Select((column, position) => (column.Name, Position: position))
            .Single(entry => string.Equals(entry.Name, columnName, StringComparison.Ordinal))
            .Position + 1;
        int liveIndex = table.ListColumns[columnName].Index;
        Assert.Equal(schemaIndex, liveIndex);
        return liveIndex;
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
