using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Office;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the workbook initialiser (work item R2.2). Tagged
/// <c>[Trait("Category","OfficeIntegration")]</c> so <c>verify-quick.ps1</c>
/// and <c>verify.ps1</c> exclude them; run via
/// <c>pwsh ./scripts/verify-office.ps1</c> on the self-hosted runner.
/// </summary>
/// <remarks>
/// Driven through the <see cref="IWorkbookInitialiser"/> port rather than the
/// AddIn command layer: this project references Core and Office but not
/// AddIn, so the command boundary stays behind the AddIn assembly. The
/// packed XLL is still loaded so the add-in's production adapter runs inside
/// Excel; the contract assertions mirror <see cref="WorkbookInitialiserTests"/>
/// (Moq PIA, no live Office) so a drift between the port and the live path is
/// caught. The typed-refusal translation itself is gated by
/// <c>InitialiseSheetCommandTests</c> (AddIn contract, no Office).
/// </remarks>
public class InitialiseSheetIntegrationTests(ITestOutputHelper output)
    {
        private readonly ITestOutputHelper _output = output;

        private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    /// <summary>
    /// The adopt path: a blank workbook created through the fixture becomes
    /// the supported state — one visible worksheet named <c>Gantt Data</c>
    /// carrying <c>tblGanttData</c>, the sheet-scoped plot anchor, and one
    /// VeryHidden <c>_GanttCreatorConfig</c> sheet. Nothing else is touched.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Initialise_adopts_a_blank_worksheet_and_builds_the_supported_state()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Excel.Worksheet active = (Excel.Worksheet)workbook.ActiveSheet;

            WorkbookInitialiseOutcome outcome =
                new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            _output.WriteLine($"Initialise path={outcome.Path} sheetName={outcome.SheetName} refusal={outcome.Refusal}");

            Assert.True(outcome.Succeeded);
            Assert.Equal(WorkbookInitialisePath.Adopted, outcome.Path);
            Assert.Equal(
                GanttWorkbookContract.GanttSheetLabel,
                outcome.SheetName,
                StringComparer.OrdinalIgnoreCase);
            Assert.Equal(
                GanttWorkbookContract.GanttSheetLabel,
                active.Name,
                StringComparer.OrdinalIgnoreCase);

            Excel.ListObject table = active.ListObjects[GanttTableSchema.TableName];
            Assert.Equal(GanttTableSchema.TableName, table.Name, StringComparer.Ordinal);
            Assert.Equal(
                GanttTableSchema.Default.Columns.Count,
                table.ListColumns.Count);

            for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
            {
                Assert.Equal(
                    GanttTableSchema.Default.Columns[index].Name,
                    table.ListColumns[index + 1].Name,
                    StringComparer.Ordinal);
            }

            var anchor = workbook.Names.Item(GanttWorkbookContract.PlotAnchorDefinedName);
            Assert.NotNull(anchor);
            Assert.Equal(
                $"='{GanttWorkbookContract.GanttSheetLabel}'!$O$1",
                anchor.RefersTo,
                StringComparer.Ordinal);

            var visibleSheets = workbook.Sheets
                .Cast<Excel.Worksheet>()
                .Count(s => s.Visible == Excel.XlSheetVisibility.xlSheetVisible);
            var veryHiddenSheets = workbook.Sheets
                .Cast<Excel.Worksheet>()
                .Count(s => s.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden);
            Assert.Equal(2, visibleSheets);
            Assert.Equal(1, veryHiddenSheets);
            Assert.Equal(
                GanttWorkbookContract.ConfigSheetName,
                workbook.Sheets[2].Name,
                StringComparer.OrdinalIgnoreCase);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// A workbook that already carries <c>tblGanttData</c> is refused and left
    /// unchanged — the typed-refusal surface exercised by the contract tests,
    /// now against live Excel.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Initialise_refuses_and_mutates_nothing_when_the_table_already_exists()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Excel.Worksheet active = (Excel.Worksheet)workbook.ActiveSheet;
            active.Name = GanttWorkbookContract.GanttSheetLabel;

            // Create a real ListObject named GanttTableSchema.TableName on the
            // renamed worksheet so the TableExists refusal path is exercised.
            Excel.Range headerRange = active.Cells[1, 1].Resize[1, 1];
            Excel.ListObject existingTable = active.ListObjects.Add(
                Excel.XlListObjectSourceType.xlSrcRange,
                headerRange,
                Type.Missing,
                Excel.XlYesNoGuess.xlYes,
                Type.Missing);
            existingTable.Name = GanttTableSchema.TableName;

            WorkbookInitialiseOutcome outcome =
                new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            _output.WriteLine($"Initialise path={outcome.Path} sheetName={outcome.SheetName} refusal={outcome.Refusal}");

            Assert.Equal(
                WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TableExists),
                outcome);
            Assert.Equal(
                GanttWorkbookContract.GanttSheetLabel,
                active.Name,
                StringComparer.OrdinalIgnoreCase);
            Assert.Equal(1, active.ListObjects.Count);
            Assert.Equal(GanttTableSchema.TableName, active.ListObjects[1].Name, StringComparer.Ordinal);
            Assert.Equal(2, workbook.Sheets.Count);
            _output.WriteLine("Refusal path left the workbook with 2 sheets and the existing table unchanged.");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// The non-blank active worksheet is left untouched and a fresh
    /// <c>Gantt Data</c> sheet is created — the create path of the hybrid
    /// sheet-selection rule (work item R2.2 decision D4).
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Initialise_creates_a_new_sheet_and_leaves_the_non_blank_active_one_untouched()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Excel.Worksheet active = (Excel.Worksheet)workbook.ActiveSheet;
            active.Range["A1", "C1"].Value2 = new object[3] { "pre", "existing", "content" };

            WorkbookInitialiseOutcome outcome =
                new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            _output.WriteLine($"Initialise path={outcome.Path} sheetName={outcome.SheetName} refusal={outcome.Refusal}");

            Assert.True(outcome.Succeeded);
            Assert.Equal(WorkbookInitialisePath.CreatedNew, outcome.Path);
            Assert.Equal("Sheet1", active.Name, StringComparer.Ordinal);
            Assert.Equal("pre", active.Range["A1"].Value2);

            var ganttSheet = workbook.Sheets
                .Cast<Excel.Worksheet>()
                .First(s => string.Equals(
                    s.Name,
                    GanttWorkbookContract.GanttSheetLabel,
                    StringComparison.OrdinalIgnoreCase));
            Excel.ListObject table = ganttSheet.ListObjects[GanttTableSchema.TableName];
            Assert.Equal(GanttTableSchema.TableName, table.Name, StringComparer.Ordinal);
            Assert.Equal(14, table.ListColumns.Count);

            var visibleSheets = workbook.Sheets
                .Cast<Excel.Worksheet>()
                .Count(s => s.Visible == Excel.XlSheetVisibility.xlSheetVisible);
            var veryHiddenSheets = workbook.Sheets
                .Cast<Excel.Worksheet>()
                .Count(s => s.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden);
            Assert.Equal(2, visibleSheets);
            Assert.Equal(1, veryHiddenSheets);
            Assert.Equal(3, workbook.Sheets.Count);
            _output.WriteLine("Create path left the non-blank active sheet untouched and added Gantt Data + config.");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
