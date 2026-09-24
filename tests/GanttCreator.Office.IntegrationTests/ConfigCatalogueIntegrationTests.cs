using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Office;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the R2.7 configuration catalogues. Tagged
/// <c>[Trait("Category","OfficeIntegration")]</c> so <c>verify-quick.ps1</c>
/// and <c>verify.ps1</c> exclude them; run via
/// <c>pwsh ./scripts/verify-office.ps1</c> on a host with Excel installed.
/// </summary>
/// <remarks>
/// This is the only place the catalogue contract meets real COM: the contract
/// tests substitute the parameterised-property seams, so the real
/// <c>ListObjects.Add</c> / <c>DataBodyRange.Value2</c> behaviour, the
/// VeryHidden sheet's writeability, and persistence across save-and-reopen are
/// evidenced here and nowhere else. The reopen half is the requirement that
/// the catalogues are workbook state, not in-memory state (ADR-0007 D1/D2).
/// </remarks>
public class ConfigCatalogueIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    /// <summary>
    /// The four catalogue tables the reader validates in full, with the
    /// expected header row and body-row count. <c>tblGanttConfig</c> is
    /// asserted separately because its values are per-workbook (workbook ID,
    /// catalogue hash, versions) rather than fixed catalogue rows.
    /// </summary>
    private static readonly (string TableName, string[] Headers, int RowCount)[] CatalogueTables =
    [
        (GanttCatalogues.TypesTableName, GanttCatalogues.TypesHeaders, GanttCatalogues.TypeRows.Count),
        (GanttCatalogues.StylesTableName, GanttCatalogues.StylesHeaders, GanttCatalogues.StylePresets.Count),
        (GanttCatalogues.MetricsTableName, GanttCatalogues.MetricsHeaders, GanttCatalogues.Metrics.Count),
        (GanttCatalogues.SettingsTableName, GanttCatalogues.SettingsHeaders, GanttCatalogues.Settings.Count),
    ];

    /// <summary>
    /// Initialise materialises all five catalogue tables on the VeryHidden
    /// configuration worksheet, the reader validates them, and both survive a
    /// save-and-reopen round trip with the workbook ID unchanged.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Initialise_materialises_the_catalogues_which_survive_save_and_reopen()
    {
        var fixture = new OfficeFixture();
        var path = Path.Combine(
            Path.GetTempPath(),
            string.Create(CultureInfo.InvariantCulture, $"GanttCreator-R27-{Guid.NewGuid():N}.xlsx"));
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(
                fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            WorkbookInitialiseOutcome outcome =
                new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            _output.WriteLine($"Initialise path={outcome.Path} sheetName={outcome.SheetName} refusal={outcome.Refusal}");
            Assert.True(outcome.Succeeded);

            Excel.Worksheet config = FindSheet(workbook, GanttWorkbookContract.ConfigSheetName);
            Assert.Equal(Excel.XlSheetVisibility.xlSheetVeryHidden, config.Visible);

            // The visible data panel keeps the neutral appearance (ADR-0007 D8).
            Excel.Worksheet gantt = FindSheet(workbook, GanttWorkbookContract.GanttSheetLabel);
            Excel.ListObject dataTable = gantt.ListObjects[GanttTableSchema.TableName];
            Assert.False(dataTable.ShowAutoFilter);
            Assert.False(dataTable.ShowTableStyleRowStripes);
            Assert.False(dataTable.ShowTableStyleColumnStripes);

            Assert.Equal(5, config.ListObjects.Count);
            foreach ((var tableName, string[] headers, var rowCount) in CatalogueTables)
            {
                AssertCatalogueTable(config, tableName, headers, rowCount);
            }

            AssertConfigMetadataTable(config);

            ConfigReadOutcome read = new ExcelConfigCatalogueReader(fixture.Excel).Read();
            _output.WriteLine($"Read (in memory): succeeded={read.Succeeded} refusal={read.Refusal}");
            Assert.True(read.Succeeded);
            Assert.NotNull(read.WorkbookId);
            Assert.Equal(GanttCatalogues.Settings.Count, read.Settings.Count);

            // Persistence: the catalogues are workbook state, so they must
            // satisfy the same contract after the workbook is saved and
            // reopened, with a stable workbook ID (ADR-0007 D1/D2).
            workbook.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook);

            Excel.Workbooks workbooks = fixture.Excel.Workbooks;
            Excel.Workbook reopened = workbooks.Open(path);
            _output.WriteLine($"Reopened '{reopened.FullName}'.");

            Excel.Worksheet reopenedConfig = FindSheet(reopened, GanttWorkbookContract.ConfigSheetName);
            Assert.Equal(5, reopenedConfig.ListObjects.Count);
            foreach ((var tableName, string[] headers, var rowCount) in CatalogueTables)
            {
                AssertCatalogueTable(reopenedConfig, tableName, headers, rowCount);
            }

            ConfigReadOutcome reread = new ExcelConfigCatalogueReader(fixture.Excel).Read();
            _output.WriteLine($"Read (reopened): succeeded={reread.Succeeded} refusal={reread.Refusal}");
            Assert.True(reread.Succeeded);
            Assert.Equal(read.WorkbookId, reread.WorkbookId);
            Assert.Equal(read.Settings.Count, reread.Settings.Count);
        }
        finally
        {
            // The fixture closes every open workbook (tracked and untracked)
            // and quits Excel; deleting the temp workbook only after that
            // avoids a file-lock race with the owning process.
            await fixture.DisposeAsync().ConfigureAwait(true);
            TryDelete(path);
        }
    }

    /// <summary>
    /// Asserts one catalogue table exists by contract name, with the contract
    /// header row in order and the contract body-row count.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="tableName">The contract table name.</param>
    /// <param name="headers">The contract header row.</param>
    /// <param name="rowCount">The expected body-row count.</param>
    private void AssertCatalogueTable(
        Excel.Worksheet config,
        string tableName,
        string[] headers,
        int rowCount)
    {
        Excel.ListObjects listObjects = config.ListObjects;
        List<string> present = listObjects.Cast<Excel.ListObject>().Select(list => list.Name).ToList();
        Assert.Contains(tableName, present);

        Excel.ListObject table = listObjects[tableName];
        Assert.Equal(tableName, table.Name, StringComparer.Ordinal);

        Excel.ListColumns columns = table.ListColumns;
        Assert.Equal(headers.Length, columns.Count);
        for (var index = 0; index < headers.Length; index++)
        {
            Excel.ListColumn column = columns[index + 1];
            Assert.Equal(headers[index], column.Name, StringComparer.Ordinal);
        }

        Excel.Range body = table.DataBodyRange;
        var actualRows = 0;
        if (body is not null)
        {
            Excel.Range rows = body.Rows;
            actualRows = rows.Count;
        }

        Assert.Equal(rowCount, actualRows);
        _output.WriteLine($"Table '{tableName}': {headers.Length} columns, {actualRows} rows.");
    }

    /// <summary>
    /// Asserts the <c>tblGanttConfig</c> metadata table has the contract
    /// header row and the four approved keys, each carrying a non-empty value.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    private void AssertConfigMetadataTable(Excel.Worksheet config)
    {
        AssertCatalogueTable(config, GanttCatalogues.ConfigTableName, GanttCatalogues.ConfigHeaders, 4);
        Excel.ListObject table = config.ListObjects[GanttCatalogues.ConfigTableName];
        Excel.Range body = table.DataBodyRange;
        Excel.Range rows = body.Rows;
        var rowCount = rows.Count;
        Excel.Range cells = body.Cells;
        List<string> keys = [];
        for (var row = 1; row <= rowCount; row++)
        {
            Excel.Range keyCell = cells[row, 1];
            Excel.Range valueCell = cells[row, 2];
            var key = Convert.ToString(keyCell.Value2, CultureInfo.InvariantCulture) ?? string.Empty;
            var value = Convert.ToString(valueCell.Value2, CultureInfo.InvariantCulture) ?? string.Empty;
            Assert.NotEmpty(key);
            Assert.NotEmpty(value);
            keys.Add(key);
        }

        Assert.Contains(GanttCatalogues.ConfigSchemaVersionKey, keys);
        Assert.Contains(GanttCatalogues.ConfigCatalogueHashKey, keys);
        Assert.Contains(GanttCatalogues.ConfigWorkbookIdKey, keys);
        Assert.Contains(GanttCatalogues.ConfigAddInVersionKey, keys);
        _output.WriteLine($"tblGanttConfig keys: {string.Join(", ", keys)}");
    }

    /// <summary>
    /// Finds a worksheet by name, case-insensitively (Excel's own sheet-name
    /// uniqueness rule).
    /// </summary>
    /// <param name="workbook">The workbook to search.</param>
    /// <param name="name">The sheet name.</param>
    /// <returns>The worksheet.</returns>
    private static Excel.Worksheet FindSheet(Excel.Workbook workbook, string name) =>
        workbook.Sheets
            .Cast<Excel.Worksheet>()
            .First(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Best-effort temp-file cleanup. The file is a test artefact, so a locked
    /// or denied file is reported rather than thrown: it must not mask the
    /// real assertion result.
    /// </summary>
    /// <param name="path">The temp workbook path.</param>
    private void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException ex)
        {
            _output.WriteLine($"Temp workbook not deleted (still locked): {ex.Message}");
        }
        catch (UnauthorizedAccessException ex)
        {
            _output.WriteLine($"Temp workbook not deleted (access denied): {ex.Message}");
        }
    }
}
