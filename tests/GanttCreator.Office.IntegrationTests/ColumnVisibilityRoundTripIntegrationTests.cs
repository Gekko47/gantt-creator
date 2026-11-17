using GanttCreator.Core;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel round trip for the column-visibility choice (R5.18h, closing O3): the host really
/// hides the worksheet column, the Refresh path leaves that choice alone, and the settings key
/// the ribbon writes can be read back.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why live.</b> Three separately-correct pieces meet here and only the host can prove they
/// do. The port writes <c>Range.Hidden</c> on the whole worksheet column; the classification
/// restorer must then decline to force the schema's visibility back over it; and the key must
/// survive a write/read cycle through the same catalogue writer and reader the ribbon uses.
/// The contract tests pin each piece against mocks, where the column ranges are fakes and
/// "hidden" is a recorded callback rather than a host state — so a reconciliation that compares
/// the wrong proxy, or off a chained COM expression, would still be green there.
/// </para>
/// <para>
/// Deliberately no <c>RegisterXll</c>: the packed add-in's background refresh would hide and
/// un-hide underneath these assertions. The three adapters are driven directly, in the order
/// the ribbon and the refresh drive them.
/// </para>
/// </remarks>
public class ColumnVisibilityRoundTripIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_hidden_column_stays_hidden_through_a_refresh_and_the_key_round_trips()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

            using var scope = new OfficeFixture.ComScope();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
            Excel.ListObject table = (Excel.ListObject)scope.Track(
                sheet.ListObjects[GanttTableSchema.TableName]);

            const string hide = "Description";
            var port = new ExcelColumnVisibilityPort(fixture.Excel);

            // ---- Nothing hidden to start with, measured on the host. ----
            ColumnVisibilityPortOutcome before = port.ReadHidden();
            Assert.True(before.Succeeded, $"read refused: {before.Refusal}");
            Assert.Empty(before.Hidden!);

            // ---- The toggle path: the port hides the whole worksheet column. ----
            ColumnVisibilityPortOutcome applied = port.Apply(new HashSet<string>([hide], StringComparer.Ordinal));
            Assert.True(applied.Succeeded, $"apply refused: {applied.Refusal}");
            Assert.Equal(1, applied.ColumnsWritten);

            Excel.ListColumn hiddenColumn = (Excel.ListColumn)scope.Track(table.ListColumns[hide]);
            Excel.Range entireColumn = (Excel.Range)scope.Track(hiddenColumn.Range.EntireColumn);
            Assert.True(
                Convert.ToBoolean(entireColumn.Hidden, System.Globalization.CultureInfo.InvariantCulture),
                "The host does not report the Description column as hidden after Apply.");

            // The read path agrees with the host, which is what seeds the dropdown's ticks.
            ColumnVisibilityPortOutcome readBack = port.ReadHidden();
            Assert.True(readBack.Succeeded, $"read refused: {readBack.Refusal}");
            _ = Assert.Single(readBack.Hidden!);
            Assert.Contains(hide, readBack.Hidden!);
            _output.WriteLine("the toggle hid the worksheet column and the port reads it back");

            // ---- A Refresh must not put it back. ----
            // This is the reconciliation R5.18h slice 2 exists for: the restorer forces only the
            // ENGINE columns' hidden flags, so a user-visible column keeps the user's choice. A
            // restorer that still forced every column would leave this assertion red.
            ColumnPresentationOutcome refreshed = new ExcelColumnPresentationRestorer(fixture.Excel)
                .EnsureClassification();
            Assert.True(refreshed.Succeeded, $"the classification restorer refused: {refreshed.Refusal}");

            Excel.Range afterRefresh = (Excel.Range)scope.Track(hiddenColumn.Range.EntireColumn);
            Assert.True(
                Convert.ToBoolean(afterRefresh.Hidden, System.Globalization.CultureInfo.InvariantCulture),
                "A Refresh un-hid a column the user chose to hide.");

            // The engine columns are still forced hidden, which is the other half of the rule:
            // the reconciliation skips user-visible columns, it does not stop classifying.
            Excel.ListColumn engineColumn = (Excel.ListColumn)scope.Track(table.ListColumns["Id"]);
            Excel.Range engineEntire = (Excel.Range)scope.Track(engineColumn.Range.EntireColumn);
            Assert.True(
                Convert.ToBoolean(engineEntire.Hidden, System.Globalization.CultureInfo.InvariantCulture),
                "An engine column is visible after a Refresh, so the classification stopped working.");
            _output.WriteLine("the Refresh kept the user's column hidden and still forced the engine columns");

            // ---- The choice is persistable: the same writer and reader the ribbon uses. ----
            string stored = GanttColumnVisibility.Format(new HashSet<string>([hide], StringComparer.Ordinal));
            ConfigWriteOutcome written = new ExcelConfigCatalogueWriter(fixture.Excel).WriteSettings(
                new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    [GanttColumnVisibility.SettingKey] = stored,
                });
            Assert.True(written.Succeeded, $"the settings write refused: {written.Refusal}");

            ConfigReadOutcome reread = new ExcelConfigCatalogueReader(fixture.Excel).Read();
            Assert.True(reread.Succeeded, $"the settings read refused: {reread.Refusal}");
            Assert.True(
                reread.Settings.TryGetValue(GanttColumnVisibility.SettingKey, out string? value),
                "The stored column visibility key did not survive a write/read cycle.");
            Assert.Equal(stored, value, StringComparer.Ordinal);

            ColumnVisibilityOutcome parsed = GanttColumnVisibility.Parse(value);
            Assert.True(parsed.Succeeded, $"the stored value did not parse: {parsed.Refusal}");
            Assert.Contains(hide, parsed.Hidden!);
            _output.WriteLine("the stored key round-tripped and parses back to the same choice");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// The all-hidden state survives the host, the classification restorer and the live
    /// measurement (ADR-0065 D1/D4).
    /// </summary>
    /// <remarks>
    /// This is the live counterpart of the probe: it hides all six user columns through the
    /// real port, proves the restorer leaves the user's choice alone, and proves the
    /// zero-column measurement succeeds with the anchor column collapsed onto the table
    /// origin. The orchestrator's full Refresh is the Office gate's job; this fact covers the
    /// three adapters the Refresh depends on.
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Every_user_column_can_be_hidden_and_the_zero_column_measurement_succeeds()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

            using var scope = new OfficeFixture.ComScope();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
            Excel.ListObject table = (Excel.ListObject)scope.Track(
                sheet.ListObjects[GanttTableSchema.TableName]);

            // Seed one body row: a freshly initialised table has a header and no body rows,
            // so DataBodyRange is legitimately null before any hiding happens (see
            // PanelGridMeasurementIntegrationTests). The probe exercised a two-row body, so
            // the "the body survives hiding every column" claim needs a body to survive.
            _ = scope.Track(table.ListRows.Add());

            var port = new ExcelColumnVisibilityPort(fixture.Excel);
            var allHidden = new HashSet<string>(
                GanttColumnVisibility.ToggleableColumnNames,
                StringComparer.Ordinal);

            ColumnVisibilityPortOutcome applied = port.Apply(allHidden);
            Assert.True(applied.Succeeded, $"apply refused: {applied.Refusal}");
            Assert.Equal(GanttColumnVisibility.ToggleableColumnNames.Count, applied.ColumnsWritten);

            foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
            {
                Excel.ListColumn column = (Excel.ListColumn)scope.Track(table.ListColumns[name]);
                Excel.Range entire = (Excel.Range)scope.Track(column.Range.EntireColumn);
                Assert.True(
                    Convert.ToBoolean(entire.Hidden, System.Globalization.CultureInfo.InvariantCulture),
                    $"{name} is visible after the all-hidden apply.");
                Assert.Equal(
                    0d,
                    Convert.ToDouble(entire.Width, System.Globalization.CultureInfo.InvariantCulture),
                    precision: 6);
            }

            ColumnVisibilityPortOutcome readBack = port.ReadHidden();
            Assert.True(readBack.Succeeded, $"read refused: {readBack.Refusal}");
            Assert.Equal(allHidden, readBack.Hidden!);

            // The classification restorer is the Refresh's presentation pass: it must leave
            // the user columns alone and still force the engine columns hidden.
            ColumnPresentationOutcome refreshed = new ExcelColumnPresentationRestorer(fixture.Excel)
                .EnsureClassification();
            Assert.True(refreshed.Succeeded, $"the classification restorer refused: {refreshed.Refusal}");
            foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
            {
                Excel.ListColumn column = (Excel.ListColumn)scope.Track(table.ListColumns[name]);
                Excel.Range entire = (Excel.Range)scope.Track(column.Range.EntireColumn);
                Assert.True(
                    Convert.ToBoolean(entire.Hidden, System.Globalization.CultureInfo.InvariantCulture),
                    $"{name} was un-hidden by the classification restorer.");
            }

            // The live measurement must accept the empty included-column list, and the anchor
            // column's left edge must equal the table's origin (the probe measured both).
            PanelGridOutcome measured = new ExcelPanelGridMeasurement(fixture.Excel).Measure([]);
            Assert.True(measured.Succeeded, $"the zero-column measurement refused: {measured.Refusal}");
            Assert.Empty(measured.Grid!.Columns);

            Excel.Range header = (Excel.Range)scope.Track(table.HeaderRowRange);
            Excel.Range anchor = (Excel.Range)scope.Track(
                sheet.Columns[GanttSheetLayout.PlotAnchorColumnIndex]);
            Assert.Equal(
                Convert.ToDouble(header.Left, System.Globalization.CultureInfo.InvariantCulture),
                Convert.ToDouble(anchor.Left, System.Globalization.CultureInfo.InvariantCulture),
                precision: 2);
            Assert.True(
                table.DataBodyRange is not null,
                "The table lost its body when every user column was hidden.");

            _output.WriteLine("all six user columns hid, the restorer kept them hidden, and the zero-column measurement succeeded");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
