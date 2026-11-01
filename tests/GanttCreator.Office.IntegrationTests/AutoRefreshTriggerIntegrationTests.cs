using System.Diagnostics;
using System.Runtime.InteropServices;
using GanttCreator.Core;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>Live Excel checks for the source event and table-column trigger boundary.</summary>
public sealed class AutoRefreshTriggerIntegrationTests
{
    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Production_scheduler_renders_a_table_edit_and_updates_the_live_shape()
    {
        OfficeFixture fixture = new();
        fixture.SuppressLeakSignal = true;
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            using OfficeFixture.ComScope scope = new();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            WorkbookInitialiseOutcome initialized =
                new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialized.Succeeded, $"Initialize refused: {initialized.Refusal}");

            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
            Excel.ListObject table = (Excel.ListObject)scope.Track(
                sheet.ListObjects[GanttTableSchema.TableName]);
            Excel.ListRow row = (Excel.ListRow)scope.Track(table.ListRows.Add());
            Excel.Range rowRange = (Excel.Range)scope.Track(row.Range);
            string rowId = NewId('1');

            SetCell("Id", rowId);
            SetCell("LaneId", NewId('a'));
            SetCell("StackIndex", 0d);
            SetCell("Type", "As-Planned Activity");
            SetCell("Description", "Scheduler initial render");
            SetCell("Start", 46266d);
            SetCell("Finish", 46270d);
            SetCell("Visible", true);

            string labelName = rowId + ":label";
            Assert.Equal(
                "Scheduler initial render",
                await WaitForShapeTextAsync(
                        scope, sheet, labelName, "Scheduler initial render", TimeSpan.FromSeconds(15))
                    .ConfigureAwait(true));

            SetCell("Description", "Scheduler updated render");
            Assert.Equal(
                "Scheduler updated render",
                await WaitForShapeTextAsync(
                        scope, sheet, labelName, "Scheduler updated render", TimeSpan.FromSeconds(15))
                    .ConfigureAwait(true));

            void SetCell(string columnName, object value)
            {
                Excel.ListColumn column = (Excel.ListColumn)scope.Track(table.ListColumns[columnName]);
                Excel.Range cell = (Excel.Range)scope.Track(rowRange.Cells[1, column.Index]);
                cell.Value2 = value;
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static async Task<string> WaitForShapeTextAsync(
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet,
        string shapeName,
        string expectedText,
        TimeSpan timeout)
    {
        long started = Stopwatch.GetTimestamp();
        while (Stopwatch.GetElapsedTime(started) < timeout)
        {
            try
            {
                Excel.Shapes shapes = (Excel.Shapes)scope.Track(sheet.Shapes);
                Excel.Shape shape = (Excel.Shape)scope.Track(shapes.Item(shapeName));
                var textFrame = scope.Track(shape.TextFrame2);
                var textRange = scope.Track(textFrame.TextRange);
                string? text = textRange.Text;
                if (string.Equals(text, expectedText, StringComparison.Ordinal))
                {
                    return text;
                }
            }
            catch (Exception exception) when (exception is COMException or ArgumentException)
            {
                // The expected shape is absent until the queued refresh runs.
                // Shapes.Item(name) reports a missing name as ArgumentException
                // ("The item with the specified name wasn't found"), not as a
                // COMException. Catching only COMException therefore aborted the
                // wait on its very first poll -- before the queued refresh could
                // create the shape -- so this test could never observe a working
                // scheduler.
            }

            await Task.Delay(TimeSpan.FromMilliseconds(25)).ConfigureAwait(true);
        }

        throw new TimeoutException(
            $"The production scheduler did not render '{expectedText}' in shape '{shapeName}' within {timeout}.");
    }

    private static string NewId(char last) => "G-" + new string('0', 31) + last;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Real_table_edit_is_reported_and_classified_as_chart_affecting()
    {
        OfficeFixture fixture = new();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            using OfficeFixture.ComScope scope = new();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            WorkbookInitialiseOutcome initialized =
                new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialized.Succeeded, $"Initialize refused: {initialized.Refusal}");

            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
            Excel.ListObject table = (Excel.ListObject)scope.Track(
                sheet.ListObjects[GanttTableSchema.TableName]);
            Excel.ListColumn description = (Excel.ListColumn)scope.Track(
                table.ListColumns["Description"]);
            Excel.ListRow row = (Excel.ListRow)scope.Track(table.ListRows.Add());
            Excel.Range rowRange = (Excel.Range)scope.Track(row.Range);
            Excel.Range cell = (Excel.Range)scope.Track(rowRange.Cells[1, description.Index]);

            List<string?> received = [];
            ExcelApplicationAdapter adapter = new(fixture.Excel);
            using IDisposable subscription = adapter.SubscribeSheetChanged(received.Add);
            cell.Value2 = "Live event trigger probe";

            string address = Assert.Single(received)!;
            Assert.Contains(workbook.Name + "|" + sheet.Name + "!", address, StringComparison.OrdinalIgnoreCase);
            Assert.True(new ExcelSheetChangeAnalyzer(fixture.Excel).ShouldAutoRefresh(address));
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
