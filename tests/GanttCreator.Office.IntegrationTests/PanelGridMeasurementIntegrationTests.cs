using GanttCreator.Core;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the R4.1 panel-grid measurement port, specifically the
/// <c>Object</c>-typed conversions. Tagged
/// <c>[Trait("Category","OfficeIntegration")]</c> so <c>verify-quick.ps1</c> and
/// <c>verify.ps1</c> exclude it; run via <c>pwsh ./scripts/verify-office.ps1</c>.
/// </summary>
/// <remarks>
/// <para>
/// The contract tests substitute a mocked <c>Range</c>, which proves the
/// adapter's rule but cannot prove the host actually reports
/// <see cref="DBNull"/> rather than a number. Only a live table can settle that,
/// and the answer decides whether the conversion guard is real or theoretical.
/// </para>
/// <para>
/// This test therefore sets two body rows to <em>different</em> heights. A range
/// spanning rows with no single height has no <c>RowHeight</c> to report, so
/// Excel is expected to return <see cref="DBNull"/>. The measurement must then be
/// the typed <see cref="PanelGridRefusalReason.InvalidMeasurement"/> refusal, not
/// the <see cref="InvalidCastException"/> a naive
/// <c>Convert.ToDouble(DBNull.Value)</c> would throw out of a read-only adapter.
/// </para>
/// <para>
/// Read-only: the test writes row heights to set up the condition, but the
/// adapter under test never mutates the workbook.
/// </para>
/// </remarks>
public class PanelGridMeasurementIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_body_with_mixed_row_heights_is_a_typed_refusal_rather_than_a_cast_exception()
    {
        var fixture = new OfficeFixture();
        var scope = new OfficeFixture.ComScope();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            var initialiser = new ExcelWorkbookInitialiser(fixture.Excel);
            WorkbookInitialiseOutcome initialised = initialiser.Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            // Uniform heights first: the measurement must SUCCEED, or the refusal
            // asserted below would prove nothing about mixed heights.
            PanelGridOutcome uniform = new ExcelPanelGridMeasurement(fixture.Excel)
                .Measure([GanttTableSchema.Default.Columns[0].Name]);
            _output.WriteLine("uniform heights: success=" + uniform.Succeeded
                + " reason=" + uniform.Refusal);
            Assert.True(uniform.Succeeded, $"Uniform-height measurement refused: {uniform.Refusal}");

            // Now make the body genuinely mixed. Excel reports no single row height
            // for such a range, which is the DBNull case under test.
            Excel.Sheets sheets = scope.Track(workbook.Sheets);
            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(sheets[GanttWorkbookContract.GanttSheetLabel]);
            Excel.ListObjects objects = scope.Track(sheet.ListObjects);
            Excel.ListObject table = scope.Track(objects[GanttTableSchema.TableName]);
            Excel.Range body = scope.Track(table.DataBodyRange);
            Assert.NotNull(body);

            Excel.Range rows = scope.Track(body.Rows);
            Assert.Equal(2, rows.Count);

            // Two rows, two different heights, written through the row objects so
            // the offsets are the ones the rows actually occupy.
            Excel.Range first = scope.Track(rows.Item[1]);
            Excel.Range second = scope.Track(rows.Item[2]);
            first.RowHeight = 15d;
            second.RowHeight = 45d;

            // Read the host's own answer back rather than assuming it.
            object? reported = body.RowHeight;
            _output.WriteLine($"Excel build {fixture.Excel.Version} (PID {fixture.ProcessId}); "
                + $"DataBodyRange.RowHeight = {(reported is null ? "null" : reported.GetType().Name + " '" + reported + "'")}");

            PanelGridOutcome mixed = new ExcelPanelGridMeasurement(fixture.Excel)
                .Measure([GanttTableSchema.Default.Columns[0].Name]);

            Assert.False(mixed.Succeeded, "A mixed-height body has no single row height to report.");
            Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, mixed.Refusal);
        }
        finally
        {
            // Release the proxies the body took before the fixture quits Excel.
            scope.Dispose();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
