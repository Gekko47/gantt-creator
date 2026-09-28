using System.Globalization;
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
    public async Task A_body_with_mixed_row_heights_measures_each_row_rather_than_being_refused()
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

            // Uniform heights first: the measurement must SUCCEED, or the per-row
            // assertions below would prove nothing about mixed heights.
            PanelGridOutcome uniform = new ExcelPanelGridMeasurement(fixture.Excel)
                .Measure([GanttTableSchema.Default.Columns[0].Name]);
            _output.WriteLine("uniform heights: success=" + uniform.Succeeded
                + " rows=" + (uniform.Grid is null ? "-" : string.Join("/", uniform.Grid.RowHeightsPt))
                + " header=" + (uniform.Grid is null ? "-" : uniform.Grid.HeaderHeightPt.ToString(CultureInfo.InvariantCulture)));
            Assert.True(uniform.Succeeded, $"Uniform-height measurement refused: {uniform.Refusal}");
            Assert.NotNull(uniform.Grid);

            // Now make the body genuinely mixed.
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

            // Read the host's own answer back rather than assuming it. The aggregate
            // over a mixed range is documented to report either the first row's height
            // or Null; the adapter must not depend on which.
            object? reported = body.RowHeight;
            _output.WriteLine($"Excel build {fixture.Excel.Version} (PID {fixture.ProcessId}); "
                + $"DataBodyRange.RowHeight = {(reported is null ? "null" : reported.GetType().Name + " '" + reported + "'")}");

            PanelGridOutcome mixed = new ExcelPanelGridMeasurement(fixture.Excel)
                .Measure([GanttTableSchema.Default.Columns[0].Name]);

            _output.WriteLine("mixed heights: success=" + mixed.Succeeded
                + " rows=" + (mixed.Grid is null ? "-" : string.Join("/", mixed.Grid.RowHeightsPt)));

            // This used to assert a typed refusal. A mixed body is a legitimate
            // worksheet, so it must now measure - and the values are the assertion,
            // not merely the success flag.
            Assert.True(mixed.Succeeded, $"Mixed-height measurement refused: {mixed.Refusal}");
            Assert.NotNull(mixed.Grid);
            Assert.Equal(2, mixed.Grid.RowHeightsPt.Count);
            Assert.Equal(15d, mixed.Grid.RowHeightsPt[0], precision: 6);
            Assert.Equal(45d, mixed.Grid.RowHeightsPt[1], precision: 6);
        }
        finally
        {
            // Release the proxies the body took before the fixture quits Excel.
            scope.Dispose();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// A read-only measurement must succeed on a protected worksheet. This adapter
    /// writes nothing, and Excel's protection blocks writes rather than reads, so a
    /// protected target is still fully measurable.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Measurement_succeeds_on_a_protected_worksheet()
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

            Excel.Sheets sheets = scope.Track(workbook.Sheets);
            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(sheets[GanttWorkbookContract.GanttSheetLabel]);
            // Contents and objects, with no password, so the fixture teardown is not
            // left holding a locked sheet.
            sheet.Protect("gantt", true, true, true);

            PanelGridOutcome outcome = new ExcelPanelGridMeasurement(fixture.Excel)
                .Measure([GanttTableSchema.Default.Columns[0].Name]);

            _output.WriteLine("protected-sheet measurement: success=" + outcome.Succeeded
                + " reason=" + outcome.Refusal);

            // The host claim under test is the premise of the policy: protection
            // blocks writes, not reads, so the measurement must still succeed. If a
            // future Office build refuses these reads, this fails and the policy needs
            // revisiting with real host evidence rather than by assumption.
            Assert.True(outcome.Succeeded, $"Protected-sheet measurement refused: {outcome.Refusal}");
        }
        finally
        {
            scope.Dispose();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
