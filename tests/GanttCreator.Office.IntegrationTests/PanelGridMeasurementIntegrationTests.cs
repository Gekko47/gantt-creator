using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the panel-grid measurement port, and specifically the
/// <c>Object</c>-typed conversions and the per-row read. Tagged
/// <c>[Trait("Category","OfficeIntegration")]</c> so <c>verify-quick.ps1</c> and
/// <c>verify.ps1</c> exclude it; run via <c>pwsh ./scripts/verify-office.ps1</c>.
/// </summary>
/// <remarks>
/// <para>
/// The contract tests substitute a mocked <c>Range</c>, which proves the adapter's
/// rule but cannot prove what the host actually reports. Only a live table settles
/// that.
/// </para>
/// <para>
/// <b>The table needs data rows before it can be measured at all.</b> This is the
/// first thing the live run found, and the contract tests could not have: the
/// initialiser creates <c>tblGanttData</c> from a single header row, so a freshly
/// initialised table has <c>ListRows.Count == 0</c> and <c>DataBodyRange</c> is
/// <see langword="null"/>. Measuring it is correctly refused as
/// <see cref="PanelGridRefusalReason.InvalidMeasurement"/>, because there is no body
/// whose bounds could be reproduced. Every test here therefore adds body rows first,
/// and the refusal is asserted separately as its own case rather than being the
/// accidental result of measuring an empty table.
/// </para>
/// <para>
/// Read-only with respect to the adapter: these tests write row heights and add rows
/// to set up the condition, but the adapter under test never mutates the workbook.
/// </para>
/// </remarks>
public class PanelGridMeasurementIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    private static Excel.ListObject ResolveGanttTable(OfficeFixture.ComScope scope, Excel.Workbook workbook)
    {
        Excel.Sheets sheets = scope.Track(workbook.Sheets);
        Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(sheets[GanttWorkbookContract.GanttSheetLabel]);
        Excel.ListObjects objects = scope.Track(sheet.ListObjects);
        return scope.Track(objects[GanttTableSchema.TableName]);
    }

    /// <summary>Adds data rows, so the table has a body to measure.</summary>
    private static void AddBodyRows(OfficeFixture.ComScope scope, Excel.ListObject table, int count)
    {
        Excel.ListRows rows = scope.Track(table.ListRows);
        for (var i = 0; i < count; i++)
        {
            scope.Track(rows.Add());
        }
    }

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

            // The precondition, asserted in this same session rather than a separate
            // one: a freshly initialised table has a header row and no data rows, so
            // DataBodyRange is null and there is no body to reproduce. Refusing is
            // correct; what was wrong before was relying on it as the accidental
            // result of a test that meant to measure a populated table.
            Excel.ListObject table = scope.Track(ResolveGanttTable(scope, workbook));
            _output.WriteLine($"fresh table ListRows.Count={scope.Track(table.ListRows).Count}");
            PanelGridOutcome empty = new ExcelPanelGridMeasurement(fixture.Excel)
                .Measure([GanttTableSchema.Default.Columns[0].Name]);
            _output.WriteLine($"empty table: success={empty.Succeeded} reason={empty.Refusal}");
            Assert.False(empty.Succeeded, "A table with no data rows has no panel body to measure.");
            Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, empty.Refusal);

            // Now add rows so there is a body.
            AddBodyRows(scope, table, 3);

            // Uniform heights first: the measurement must SUCCEED, or the per-row
            // assertions below would prove nothing about mixed heights.
            PanelGridOutcome uniform = new ExcelPanelGridMeasurement(fixture.Excel)
                .Measure([GanttTableSchema.Default.Columns[0].Name]);
            _output.WriteLine("uniform heights: success=" + uniform.Succeeded
                + " rows=" + (uniform.Grid is null ? "-" : string.Join("/", uniform.Grid.RowHeightsPt))
                + " header=" + (uniform.Grid is null ? "-" : uniform.Grid.HeaderHeightPt.ToString(CultureInfo.InvariantCulture)));
            Assert.True(uniform.Succeeded, $"Uniform-height measurement refused: {uniform.Refusal}");
            Assert.NotNull(uniform.Grid);

            // Make the body genuinely mixed. The expected list is derived from the
            // row count the host actually reports rather than from how many rows were
            // requested: the initialised table's range means the two differ, and a
            // hardcoded count here would be a fixture bug disguised as a measurement
            // failure.
            Excel.Range rows = scope.Track(table.DataBodyRange!.Rows);
            double[] expected = [.. Enumerable.Range(0, rows.Count).Select(i => 15d + (i * 10d))];
            for (var i = 1; i <= rows.Count; i++)
            {
                scope.Track(rows.Item[i]).RowHeight = expected[i - 1];
            }

            // Read the host's own answer back rather than assuming it. The aggregate
            // over a mixed range is documented to report either the first row's height
            // or Null; the adapter must not depend on which.
            object? reported = table.DataBodyRange!.RowHeight;
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
            for (var i = 0; i < expected.Length; i++)
            {
                Assert.Equal(expected[i], mixed.Grid.RowHeightsPt[i], precision: 6);
            }

            // The protected-target claim shares this session deliberately. It is a
            // genuine second claim - worksheet protection blocks writes, not reads, so
            // a read-only measurement must still succeed - but giving it its own Excel
            // instance adds a whole live session, and the COM-proxy leak ratchet counts
            // sessions. Merging keeps the evidence and the suite's cost honest.
            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                scope.Track(workbook.Sheets)[GanttWorkbookContract.GanttSheetLabel]);
            sheet.Protect("gantt", true, true, true);
            try
            {
                PanelGridOutcome outcome = new ExcelPanelGridMeasurement(fixture.Excel)
                    .Measure([GanttTableSchema.Default.Columns[0].Name]);

                _output.WriteLine("protected-sheet measurement: success=" + outcome.Succeeded
                    + " reason=" + outcome.Refusal);

                // If a future Office build refuses these reads, this fails and the
                // policy needs revisiting with real host evidence rather than by
                // assumption.
                Assert.True(outcome.Succeeded, $"Protected-sheet measurement refused: {outcome.Refusal}");
                Assert.NotNull(outcome.Grid);
                Assert.Equal(expected.Length, outcome.Grid.RowHeightsPt.Count);
            }
            finally
            {
                // A sheet left protected is a locked COM object at teardown, which the
                // COM-proxy leak ratchet would then count against this suite.
                sheet.Unprotect("gantt");
            }
        }
        finally
        {
            // Release the proxies the body took before the fixture quits Excel.
            scope.Dispose();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
