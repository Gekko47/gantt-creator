using System.Globalization;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// The row heights a workbook has the moment Initialise has finished and its first
/// automatic render has normalised them — with NO body rows, which is the state the user
/// sees before they type anything.
/// </summary>
/// <remarks>
/// <para>
/// Reported from the host on 2026-11-14: "row 5 and 6 are not the correct height on
/// initialise". The layout rows are positional — the top padding row, the reserved row that
/// carries the title and the year band, the header row that IS the period band
/// (ADR-0030 D4), then the anchor and bottom padding rows below the body — so with an empty
/// body every one of those is decided by an offset from a body that does not exist. This
/// prints the whole column of heights before and after the normaliser rather than asserting
/// a layout nobody has yet written down for the zero-row case.
/// </para>
/// <para>
/// No assertion about the heights is made yet: the measurement is the point. Once the
/// expected zero-body layout is agreed, these printouts become the assertions.
/// </para>
/// </remarks>
public class InitialiseRowHeightIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_zero_body_layout_rows_have_the_catalogue_heights_after_normalise()
    {
        double bodyToken = GanttCatalogues.MetricDefault("GanttRowHeightPt");
        double splitterToken = GanttCatalogues.MetricDefault("SplitterHeightPt");
        double spacerToken = GanttCatalogues.MetricDefault("SpacerHeightPt");
        double periodBandToken = GanttCatalogues.MetricDefault("PeriodBandHeightPt");
        double yearBandToken = GanttCatalogues.MetricDefault("YearBandHeightPt");
        double paddingToken = GanttCatalogues.MetricDefault("ChartPaddingRowHeightPt");
        double anchorToken = GanttCatalogues.MetricDefault("ChartAnchorRowHeightPt");

        _output.WriteLine(string.Create(
            CultureInfo.InvariantCulture,
            $"tokens: padding={paddingToken} yearBand={yearBandToken} periodBand={periodBandToken}"
            + $" body={bodyToken} anchor={anchorToken}"));

        OfficeFixture fixture = new();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Excel.Workbook workbook = fixture.CreateWorkbook();

            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"initialise path={initialised.Path} sheet={initialised.SheetName} refusal={initialised.Refusal}"));
            Assert.True(initialised.Succeeded);

            using OfficeFixture.ComScope scope = new();
            Excel.Worksheet sheet = scope.Track((Excel.Worksheet)workbook.ActiveSheet);
            Excel.ListObjects objects = scope.Track(sheet.ListObjects);
            Excel.ListObject table = scope.Track(objects[GanttTableSchema.TableName]);
            Excel.ListRows rows = scope.Track(table.ListRows);
            _output.WriteLine("body rows = " + rows.Count.ToString(CultureInfo.InvariantCulture));

            DumpRows(scope, sheet, "initialise");

            // The same call the refresh makes, with the same catalogue defaults, and no type
            // texts because there are no body rows to classify.
            ExcelRowHeightNormaliser normaliser = new(fixture.Excel);
            RowHeightNormalisationOutcome normalised = normaliser.Normalise(
                bodyToken,
                splitterToken,
                spacerToken,
                periodBandToken,
                yearBandToken,
                paddingToken,
                anchorToken);
            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"normalise: succeeded={normalised.Succeeded} rowsWritten={normalised.RowsWritten} refusal={normalised.Refusal}"));

            DumpRows(scope, sheet, "after normalise");

            PanelGridOutcome grid = new ExcelPanelGridMeasurement(fixture.Excel).Measure(["Type"]);
            _output.WriteLine("grid: success=" + grid.Succeeded + " reason=" + grid.Refusal);
            if (grid.Grid is { } measured)
            {
                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"grid: originTop={measured.OriginTopPt:F2} header={measured.HeaderHeightPt:F2}"
                    + $" topPadding={measured.TopPaddingHeightPt:F2}"
                    + $" bodyRows={measured.RowHeightsPt.Count}"));
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private void DumpRows(OfficeFixture.ComScope scope, Excel.Worksheet sheet, string label)
    {
        for (int row = 1; row <= 8; row++)
        {
            Excel.Range line = scope.Track(sheet.Rows[row]);
            double top = Convert.ToDouble(line.Top, CultureInfo.InvariantCulture);
            double height = Convert.ToDouble(line.RowHeight, CultureInfo.InvariantCulture);
            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"{label} row {row}: top={top:F2} height={height:F2}"));
        }
    }
}
