using System.Diagnostics;
using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live measurement of what makes the first render slow, and of the one in-process
/// mechanism that is still available after <c>Application.Calculation</c> was proven
/// unusable (<see cref="ApplicationStateScopeIntegrationTests"/>).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this exists.</b> R5.18n measured 6626.6 ms to write 52 shapes on the first
/// render after Initialise, and 80.5 ms for the same 52 shapes after an
/// out-of-process <c>Application.Calculation = manual</c>. That setter cannot be used
/// in process, so the question this test answers with numbers is whether the
/// per-sheet equivalent — <c>Worksheet.EnableCalculation = false</c> — produces the
/// same effect, and whether the cost is a first-render artefact or per-shape work.
/// </para>
/// <para>
/// <b>No timing assertion.</b> A wall-clock threshold would fail on a loaded machine
/// and teach nothing; the assertions cover the shape count and the property round
/// trip, and the timings are printed as evidence. The decision they inform is a
/// product decision, recorded in the ADR, not a test.
/// </para>
/// </remarks>
public class ShapeWriteCostIntegrationTests(ITestOutputHelper output)
{
    /// <summary>
    /// A legal owned-shape identifier: <c>CouldBeOwnedName</c> requires the
    /// <c>owner:role</c> form, and the first run of this probe wrote zero shapes because
    /// <c>probe.0</c> was refused as <c>UnsupportedIdentifierForm</c>.
    /// </summary>
    private const string ProbePrefix = "probe:";

    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_first_render_cost_is_measured_with_and_without_worksheet_calculation()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Excel.Workbook workbook = fixture.CreateWorkbook();

            // The real initialised sheet, not a blank one: the cost only reproduces on a
            // sheet that carries the table, its style, the defined name and the
            // validation.
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            _output.WriteLine(
                $"initialise path={initialised.Path} sheet={initialised.SheetName} refusal={initialised.Refusal}");
            Assert.True(initialised.Succeeded);

            using OfficeFixture.ComScope scope = new();
            Excel.Worksheet sheet = scope.Track((Excel.Worksheet)workbook.ActiveSheet);
            sheet.Activate();

            Excel.Range used = scope.Track(sheet.UsedRange);
            _output.WriteLine(
                $"sheet='{sheet.Name}' used={used.Address[false, false]} shapes={sheet.Shapes.Count} "
                + $"excel={fixture.Excel.Version}");

            // Initialise leaves the target unprotected at this point in its own sequence,
            // so no protection window is needed here; it is recorded because a future
            // initialiser change that protects the sheet would otherwise turn this probe
            // into a silent zero-shape measurement.
            _output.WriteLine($"ProtectContents before = {sheet.ProtectContents}");

            var writer = new ExcelShapeWriter(fixture.Excel);

            // Phase 1: the state a real refresh runs in, on a sheet Excel has just built.
            double firstMs = Measure(writer, sheet, 52, "1");
            Delete(sheet);
            Report("1 first render, calculation on", 52, firstMs);

            // Phase 2: the same work with the worksheet's own calculation switched off.
            bool calculationBefore = sheet.EnableCalculation;
            _output.WriteLine($"EnableCalculation before = {calculationBefore}");
            sheet.EnableCalculation = false;
            bool calculationDuring = sheet.EnableCalculation;
            double suppressedMs = Measure(writer, sheet, 52, "2");
            sheet.EnableCalculation = calculationBefore;
            bool calculationAfter = sheet.EnableCalculation;
            Delete(sheet);
            Report("2 same render, EnableCalculation off", 52, suppressedMs);

            // Phase 3: repeat of phase 1, so a first-run artefact cannot be mistaken for
            // the cause. If 3 is as slow as 1, the cost is per-shape; if it is fast, the
            // cost was a one-off rebuild.
            double repeatMs = Measure(writer, sheet, 52, "3");
            Delete(sheet);
            Report("3 repeat, calculation on", 52, repeatMs);

            // Phase 4: the same 52 shapes carrying the styling and text a real scene has.
            // A plain rectangle is the cheapest possible write; the live chart's
            // primitives are labelled text boxes, dashed lines, and filled bands, and the
            // 127 ms/shape the live trace reported has to be explained by something.
            double styledMs = MeasureStyled(writer, sheet, 52);
            Report("4 styled + text, calculation on", 52, styledMs);

            // Phase 5: the single z-order pass, measured on the shapes phase 4 left.
            var ids = new List<string>(52);
            for (var index = 0; index < 52; index++)
            {
                ids.Add(ProbePrefix + index.ToString(CultureInfo.InvariantCulture));
            }

            var zOrderStopwatch = Stopwatch.StartNew();
            ShapeWriteOutcome zOrder = writer.ApplyZOrder(ids);
            zOrderStopwatch.Stop();
            _output.WriteLine($"5 ApplyZOrder({ids.Count}): {zOrderStopwatch.Elapsed.TotalMilliseconds:N1} ms refusal={zOrder.Refusal}");
            Assert.True(zOrder.Succeeded);
            Delete(sheet);

            _output.WriteLine($"EnableCalculation round trip: {calculationBefore} -> {calculationDuring} -> {calculationAfter}");

            // The property must be usable and restorable, or it cannot be a fix: this is
            // the positive half of the probe (the negative half is the Calculation probe).
            Assert.True(calculationBefore);
            Assert.False(calculationDuring);
            Assert.True(calculationAfter);
            Assert.Equal(0, CountProbeShapes(sheet));
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private void Report(string phase, int shapeCount, double elapsedMs) =>
        _output.WriteLine(string.Create(
            CultureInfo.InvariantCulture,
            $"{phase}: {elapsedMs:N1} ms for {shapeCount} shapes = {elapsedMs / shapeCount:N2} ms/shape"));

    private double Measure(ExcelShapeWriter writer, Excel.Worksheet sheet, int shapeCount, string phase)
    {
        var stopwatch = new Stopwatch();
        // A batch is what the live renderer opens (ADR-0054), so the measurement carries
        // the same shape-independent setup cost the real write pass has.
        using (writer.BeginBatch())
        {
            stopwatch.Start();
            for (var index = 0; index < shapeCount; index++)
            {
                var bounds = new RectD(
                    20d + (index % 12 * 70d),
                    40d + (index / 12 * 14d),
                    60d,
                    12d);
                ShapeWriteOutcome outcome = writer.Create(new OfficeShapeRequest(
                    ProbePrefix + index.ToString(CultureInfo.InvariantCulture),
                    OfficeShapeKind.Rectangle,
                    new OfficeShapeGeometry(Bounds: bounds),
                    ZLayer.Background));
                if (index == 0)
                {
                    _output.WriteLine($"phase {phase} first create: {outcome}");
                }
            }

            stopwatch.Stop();
        }
        Assert.Equal(shapeCount, CountProbeShapes(sheet));
        return stopwatch.Elapsed.TotalMilliseconds;
    }

    private static double MeasureStyled(ExcelShapeWriter writer, Excel.Worksheet sheet, int shapeCount)
    {
        var stopwatch = Stopwatch.StartNew();
        using (writer.BeginBatch())
        {
            for (var index = 0; index < shapeCount; index++)
            {
                var box = new RectD(
                    20d + (index % 12 * 70d),
                    40d + (index / 12 * 14d),
                    60d,
                    12d);
                string id = ProbePrefix + index.ToString(CultureInfo.InvariantCulture);
                OfficeShapeRequest request = (index % 3) switch
                {
                    0 => new OfficeShapeRequest(
                        id,
                        OfficeShapeKind.Rectangle,
                        new OfficeShapeGeometry(Bounds: box),
                        ZLayer.ActivityBody,
                        FillColour: GanttTableAppearance.HeaderFill,
                        StrokeColour: GanttTableAppearance.HeaderTextColour,
                        LineWidthPt: 1d),
                    1 => new OfficeShapeRequest(
                        id,
                        OfficeShapeKind.Line,
                        new OfficeShapeGeometry(
                            From: new PointD(box.Left, box.Top),
                            To: new PointD(box.Right, box.Top)),
                        ZLayer.Grid,
                        StrokeColour: GanttTableAppearance.HeaderTextColour,
                        LineWidthPt: 0.75d,
                        StrokeDash: GanttStrokeDash.Dash),
                    _ => new OfficeShapeRequest(
                        id,
                        OfficeShapeKind.TextBox,
                        new OfficeShapeGeometry(Bounds: box),
                        ZLayer.Label,
                        FontFamily: "Calibri",
                        FontSizePt: 8d,
                        Bold: false,
                        Text: "Jan 2026",
                        Alignment: GanttTextAlignment.Centre,
                        TextColour: GanttTableAppearance.HeaderTextColour),
                };

                ShapeWriteOutcome outcome = writer.Create(request);
                Assert.True(outcome.Succeeded, $"styled create {id} refused: {outcome.Refusal}");
            }
        }

        stopwatch.Stop();
        return stopwatch.Elapsed.TotalMilliseconds;
    }

    private static int CountProbeShapes(Excel.Worksheet sheet)
    {
        var count = 0;
        for (var index = 1; index <= sheet.Shapes.Count; index++)
        {
            Excel.Shape shape = sheet.Shapes.Item(index);
            if (shape.Name.StartsWith(ProbePrefix, StringComparison.Ordinal))
            {
                count++;
            }
        }

        return count;
    }

    private static void Delete(Excel.Worksheet sheet)
    {
        for (var index = sheet.Shapes.Count; index >= 1; index--)
        {
            Excel.Shape shape = sheet.Shapes.Item(index);
            if (shape.Name.StartsWith(ProbePrefix, StringComparison.Ordinal))
            {
                shape.Delete();
            }
        }
    }
}
