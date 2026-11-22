using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R7.4 Phase 0 (work item T-0.2): measures the three Step-0 resize routes on the reference
/// host with the representative staged chart, and records the evidence table the D-R7.4
/// margin/route decisions consume. This body deliberately asserts only the preconditions and
/// records the measurements — the route is chosen by the human from this evidence, not by a
/// passing assertion.
/// </summary>
/// <remarks>
/// <para>
/// <b>Evidence, not a verdict.</b> For each slide size (4:3, 16:9, a custom wide size) and each
/// route — (a) absolute four-member box write, (b) <c>LockAspectRatio</c> + width, (c) uniform
/// <c>ScaleWidth</c>/<c>ScaleHeight</c> — the body records the pasted group box, the fitted
/// target, the applied box, the aspect delta and the text-child bound dimensions. A divergent
/// form is evidence for the decision, never a reason to widen a validator.
/// </para>
/// <para>
/// <b>Harness shape (audit A-4).</b> The probe needs the production-scale staged Excel group on
/// the clipboard, so it reuses the R7.3 staging → copy path and the fixture's typed
/// <c>PasteSpecialShape</c>. The Excel host owns the staging lifecycle; the PowerPoint host is
/// fixture-started and quit only by the fixture.
/// </para>
/// </remarks>
public sealed class PowerPointSlideFitRouteProbeIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_three_resize_routes_are_measured_on_the_reference_host()
    {
        OfficeFixture excelFixture = new();
        PowerPointFixture powerPointFixture = new();

        try
        {
            await excelFixture.InitializeAsync().ConfigureAwait(true);

            // PowerPoint starts BEFORE the copy so launching it cannot disturb the clipboard
            // the lifecycle is about to fill (the R6.7/R7.3 ordering).
            powerPointFixture.Initialize(withWindow: true, slideCount: 1);

            using OfficeFixture.ComScope scope = new();
            (Excel.Worksheet sheet, Excel.Sheets sheets, ExportBoundsOutcome resolution, Excel.Shape sentinel, Excel.Range witness, object? witnessBefore, int sheetCountBefore) =
                StagingLifecycleIntegrationTests.Prepare(excelFixture, scope);

            ExcelShapeWriter stagingWriter = new(
                excelFixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);
            ExcelShapeWriter liveWriter = new(excelFixture.Excel);
            Assert.True(liveWriter.Create(StagingLifecycleIntegrationTests.LiveProbe("live:probe-1")).Succeeded);

            (int scrollRow, int scrollColumn, double zoom, string selection) =
                StagingLifecycleIntegrationTests.CaptureView(excelFixture, scope);

            using ExcelApplicationStateScope state = new(excelFixture.Excel);
            StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
                state,
                stagingWriter,
                new ExcelShapeGrouper(excelFixture.Excel),
                resolution,
                Copy: group => StagedShapeCopier.Copy(
                    group,
                    new ExcelShapeCopier(excelFixture.Excel),
                    new WindowsClipboardFormatReader())));

            Assert.True(outcome.Succeeded, $"Lifecycle refused: {outcome.Refusal}");
            Assert.NotNull(outcome.Copy);
            Assert.True(
                outcome.Copy.Succeeded,
                $"Copy refused: {outcome.Copy.Refusal} {outcome.Copy.PortRefusal} {outcome.Copy.Detail}");

            // 4:3, 16:9 and one custom wide size. The inset is the RECORDED D-R7.4 policy value
            // (ChartPaddingPt.Uniform of the chart's outer padding, 4 pt at scene scale), so the
            // measured target is the shipped placement and the routes stay like-for-like.
            (double WidthPt, double HeightPt)[] slideSizes = [(720, 540), (960, 540), (1200, 400)];
            string[] routes = ["a", "b", "c"];
            double insetPt = ChartPaddingPt.Uniform(
                GanttCatalogues.MetricDefault("ChartOuterPaddingPt")).MaxPt;
            bool sawTextChild = false;

            foreach ((double slideWidthPt, double slideHeightPt) in slideSizes)
            {
                powerPointFixture.SetSlideSize(slideWidthPt, slideHeightPt);
                (double readWidthPt, double readHeightPt) = powerPointFixture.ReadSlideSize();
                Assert.True(
                    Math.Abs(readWidthPt - slideWidthPt) <= 0.5,
                    $"The host read back slide width {readWidthPt} for a requested {slideWidthPt}.");
                Assert.True(
                    Math.Abs(readHeightPt - slideHeightPt) <= 0.5,
                    $"The host read back slide height {readHeightPt} for a requested {slideHeightPt}.");

                foreach (string route in routes)
                {
                    PowerPointPasteOutcome pasted = powerPointFixture.PasteSpecialShape();
                    try
                    {
                        PowerPointGroupBox groupBox = powerPointFixture.LastPastedGroupBox();
                        Assert.True(
                            groupBox.ChildCount > 0,
                            "The pasted group exposes no children, so no route can be measured.");

                        // The NATURAL (unresized) label state is measured first. It separates a
                        // label the composition already overflows from one the resize breaks —
                        // the distinction work item D5's stop condition turns on.
                        IReadOnlyList<PowerPointChildTextMetrics> naturalText =
                            powerPointFixture.MeasurePastedGroupText();
                        int naturalChildren = naturalText.Count(child => child.HasText);
                        int naturalFits = naturalText.Count(child => !child.HasText || Fits(child));
                        string naturalFonts = Fonts(naturalText);
                        string[] naturalOverflow = DescribeOverflow(naturalText);

                        SlideFitOutcome fit = SlideFit.Fit(
                            new RectD(0, 0, groupBox.WidthPt, groupBox.HeightPt),
                            readWidthPt,
                            readHeightPt,
                            insetPt);
                        Assert.True(fit.Succeeded, $"Fit refused: {fit.Refusal}");
                        SlidePlacement placement = Assert.IsType<SlidePlacement>(fit.Placement);

                        PowerPointGroupBox applied = route switch
                        {
                            "a" => powerPointFixture.ApplyRouteAbsoluteBox(
                                placement.Left,
                                placement.Top,
                                placement.Width,
                                placement.Height),
                            "b" => powerPointFixture.ApplyRouteLockAspectWidth(placement.Width),
                            "c" => powerPointFixture.ApplyRouteUniformScale((float)placement.Scale),
                            _ => throw new InvalidOperationException("Unknown probe route."),
                        };

                        IReadOnlyList<PowerPointChildTextMetrics> text =
                            powerPointFixture.MeasurePastedGroupText();
                        sawTextChild |= text.Any(child => child.HasText);
                        int textChildren = text.Count(child => child.HasText);
                        int textFitsBox = text.Count(child => !child.HasText || Fits(child));
                        string postFonts = Fonts(text);
                        string[] overflow = DescribeOverflow(text);

                        double appliedAspect = applied.WidthPt / applied.HeightPt;
                        double targetAspect = placement.Width / placement.Height;
                        _output.WriteLine(string.Create(
                            CultureInfo.InvariantCulture,
                            $"host={powerPointFixture.BuildDescription} slide={slideWidthPt}x{slideHeightPt}"
                                + $" route={route} pastedShapes={pasted.ShapeCount} children={groupBox.ChildCount}"
                                + $" pastedBox={groupBox.WidthPt:0.##}x{groupBox.HeightPt:0.##}"
                                + $" target=({placement.Left:0.##},{placement.Top:0.##},{placement.Width:0.##},{placement.Height:0.##})"
                                + $" scale={placement.Scale:0.#####}"
                                + $" applied=({applied.LeftPt:0.##},{applied.TopPt:0.##},{applied.WidthPt:0.##},{applied.HeightPt:0.##})"
                                + $" aspectDelta={Math.Abs(appliedAspect - targetAspect):0.#####}"
                                + $" natural={naturalFits}/{naturalChildren} naturalFonts={naturalFonts}"
                                + $" naturalOverflow=[{string.Join("; ", naturalOverflow)}]"
                                + $" textChildren={textChildren} textFitsBox={textFitsBox} postFonts={postFonts}"
                                + $" overflow=[{string.Join("; ", overflow)}]"));
                    }
                    finally
                    {
                        powerPointFixture.DeletePastedShapes();
                    }
                }
            }

            // The label evidence must not be vacuous: the representative chart carries text.
            Assert.True(
                sawTextChild,
                "The representative chart exposed no text child, so the label evidence is empty.");

            // Excel-side hygiene, mirroring the R7.3 gate.
            Assert.Empty(stagingWriter.ListOwned());
            Assert.Equal(0, StagingLifecycleIntegrationTests.CountStaged(scope, sheet));
            Assert.Equal(["live:probe-1"], liveWriter.ListOwned());

            StagingLifecycleIntegrationTests.AssertViewRestored(
                excelFixture,
                scope,
                sheets,
                sentinel,
                witness,
                witnessBefore,
                sheetCountBefore,
                scrollRow,
                scrollColumn,
                zoom,
                selection);

            StagingLifecycleIntegrationTests.CleanupLive(liveWriter, sentinel);
        }
        finally
        {
            await powerPointFixture.DisposeAsync().ConfigureAwait(true);
            await excelFixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static bool Fits(PowerPointChildTextMetrics child) =>
        child.TextWidthPt <= child.ShapeWidthPt + HostOwnedGeometry.GeometryMatchTolerancePt
        && child.TextHeightPt <= child.ShapeHeightPt + HostOwnedGeometry.GeometryMatchTolerancePt;

    private static string Fonts(IReadOnlyList<PowerPointChildTextMetrics> text) =>
        string.Join(",", text.Where(child => child.HasText).Select(child => child.FontSizePt).Distinct());

    private static string[] DescribeOverflow(IReadOnlyList<PowerPointChildTextMetrics> text) =>
        [.. text
            .Where(child => child.HasText && !Fits(child))
            .Select(child => string.Create(
                CultureInfo.InvariantCulture,
                $"{child.Name}:shape={child.ShapeWidthPt:0.#}x{child.ShapeHeightPt:0.#},text={child.TextWidthPt:0.#}x{child.TextHeightPt:0.#},wrap={child.WordWrapOn},autosize={child.AutoSize},font={child.FontSizePt:0.##}"))];
}
