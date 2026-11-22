using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R7.4 Office gate (work item T-3.1): the pasted production group is fitted and placed
/// <b>once</b> at 4:3, 16:9 and one recorded custom size, through the production paste and
/// placement stages, with the placement, the policy margin, the aspect, the label state and the
/// editability inspected on the host.
/// </summary>
/// <remarks>
/// <para>
/// <b>One staging pass, three cycles (audit A-6).</b> The R7.3 staging → copy path runs once; each
/// size then pastes, fits, applies, verifies and deletes on the fixture-owned deck. The slide size
/// is changed on that deck only — a user deck is never touched. Editability is checked once, on
/// the final size, because ungrouping destroys the group the by-id cleanup deletes and the
/// fixture discards its deck at teardown regardless.
/// </para>
/// <para>
/// <b>Label fidelity is the measured rule (D5, 2026-11-22).</b> The host does not scale text with
/// the group, and the staged composition already overflows four of 81 text children. The gate
/// therefore measures the composition's own label state <em>before</em> any placement and asserts
/// that the placement adds no overflowing child; the absolute counts are recorded beside it.
/// </para>
/// <para>
/// <b>The margin is the recorded D-R7.4 policy.</b> The inset is the chart's own outer padding
/// (<see cref="ChartPaddingPt.Uniform"/> of the <c>ChartOuterPaddingPt</c> token) at scene scale,
/// so the expected slide margin is <c>inset × scale</c> on the binding axis. M-5's custom size is
/// the extreme 1200 × 400.
/// </para>
/// </remarks>
public sealed class PowerPointSlideFitIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private const double _tolerancePt = HostOwnedGeometry.GeometryMatchTolerancePt;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_pasted_group_is_fitted_and_placed_once_on_three_slide_sizes()
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
            StagingLifecycleOutcome lifecycle = StagingLifecycle.Run(new(
                state,
                stagingWriter,
                new ExcelShapeGrouper(excelFixture.Excel),
                resolution,
                Copy: group => StagedShapeCopier.Copy(
                    group,
                    new ExcelShapeCopier(excelFixture.Excel),
                    new WindowsClipboardFormatReader())));

            Assert.True(lifecycle.Succeeded, $"Lifecycle refused: {lifecycle.Refusal}");
            Assert.NotNull(lifecycle.Copy);
            Assert.True(
                lifecycle.Copy.Succeeded,
                $"Copy refused: {lifecycle.Copy.Refusal} {lifecycle.Copy.PortRefusal} {lifecycle.Copy.Detail}");

            StagedGroupOutcome group = Assert.IsType<StagedGroupOutcome>(lifecycle.Group);
            Assert.True(group.Succeeded);

            // D6: the aspect source is the resolved zero-origin ExportBounds.Bounds. PlacementDelta is
            // already consumed by the staging composition and is never re-applied here.
            RectD bounds = Assert.IsType<RectD>(group.Bounds);
            int stagedChildren = group.ChildOrder.Count;

            // D-R7.4: the inset is the chart's own outer padding at scene scale.
            double insetPt = ChartPaddingPt.Uniform(
                GanttCatalogues.MetricDefault("ChartOuterPaddingPt")).MaxPt;
            Assert.True(insetPt > 0, "The recorded inset must be positive.");

            // 4:3, 16:9 and M-5's recorded custom size (an extreme wide aspect).
            (double WidthPt, double HeightPt)[] slideSizes = [(720, 540), (960, 540), (1200, 400)];

            PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());
            try
            {
                // The adapter must attach before any stage reads the host: an unattached adapter
                // reports a refused observation, which the paste stage classifies as NoSlide
                // (measured 2026-11-22 on the first run of this gate).
                PowerPointApplicationFacts attached = adapter.AttachOrStart();
                Assert.True(
                    attached.Available,
                    $"The adapter could not attach to PowerPoint: hresult=0x{attached.HResult:X8}");

                PowerPointPasteStage pasteStage = new(adapter);
                PowerPointPlacementStage placementStage = new(adapter);

                for (int cycle = 0; cycle < slideSizes.Length; cycle++)
                {
                    (double slideWidthPt, double slideHeightPt) = slideSizes[cycle];
                    powerPointFixture.SetSlideSize(slideWidthPt, slideHeightPt);
                    (double readWidthPt, double readHeightPt) = powerPointFixture.ReadSlideSize();
                    Assert.True(
                        Math.Abs(readWidthPt - slideWidthPt) <= 0.5,
                        $"The host read back slide width {readWidthPt} for a requested {slideWidthPt}.");
                    Assert.True(
                        Math.Abs(readHeightPt - slideHeightPt) <= 0.5,
                        $"The host read back slide height {readHeightPt} for a requested {slideHeightPt}.");

                    PowerPointPasteStageResult paste = pasteStage.Paste(
                        new PowerPointPasteExpectation(stagedChildren, bounds.Width, bounds.Height));
                    Assert.True(
                        paste.Succeeded,
                        $"Paste refused: {paste.Refusal} {paste.Validation?.Refusal}");
                    int groupId = paste.GroupId
                        ?? throw new InvalidOperationException("The paste stage reported no group identity.");

                    // The composition's own label state, BEFORE any placement.
                    IReadOnlyList<PowerPointChildTextMetrics> natural = powerPointFixture.MeasureGroupTextById(groupId);
                    string[] naturalOverflow = OverflowingNames(natural);

                    PowerPointPlacementResult placed = placementStage.Place(groupId, bounds, insetPt);
                    Assert.True(
                        placed.Succeeded,
                        $"Placement refused: {placed.Refusal} fit={placed.FitRefusal} hresult=0x{placed.HResult:X8}");
                    SlidePlacement placement = Assert.IsType<SlidePlacement>(placed.Placement);
                    PowerPointPlacementDeltas deltas = Assert.IsType<PowerPointPlacementDeltas>(placed.Deltas);

                    // The read-back is the acceptance guard: the host accepted the fitted numbers.
                    Assert.True(
                        deltas.MaxPt <= _tolerancePt,
                        $"Read-back drifted by {deltas.MaxPt} pt.");

                    // The policy margin: inset x scale on the binding axis, never smaller elsewhere.
                    double marginPt = insetPt * placement.Scale;
                    double horizontalMarginPt = (readWidthPt - placement.Width) / 2;
                    double verticalMarginPt = (readHeightPt - placement.Height) / 2;
                    Assert.True(horizontalMarginPt >= marginPt - _tolerancePt);
                    Assert.True(verticalMarginPt >= marginPt - _tolerancePt);
                    Assert.True(
                        Math.Abs(horizontalMarginPt - marginPt) <= _tolerancePt
                            || Math.Abs(verticalMarginPt - marginPt) <= _tolerancePt,
                        $"Neither axis carries the policy margin {marginPt}; horizontal {horizontalMarginPt}, vertical {verticalMarginPt}.");

                    // Containment, and the aspect the export bounds define (D2/D6).
                    Assert.True(placement.Left >= -_tolerancePt);
                    Assert.True(placement.Top >= -_tolerancePt);
                    Assert.True(placement.Left + placement.Width <= readWidthPt + _tolerancePt);
                    Assert.True(placement.Top + placement.Height <= readHeightPt + _tolerancePt);
                    double expectedAspect = bounds.Width / bounds.Height;
                    double appliedAspect = placement.Width / placement.Height;
                    Assert.True(
                        Math.Abs(appliedAspect - expectedAspect) <= _tolerancePt,
                        $"Aspect drifted: expected {expectedAspect}, applied {appliedAspect}.");

                    // D5 measured refinement: placement may remove label overflow, never add it.
                    IReadOnlyList<PowerPointChildTextMetrics> placedText =
                        powerPointFixture.MeasureGroupTextById(groupId);
                    string[] placedOverflow = OverflowingNames(placedText);
                    string[] addedOverflow = [.. placedOverflow.Except(naturalOverflow, StringComparer.Ordinal)];
                    Assert.Empty(addedOverflow);

                    _output.WriteLine(string.Create(
                        CultureInfo.InvariantCulture,
                        $"host={powerPointFixture.BuildDescription} slide={readWidthPt:0.#}x{readHeightPt:0.#}"
                            + $" bounds={bounds.Width:0.##}x{bounds.Height:0.##} children={stagedChildren}"
                            + $" inset={insetPt:0.##} scale={placement.Scale:0.#####}"
                            + $" placed=({placement.Left:0.##},{placement.Top:0.##},{placement.Width:0.##},{placement.Height:0.##})"
                            + $" deltas=({deltas.LeftPt:0.###},{deltas.TopPt:0.###},{deltas.WidthPt:0.###},{deltas.HeightPt:0.###})"
                            + $" margin={marginPt:0.###} appliedMargin=({horizontalMarginPt:0.###},{verticalMarginPt:0.###})"
                            + $" aspectDelta={Math.Abs(appliedAspect - expectedAspect):0.#####}"
                            + $" naturalOverflow=[{string.Join("; ", naturalOverflow)}]"
                            + $" placedOverflow=[{string.Join("; ", placedOverflow)}]"));

                    // The once rule at the live seam: the consumed placement cannot be re-applied.
                    PowerPointPlacementResult second = placementStage.Apply(groupId, placement);
                    Assert.False(second.Succeeded);
                    Assert.Equal(PowerPointPlacementRefusal.AlreadyApplied, second.Refusal);

                    if (cycle == slideSizes.Length - 1)
                    {
                        // Editability, once, on the final size: the ungroup yields the staged
                        // children as editable shapes. The fixture tracks them for the delete.
                        int editable = powerPointFixture.UngroupShapeById(groupId);
                        Assert.Equal(stagedChildren, editable);
                        powerPointFixture.DeletePastedShapes();
                    }
                    else
                    {
                        _ = powerPointFixture.DeleteShapeById(groupId);
                    }
                }
            }
            finally
            {
                adapter.Dispose();
            }

            // Zero staging residue and the live chart intact.
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

    private static string[] OverflowingNames(IReadOnlyList<PowerPointChildTextMetrics> text) =>
        [.. text.Where(child => child.HasText && !Fits(child)).Select(child => child.Name)];

    private static bool Fits(PowerPointChildTextMetrics child) =>
        child.TextWidthPt <= child.ShapeWidthPt + _tolerancePt
        && child.TextHeightPt <= child.ShapeHeightPt + _tolerancePt;
}
