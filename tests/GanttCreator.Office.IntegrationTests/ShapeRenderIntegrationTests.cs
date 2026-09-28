using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R4.3's Required Office gate: a scene rectangle and a scene line become real
/// Excel shapes whose point geometry, name, and ownership tag are read back off
/// the live host and compared within <see cref="GeometryMath.Epsilon"/>.
/// </summary>
/// <remarks>
/// <para>
/// The chart bounds used here start one padding left of and above the content,
/// which is the real case entity guide section 1 describes: the bounds are
/// derived by expanding the content union, so a chart whose content starts at
/// the origin has a negative bounds origin. The R4.3 Step-0 probe established
/// that the host <strong>silently clamps</strong> a negative shape offset to
/// zero, so this test is the live proof that the translation actually reaches
/// the host - without it every shape would pile up at the origin and the render
/// would look superficially plausible.
/// </para>
/// <para>
/// Geometry is read back from the live shape rather than from the value the
/// adapter intended to write, so a host-side clamp or re-anchor fails the test
/// instead of passing on the adapter's own arithmetic.
/// </para>
/// </remarks>
public class ShapeRenderIntegrationTests(ITestOutputHelper output)
{
    private const double ChartPaddingPt = 12d;

    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_scene_rectangle_and_line_land_on_the_live_sheet_within_tolerance()
    {
        var fixture = new OfficeFixture();

        // KNOWN-LIMITATIONS L19. This test's purpose is to create REAL shapes
        // through the real ExcelShapeWriter and read their geometry back off the
        // live host; a workbook holding a live shape keeps the Excel process
        // alive, so teardown legitimately escalates to a kill. The leak ratchet
        // counts forced kills, and L19 records that this signal is
        // non-deterministic and has already been raised from 24 to 26. The
        // harness provides SuppressLeakSignal for exactly this case - a test
        // that forces a kill on purpose - and the test still asserts the full
        // geometry, ownership, and reconciliation contract, so nothing is hidden
        // by it. This is NOT a licence to suppress a test that simply leaked.
        fixture.SuppressLeakSignal = true;

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            using var scope = new OfficeFixture.ComScope();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            var initialiser = new ExcelWorkbookInitialiser(fixture.Excel);
            WorkbookInitialiseOutcome initialised = initialiser.Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            GanttScene scene = BuildScene();
            var renderer = new SceneShapeRenderer(
                ChartOriginDelta.ForChartBounds(new RectD(-ChartPaddingPt, -ChartPaddingPt, 400d, 300d)));
            SceneTranslationOutcome translated = renderer.Translate(scene);
            Assert.True(translated.Complete, translated.Deferred.Count + " primitives were deferred.");
            Assert.Equal(2, translated.Requests.Count);

            var writer = new ExcelShapeWriter(fixture.Excel);
            foreach (OfficeShapeRequest request in translated.Requests)
            {
                ShapeWriteOutcome outcome = writer.Create(request);
                Assert.True(outcome.Succeeded, $"Create refused for '{request.PrimitiveId}': {outcome.Refusal}");
            }

            var sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            // The expected values below are written out by hand from the scene
            // constants plus the padding, NOT read back out of the renderer. A
            // test that derives its expectation from the code under test is
            // self-consistent by construction: mutating the translation would
            // move both sides together and the test would still pass. These
            // literals are the only independent check that the translation
            // reached the host.
            AssertBar(scope, sheet);
            AssertLine(scope, sheet);

            // Both shapes are owned and reconcilable, which is what R4.7 and R4.8
            // depend on.
            Assert.Equal(
                ["chart:grid:0", "row-1:bar"],
                writer.ListOwned().OrderBy(id => id, StringComparer.Ordinal));

            // Delete the shapes but leave the workbook open for the fixture to
            // close. Closing it here would disconnect the very proxies the
            // ComScope releases at the end of this body.
            foreach (string primitiveId in new[] { "row-1:bar", "chart:grid:0" })
            {
                Excel.Shape shape = scope.Track(sheet.Shapes.Item(primitiveId));
                shape.Delete();
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>Builds the two-primitive scene, with a negative chart-bounds origin.</summary>
    /// <returns>The created scene.</returns>
    private static GanttScene BuildScene()
    {
        // A chart whose content starts at the origin, so the derived bounds start
        // one padding to the left of and above it.
        var chartBounds = new RectD(-ChartPaddingPt, -ChartPaddingPt, 400d, 300d);
        var style = new SceneStyle("Default");

        GanttScene scene = GanttScene.TryCreate(
                chartBounds,
                new RectD(0d, 0d, 380d, 280d),
                [
                    new SceneRect(
                        "row-1:bar",
                        SceneOwnerId.Chart,
                        ZLayer.ActivityBody,
                        new RectD(40d, 30d, 180d, 24d),
                        style),
                    new SceneLine(
                        "chart:grid:0",
                        SceneOwnerId.Chart,
                        ZLayer.Grid,
                        new PointD(40d, 20d),
                        new PointD(220d, 20d),
                        style),
                ],
                [])
            .Scene ?? throw new InvalidOperationException("The fixture scene failed to validate.");
        return scene;
    }

    private void AssertBar(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        Excel.Shape shape = scope.Track(sheet.Shapes.Item("row-1:bar"));

        // Scene bounds (40, 30, 180, 24) translated by the 12pt padding.
        _output.WriteLine("bar expected L=52 T=42 W=180 H=24");
        _output.WriteLine("bar observed L=" + shape.Left + " T=" + shape.Top
            + " W=" + shape.Width + " H=" + shape.Height);

        AssertPoint("bar Left", 52d, shape.Left);
        AssertPoint("bar Top", 42d, shape.Top);
        AssertPoint("bar Width", 180d, shape.Width);
        AssertPoint("bar Height", 24d, shape.Height);

        // The translation is what keeps this shape off the origin. Had the
        // renderer passed scene coordinates through, the host would have clamped
        // the negative offsets to zero and these four assertions would fail.
        Assert.True(shape.Left > 0f && shape.Top > 0f,
            "The bar landed at the origin, so the host's negative-offset clamp fired.");

        AssertOwnership(shape, "row-1:bar");
    }

    private void AssertLine(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        Excel.Shape shape = scope.Track(sheet.Shapes.Item("chart:grid:0"));

        // Scene endpoints (40, 20) to (220, 20) translated by the 12pt padding:
        // a horizontal left-to-right line, so the bounding box is exact.
        _output.WriteLine("line expected L=52 T=32 W=180 H=0");
        _output.WriteLine("line observed L=" + shape.Left + " T=" + shape.Top
            + " W=" + shape.Width + " H=" + shape.Height);

        // Shape exposes only Left/Top/Width/Height, so a line's endpoints are
        // observed through its bounding box.
        AssertPoint("line Left (from.X)", 52d, shape.Left);
        AssertPoint("line Top (from.Y)", 32d, shape.Top);
        AssertPoint("line Width (to.X - from.X)", 180d, shape.Width);
        AssertPoint("line Height (to.Y - from.Y)", 0d, shape.Height);

        AssertOwnership(shape, "chart:grid:0");
    }

    private static void AssertOwnership(Excel.Shape shape, string primitiveId)
    {
        Assert.Equal(primitiveId, shape.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId(primitiveId), shape.AlternativeText);
        Assert.True(ShapeOwnershipTag.IsOwnedTag(shape.AlternativeText));
    }

    private static void AssertPoint(string what, double expected, double actual)
    {
        Assert.True(
            GeometryMath.ApproximatelyEqual(expected, actual),
            string.Create(
                CultureInfo.InvariantCulture,
                $"{what}: expected {expected:R} but the live host reported {actual:R}; "
                    + $"difference {Math.Abs(expected - actual):R} exceeds the documented "
                    + $"tolerance {GeometryMath.Epsilon:R}."));
    }
}
