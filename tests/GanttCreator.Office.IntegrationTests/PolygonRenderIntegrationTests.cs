using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R4.5's Required Office gate: a four-point milestone diamond becomes a real
/// Excel freeform at the scene's exact points, and the scene's back-to-front
/// order is applied to the live sheet.
/// </summary>
/// <remarks>
/// <para>
/// The expected geometry is hand-written, never derived from the renderer, for
/// the reason the R4.3 gate's comment gives.
/// </para>
/// <para>
/// The freeform assertions deliberately include the <strong>tip-to-tip
/// extent</strong>, not just the position. A milestone that collapsed to a
/// degenerate shape would sit exactly where a real one would, so a
/// position-only test would pass against a zero-size diamond.
/// </para>
/// <para>
/// Every geometry assertion uses the single documented
/// <see cref="GeometryMath.Epsilon"/>. There is no second tolerance: the host
/// places a diamond auto-shape exactly (0.000 EMU across exact spans of 10, 20,
/// 37.5, 100 and 253 pt), so the milestone family needs no exception to the
/// project's one rounding boundary.
/// </para>
/// </remarks>
public class PolygonRenderIntegrationTests(ITestOutputHelper output)
{
    private const double Padding = 12d;
    private const double CentreY = 100d;
    private const double Radius = 10d;

    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_milestone_diamond_renders_as_a_diamond_auto_shape_and_the_scene_order_is_applied()
    {
        var fixture = new OfficeFixture();

        // L19: this test creates REAL shapes through the real ExcelShapeWriter,
        // and a workbook holding a live shape keeps the Excel process alive, so
        // teardown legitimately escalates to a kill.
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
                ChartOriginDelta.ForChartBounds(new RectD(-Padding, -Padding, 400d, 300d)));
            SceneTranslationOutcome translated = renderer.Translate(scene);
            Assert.True(translated.Complete, translated.Deferred.Count + " primitives were deferred.");
            Assert.Equal(3, translated.Requests.Count);

            // The bar is on ActivityBody (40) and the markers on Milestone (60),
            // so the scene's own back-to-front order is bar, marker-a, marker-b.
            Assert.Equal(
                ["row-1:bar", "row-1:marker-a", "row-1:marker-b"],
                translated.Requests.Select(request => request.PrimitiveId));

            var writer = new ExcelShapeWriter(fixture.Excel);
            foreach (OfficeShapeRequest request in translated.Requests)
            {
                ShapeWriteOutcome outcome = writer.Create(request);
                Assert.True(outcome.Succeeded, $"Create refused for '{request.PrimitiveId}': {outcome.Refusal}");
            }

            ShapeWriteOutcome zOrder = writer.ApplyZOrder(
                [.. translated.Requests.Select(request => request.PrimitiveId)]);
            Assert.True(zOrder.Succeeded, $"ApplyZOrder refused: {zOrder.Refusal}");

            var sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            // Scene diamond centres 80 and 100; each spans centre +/- radius.
            AssertDiamond(scope, sheet, "row-1:marker-a", 80d);
            AssertDiamond(scope, sheet, "row-1:marker-b", 100d);

            Excel.Shape bar = scope.Track(sheet.Shapes.Item("row-1:bar"));
            AssertPoint("bar Left", 72d, bar.Left);
            AssertPoint("bar Top", 108d, bar.Top);
            AssertPoint("bar Width", 40d, bar.Width);
            AssertPoint("bar Height", 8d, bar.Height);

            AssertZOrder(scope, sheet, bar);
            AssertOwnershipAndCleanup(scope, sheet);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static GanttScene BuildScene()
    {
        var style = new SceneStyle("Default");

        // Two same-date milestones 20pt apart plus a bar beneath them: the
        // overlap case the entity guide requires be ordered, not moved.
        return GanttScene.TryCreate(
                new RectD(-Padding, -Padding, 400d, 300d),
                new RectD(0d, 0d, 380d, 280d),
                [
                    new SceneRect("row-1:bar", SceneOwnerId.Chart, ZLayer.ActivityBody, new RectD(60d, 96d, 40d, 8d), style),
                    new ScenePolygon("row-1:marker-a", SceneOwnerId.Chart, ZLayer.Milestone, Diamond(80d), style),
                    new ScenePolygon("row-1:marker-b", SceneOwnerId.Chart, ZLayer.Milestone, Diamond(100d), style),
                ],
                [])
            .Scene ?? throw new InvalidOperationException("The fixture scene failed to validate.");
    }

    /// <summary>Builds a four-point diamond centred on the given X.</summary>
    /// <param name="centreX">The diamond's horizontal centre.</param>
    /// <returns>The diamond's four points in draw order.</returns>
    private static PointD[] Diamond(double centreX) =>
    [
        new(centreX, CentreY - Radius),
        new(centreX + Radius, CentreY),
        new(centreX, CentreY + Radius),
        new(centreX - Radius, CentreY),
    ];

    /// <summary>Asserts a live freeform matches the box its scene points describe.</summary>
    /// <param name="scope">The COM scope.</param>
    /// <param name="sheet">The worksheet holding the shape.</param>
    /// <param name="primitiveId">The shape name.</param>
    /// <param name="centreX">The diamond's scene centre X.</param>
    private void AssertDiamond(
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet,
        string primitiveId,
        double centreX)
    {
        Excel.Shape shape = scope.Track(sheet.Shapes.Item(primitiveId));

        _output.WriteLine(primitiveId + " observed L=" + shape.Left + " T=" + shape.Top
            + " W=" + shape.Width + " H=" + shape.Height);

        // The bounding box of a diamond centred at centreX spans centreX-Radius to
        // centreX+Radius, translated by the 12pt padding.
        AssertPoint(primitiveId + " Left", centreX - Radius + Padding, shape.Left);
        AssertPoint(primitiveId + " Top", CentreY - Radius + Padding, shape.Top);

        // A 20pt tip-to-tip diamond, so the bounding box is 20x20. Asserting the
        // extent is what catches a collapsed shape, which would report a
        // plausible position with ~zero width.
        AssertPoint(primitiveId + " Width (tip-to-tip)", Radius * 2, shape.Width);
        AssertPoint(primitiveId + " Height (tip-to-tip)", Radius * 2, shape.Height);
    }

    /// <summary>
    /// The scene's order must be the host's. <c>ZOrderPosition</c> is read-only
    /// and higher means nearer the front, so this is the host's own answer rather
    /// than a restatement of what the adapter asked for.
    /// </summary>
    /// <param name="scope">The COM scope.</param>
    /// <param name="sheet">The worksheet holding the shapes.</param>
    /// <param name="bar">The bar shape, already tracked.</param>
    private void AssertZOrder(OfficeFixture.ComScope scope, Excel.Worksheet sheet, Excel.Shape bar)
    {
        int barPosition = bar.ZOrderPosition;
        int markerAPosition = scope.Track(sheet.Shapes.Item("row-1:marker-a")).ZOrderPosition;
        int markerBPosition = scope.Track(sheet.Shapes.Item("row-1:marker-b")).ZOrderPosition;

        _output.WriteLine("z-order positions: bar=" + barPosition
            + " marker-a=" + markerAPosition + " marker-b=" + markerBPosition);

        Assert.True(barPosition < markerAPosition, "The bar must sit behind the first milestone.");
        Assert.True(markerAPosition < markerBPosition, "The second milestone must sit in front of the first.");
    }

    /// <summary>Asserts every rendered shape is owned, then deletes it.</summary>
    /// <param name="scope">The COM scope.</param>
    /// <param name="sheet">The worksheet holding the shapes.</param>
    private static void AssertOwnershipAndCleanup(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        foreach (string id in new[] { "row-1:bar", "row-1:marker-a", "row-1:marker-b" })
        {
            Excel.Shape shape = scope.Track(sheet.Shapes.Item(id));
            Assert.Equal(id, shape.Name);
            Assert.Equal(ShapeOwnershipTag.ForPrimitiveId(id), shape.AlternativeText);
            shape.Delete();
        }
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
