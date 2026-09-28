using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R4.7's Required Office gate: two refreshes over the same scene produce the
/// same owned identifiers, the same count, the same bounds, and no duplicates,
/// and the user's selection survives both.
/// </summary>
/// <remarks>
/// <para>
/// The contract tests prove the plan is a fixed point; this proves the <em>host</em>
/// converges to it. A planner bug would be caught there and a writer bug only
/// here — in particular a duplicate shape, which is the one failure a user
/// notices immediately and which no pure assertion can detect.
/// </para>
/// <para>
/// Bounds are read back off the live shapes rather than taken from the request,
/// so a host-side clamp or a stale shape left over from the first refresh fails
/// the test instead of passing on the adapter's own arithmetic.
/// </para>
/// </remarks>
public class RefreshIdempotenceIntegrationTests(ITestOutputHelper output)
{
    private const double PaddingPt = 12d;

    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Two_refreshes_leave_identical_owned_shapes_with_no_duplicates_and_intact_selection()
    {
        var fixture = new OfficeFixture();

        // KNOWN-LIMITATIONS L19, for the same reason the other render tests need
        // it: this test creates REAL shapes through the real writer and reads
        // them back, and a workbook holding live shapes keeps the Excel process
        // alive, so teardown escalates to a kill. Every idempotence claim is
        // still asserted from the host, so nothing is hidden by the suppression.
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

            var renderer = new SceneShapeRenderer(
                ChartOriginDelta.ForChartBounds(new RectD(-PaddingPt, -PaddingPt, 400d, 300d)));
            var writer = new ExcelShapeWriter(fixture.Excel);

            GanttScene scene = BuildScene();
            SceneTranslationOutcome translated = renderer.Translate(scene);
            Assert.True(translated.Complete,
                $"{translated.Deferred.Count} deferred, {translated.Refusals.Count} refused.");
            OfficeShapeRequest[] desired = [.. translated.Requests];
            Assert.NotEmpty(desired);

            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            // Put the user's cursor somewhere specific, so a scope that failed to
            // capture and restore it would be caught rather than tolerated.
            Excel.Range anchor = (Excel.Range)scope.Track(sheet.Cells[5, 3]);
            anchor.Select();
            object? selectionBefore = fixture.Excel.Selection;

            ShapeReconcileOutcome first = ShapeReconciler.Reconcile(
                writer,
                new ShapeReconcileRequest(desired));

            Assert.True(first.Succeeded, $"First refresh refused: {first.Refusal} / {first.Failure}");
            IReadOnlyList<string> ownedAfterFirst = writer.ListOwned();
            List<ShapeBounds> boundsAfterFirst = ReadBounds(scope, sheet, ownedAfterFirst);

            // Second refresh, same scene. This is the live fixed point.
            ShapeReconcileOutcome second = ShapeReconciler.Reconcile(
                writer,
                new ShapeReconcileRequest(desired));

            Assert.True(second.Succeeded, $"Second refresh refused: {second.Refusal} / {second.Failure}");
            IReadOnlyList<string> ownedAfterSecond = writer.ListOwned();
            List<ShapeBounds> boundsAfterSecond = ReadBounds(scope, sheet, ownedAfterSecond);

            _output.WriteLine($"first owned:  {ownedAfterFirst.Count}");
            _output.WriteLine($"second owned: {ownedAfterSecond.Count}");

            // Identical identifiers, count, and bounds across both refreshes.
            Assert.Equal(ownedAfterFirst, ownedAfterSecond);
            Assert.Equal(ownedAfterFirst.Count, ownedAfterSecond.Count);
            Assert.Equal(boundsAfterFirst, boundsAfterSecond);

            // No duplicates. ListOwned returns ordinal-sorted names, so a duplicate
            // would also show as a length mismatch; the explicit check names the
            // failure a user would actually see.
            Assert.Equal(
                ownedAfterSecond.Count,
                ownedAfterSecond.Distinct(StringComparer.Ordinal).Count());

            // Every declared identifier is present exactly once, which is the
            // property a duplicated shape would break.
            Assert.Equal(
                desired.Select(request => request.PrimitiveId).Order(StringComparer.Ordinal),
                ownedAfterSecond.Order(StringComparer.Ordinal));

            // Selection survives both refreshes (D4). The anchor's address is
            // compared rather than a proxy identity, because Selection is
            // re-proxied on every read and a reference comparison would report
            // corruption where none occurred.
            Assert.NotNull(selectionBefore);
            Assert.Equal(
                ((Excel.Range)selectionBefore!).Address,
                ((Excel.Range)fixture.Excel.Selection).Address);

            // A blocked refresh must not disturb the converged chart.
            ShapeReconcileOutcome blocked = ShapeReconciler.Reconcile(
                writer,
                new ShapeReconcileRequest(desired, HasBlockingErrors: true));

            Assert.Equal(ShapeReconcileRefusal.BlockingErrors, blocked.Refusal);
            Assert.False(blocked.MutatedAnyShape);
            Assert.Equal(ownedAfterSecond, writer.ListOwned());
            Assert.Equal(boundsAfterSecond, ReadBounds(scope, sheet, ownedAfterSecond));

            // Delete the shapes but leave the workbook open for the fixture to
            // close, as the other render tests do.
            foreach (string identifier in ownedAfterSecond)
            {
                scope.Track(sheet.Shapes.Item(identifier)).Delete();
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>One shape's live geometry, for comparing two refreshes.</summary>
    private readonly record struct ShapeBounds(
        string Identifier,
        double Left,
        double Top,
        double Width,
        double Height);

    /// <summary>
    /// Reads the host's reported geometry for each owned shape. A first refresh
    /// that silently failed to move a shape would be invisible without this.
    /// </summary>
    private static List<ShapeBounds> ReadBounds(
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet,
        IReadOnlyList<string> identifiers)
    {
        List<ShapeBounds> bounds = [];
        foreach (string identifier in identifiers)
        {
            Excel.Shape shape = (Excel.Shape)scope.Track(sheet.Shapes.Item(identifier));
            bounds.Add(
                new ShapeBounds(
                    identifier,
                    GeometryMath.SnapToDisplayPrecision(shape.Left),
                    GeometryMath.SnapToDisplayPrecision(shape.Top),
                    GeometryMath.SnapToDisplayPrecision(shape.Width),
                    GeometryMath.SnapToDisplayPrecision(shape.Height)));
        }

        return bounds;
    }

    /// <summary>
    /// A small scene mixing shape families — a background, a bar, a label, and a
    /// grid line — so the gate covers more than rectangles.
    /// </summary>
    private static GanttScene BuildScene()
    {
        SceneOwnerId owner = SceneOwnerId.Chart;
        var style = new SceneStyle("Default");

        SceneRect background = new(
            "chart:background",
            owner,
            ZLayer.Background,
            new RectD(0, 0, 380, 260),
            style);
        SceneRect bar = new("row-1:bar", owner, ZLayer.ActivityBody, new RectD(40, 60, 120, 20), style);
        SceneText label = new(
            "row-1:bar-label",
            owner,
            ZLayer.Label,
            "Procurement",
            new RectD(168, 60, 80, 20),
            style,
            GanttTextAlignment.Left);
        SceneLine grid = new(
            "chart:grid:1",
            owner,
            ZLayer.Grid,
            new PointD(0, 50),
            new PointD(380, 50),
            style);

        return GanttScene
            .TryCreate(
                new RectD(-PaddingPt, -PaddingPt, 400d, 300d),
                new RectD(0, 0, 380, 260),
                [background, bar, label, grid],
                [])
            .Scene
            ?? throw new InvalidOperationException("The fixture scene failed to validate.");
    }
}
