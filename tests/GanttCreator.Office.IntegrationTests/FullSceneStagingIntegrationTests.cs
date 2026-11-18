using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R6.3 Office gate: the complete export scene - panel plus every chart entity - staged
/// beside real content at the deterministic origin, with the host's own read-back proving
/// the bounds-equality contract, the scene z-order, the scene-owned placement class, and
/// that neither the workbook's cells, the user's sentinel shape, nor the view was
/// disturbed. A live coalesced-refresh simulation then reconciles the live chart while the
/// staged set exists.
/// </summary>
/// <remarks>
/// The fixture drives the production chain: initialise a real table, write real rows, read
/// and validate them, measure the real grid, read the workbook's settings and styles, build
/// the EXPORT request through <see cref="ExcelSceneBuildRequestFactory.CreateExport"/>,
/// resolve the bounds through <see cref="ExportBoundsResolver"/>, and stage through an
/// <see cref="ExcelShapeWriter"/> constructed with
/// <see cref="ShapeOwnershipScheme.Staging"/>. The staging operation runs inside the R4.2
/// application-state scope, and the scroll/zoom/selection readings either side are the D3
/// evidence.
/// </remarks>
public sealed class FullSceneStagingIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_full_scene_stages_beside_real_content_and_cleans_up()
    {
        OfficeFixture fixture = new();

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            using OfficeFixture.ComScope scope = new();

            Excel.Workbook workbook = fixture.CreateWorkbook();
            WorkbookInitialiseOutcome initialised =
                new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            // The add-in is not loaded here, so no coalesced pass races these writes; the
            // chart lock is still switched off through the user setting, as every
            // operating body does, so a later protection state cannot refuse a write.
            OfficeFixture.DisableChartProtection(workbook);

            Excel.Sheets sheets = scope.Track(workbook.Worksheets);
            int sheetCountBefore = sheets.Count;
            Excel.Worksheet sheet = scope.Track(
                (Excel.Worksheet)sheets[GanttWorkbookContract.GanttSheetLabel]);
            Excel.ListObject table = ResolveTable(scope, sheet);
            Excel.ListRows listRows = scope.Track(table.ListRows);

            while (listRows.Count > 0)
            {
                scope.Track(listRows[1]).Delete();
            }

            Excel.ListRow spanRow = scope.Track(listRows.Add());
            Excel.ListRow milestoneRow = scope.Track(listRows.Add());
            SetRowCell(scope, spanRow, "Id", GanttRowId.New().Value);
            SetRowCell(scope, spanRow, "Type", "As-Planned Activity");
            SetRowCell(scope, spanRow, "Description", "Build phase");
            SetRowCell(scope, spanRow, "Start", 45658.0);
            SetRowCell(scope, spanRow, "Finish", 45670.0);
            SetRowCell(scope, milestoneRow, "Id", GanttRowId.New().Value);
            SetRowCell(scope, milestoneRow, "Type", "As-Planned Milestone");
            SetRowCell(scope, milestoneRow, "Start", 45675.0);
            TrimBlankBodyRows(scope, fixture.Excel, listRows);

            GanttTableReadOutcome read = new ExcelGanttTableReader(fixture.Excel).Read();
            Assert.True(read.Succeeded, $"Read refused: {read.Refusal}");
            GanttValidationOutcome validation = GanttRowValidator.Validate(read.Rows);
            Assert.Equal(validation.Events.Count, read.Rows.Count);
            Assert.Equal(2, validation.Events.Count);

            string[] measuredColumns =
            [
                .. GanttTableSchema.Default.Columns
                    .Where(column => !column.IsHidden)
                    .Select(column => column.Name),
            ];
            PanelGridOutcome gridOutcome =
                new ExcelPanelGridMeasurement(fixture.Excel).Measure(measuredColumns);
            Assert.True(gridOutcome.Succeeded, $"Grid refused: {gridOutcome.Refusal}");
            PanelCellGrid grid = gridOutcome.Grid!;

            ConfigReadOutcome config = new ExcelConfigCatalogueReader(fixture.Excel).Read();
            Assert.True(config.Succeeded, $"Settings read refused: {config.Refusal}");

            // The production measurement seam, exactly as the Refresh command constructs
            // the factory: the scene cannot be composed without one.
            SceneBuildRequestOutcome request = new ExcelSceneBuildRequestFactory(new AptosTextMetrics()).CreateExport(
                validation.Events,
                config.Settings,
                config.Styles,
                grid,
                config.Metrics);
            Assert.True(request.Succeeded, $"Export request refused: {request.Message}");

            ExportBoundsOutcome resolution = ExportBoundsResolver.TryResolve(
                request.Request,
                ExportPanelTheme.Resolve(),
                ExportSelection.TablePlusChart);
            Assert.True(resolution.Succeeded, $"Resolution refused: {resolution.Refusal}");
            Assert.NotNull(resolution.Scene);
            Assert.True(resolution.Bounds is not null);

            ExcelShapeWriter stagingWriter = new(
                fixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);
            ExcelShapeWriter liveWriter = new(fixture.Excel);
            Assert.True(liveWriter.Create(LiveProbe("live:probe-1")).Succeeded);
            Assert.True(liveWriter.Create(LiveProbe("live:probe-2")).Succeeded);

            // The R4.8 sentinels: an unowned user shape with a name that could collide
            // with nothing, and a witness cell. Neither may be touched by the staging
            // operation, its cleanup, or the live reconciliation.
            Excel.Shapes shapes = scope.Track(sheet.Shapes);
            Excel.Shape sentinel = scope.Track(shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                40,
                500,
                120,
                24));
            sentinel.Name = "User note";
            sentinel.TextFrame2.TextRange.Text = "Do not touch";
            Excel.Range spanRange = scope.Track(spanRow.Range);
            Excel.Range witness = scope.Track(spanRange.Cells[1, ColumnIndex("Description")]);
            object? witnessBefore = witness.Value2;

            // D3: set a non-default view so the after-reading would notice a disturbance.
            Excel.Window window = scope.Track(fixture.Excel.ActiveWindow);
            window.ScrollRow = 2;
            window.ScrollColumn = 1;
            window.Zoom = 90;
            int scrollRowBefore = window.ScrollRow;
            int scrollColumnBefore = window.ScrollColumn;
            double zoomBefore = window.Zoom;
            string selectionBefore = SelectionAddress(scope, fixture.Excel);

            StagingCompositionOutcome staged;
            using (ExcelApplicationStateScope state = new(fixture.Excel))
            {
                state.SuppressScreenUpdating();
                state.SuppressEvents();
                state.SuppressAlerts();
                state.CaptureSelection();
                staged = EditablePanelStagingComposer.ComposeFullScene(resolution, stagingWriter);
            }

            string[] stagedNames = [];
            try
            {
                Assert.True(staged.Succeeded, $"Composition refused: {staged.Refusal}");
                Assert.True(staged.StagedBounds is not null);

                stagedNames =
                [
                    .. staged.StagedPrimitiveIds.Select(id => ShapeStagingTag.NamePrefix + id),
                ];

                // Completeness: every scene primitive staged exactly once.
                Assert.Equal(resolution.Scene.Primitives.Count, stagedNames.Length);

                Excel.Shapes stagedCollection = scope.Track(sheet.Shapes);
                Dictionary<string, Excel.Shape> stagedByName = [];
                int shapeCount = stagedCollection.Count;
                for (int index = 1; index <= shapeCount; index++)
                {
                    Excel.Shape shape = scope.Track(stagedCollection.Item(index));
                    if (shape.Name.StartsWith(ShapeStagingTag.NamePrefix, StringComparison.Ordinal))
                    {
                        stagedByName[shape.Name] = shape;
                    }
                }

                Assert.Equal(stagedNames.Length, stagedByName.Count);

                RectD envelope = staged.StagedBounds!.Value;
                List<RectD> hostEnvelopes = [];
                int previousZOrder = int.MinValue;
                foreach (string primitiveId in staged.StagedPrimitiveIds)
                {
                    Assert.True(
                        stagedByName.TryGetValue(ShapeStagingTag.NamePrefix + primitiveId, out Excel.Shape? shape),
                        $"A staged primitive is missing: {primitiveId}");
                    Assert.NotNull(shape);

                    // Staging identity, proved from the host's alternative text.
                    Assert.True(ShapeStagingTag.IsStagingTag(shape.AlternativeText));
                    Assert.False(ShapeOwnershipTag.IsOwnedTag(shape.AlternativeText));

                    // D1: staged shapes never enter the host-owned resize class.
                    Assert.Equal(Excel.XlPlacement.xlMove, shape.Placement);

                    // Paint order: created back to front, so z-order strictly increases.
                    Assert.True(shape.ZOrderPosition > previousZOrder);
                    previousZOrder = shape.ZOrderPosition;

                    RectD hostEnvelope = new(
                        Convert.ToDouble(shape.Left, CultureInfo.InvariantCulture),
                        Convert.ToDouble(shape.Top, CultureInfo.InvariantCulture),
                        Convert.ToDouble(shape.Width, CultureInfo.InvariantCulture),
                        Convert.ToDouble(shape.Height, CultureInfo.InvariantCulture));
                    hostEnvelopes.Add(hostEnvelope);
                }

                // D2 on the host: the staged set's union equals the R6.1 export bounds.
                // The EXACT GeometryMath.Epsilon equality is pinned by
                // FullSceneStagingComposerTests on the translated requests; the host
                // stores shape geometry in single precision (measured 2026-11-18:
                // 514.3622047pt requested, 514.3621826pt read back, ~2.2e-5pt), so the
                // live read-back uses the documented host tolerance - above float
                // rounding, below display resolution - and never a widened scene
                // tolerance or a padded bound.
                RectD union = Union(hostEnvelopes);
                double hostTolerance = HostOwnedGeometry.GeometryMatchTolerancePt;
                Assert.True(
                    Math.Abs(union.Left - envelope.Left) <= hostTolerance
                    && Math.Abs(union.Top - envelope.Top) <= hostTolerance
                    && Math.Abs(union.Right - envelope.Right) <= hostTolerance
                    && Math.Abs(union.Bottom - envelope.Bottom) <= hostTolerance,
                    $"Host envelope {union} does not equal {envelope} within {hostTolerance}pt.");
                Assert.All(
                    hostEnvelopes,
                    candidate => Assert.True(
                        candidate.Left >= envelope.Left - hostTolerance
                        && candidate.Top >= envelope.Top - hostTolerance
                        && candidate.Right <= envelope.Right + hostTolerance
                        && candidate.Bottom <= envelope.Bottom + hostTolerance,
                        $"A staged shape escapes the envelope: {candidate}"));

                // The full scene carries the families the panel-only row never produced.
                Assert.Contains(stagedByName.Values, shape => shape.Type == MsoShapeType.msoLine);
                Assert.Contains(
                    stagedByName.Values,
                    shape => shape.AutoShapeType == MsoAutoShapeType.msoShapeDiamond);
                Assert.Contains(
                    stagedByName.Values,
                    shape => shape.AutoShapeType == MsoAutoShapeType.msoShapeRectangle);
                Assert.Contains(stagedByName.Values, shape => shape.Type == MsoShapeType.msoTextBox);

                // D5 live: a coalesced refresh reconciles the live chart while the staged
                // set exists, and is blind to it. The staging operation writes no cells,
                // so this simulates the concurrent user edit that schedules one.
                ShapeReconcileOutcome reconciled = ShapeReconciler.Reconcile(
                    liveWriter,
                    new ShapeReconcileRequest(
                        [LiveProbe("live:probe-1"), LiveProbe("live:probe-3")]));
                Assert.Null(reconciled.Refusal);
                Assert.Equal(["live:probe-1", "live:probe-3"], liveWriter.ListOwned());

                Excel.Shapes afterReconcile = scope.Track(sheet.Shapes);
                int afterReconcileCount = afterReconcile.Count;
                for (int index = 1; index <= afterReconcileCount; index++)
                {
                    Excel.Shape shape = scope.Track(afterReconcile.Item(index));
                    bool isStaged = shape.Name.StartsWith(ShapeStagingTag.NamePrefix, StringComparison.Ordinal);
                    if (isStaged)
                    {
                        Assert.Contains(shape.Name, stagedNames);
                    }
                }

                Assert.Equal(stagedNames.Length, CountStaged(scope, sheet));

                // No cells, no sheet, no sentinels were touched.
                Assert.Equal(sheetCountBefore, sheets.Count);
                Assert.Equal(witnessBefore, witness.Value2);
                Assert.True(sentinel.Name == "User note");
                Assert.Equal("Do not touch", sentinel.TextFrame2.TextRange.Text);

                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"Excel build {fixture.Excel.Version}; staged={stagedNames.Length}"
                    + $" envelope={envelope.X:F1},{envelope.Y:F1},{envelope.Width:F1},{envelope.Height:F1}"));
            }
            finally
            {
                // Cleanup through the staging scheme only, newest first, so the live
                // chart could not be touched even under a scheme regression.
                for (int index = staged.StagedPrimitiveIds.Count - 1; index >= 0; index--)
                {
                    _ = stagingWriter.Delete(staged.StagedPrimitiveIds[index]);
                }
            }

            // Zero staged residue; the sentinel and the live chart survive cleanup.
            Assert.Equal(0, CountStaged(scope, sheet));
            Assert.Equal(sheetCountBefore, sheets.Count);
            Assert.Equal(witnessBefore, witness.Value2);
            Assert.True(sentinel.Name == "User note");
            Assert.Equal(["live:probe-1", "live:probe-3"], liveWriter.ListOwned());

            // D3: the view state the operation ran under is still the user's.
            Assert.Equal(scrollRowBefore, window.ScrollRow);
            Assert.Equal(scrollColumnBefore, window.ScrollColumn);
            Assert.Equal(zoomBefore, window.Zoom, precision: 6);
            Assert.Equal(selectionBefore, SelectionAddress(scope, fixture.Excel));

            // Live cleanup, so the fixture leaves no owned probe behind.
            foreach (string identifier in liveWriter.ListOwned())
            {
                _ = liveWriter.Delete(identifier);
            }

            sentinel.Delete();
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static RectD Union(IReadOnlyList<RectD> envelopes)
    {
        double left = envelopes.Min(envelope => envelope.Left);
        double top = envelopes.Min(envelope => envelope.Top);
        double right = envelopes.Max(envelope => envelope.Right);
        double bottom = envelopes.Max(envelope => envelope.Bottom);
        return new RectD(left, top, right - left, bottom - top);
    }

    private static OfficeShapeRequest LiveProbe(string id) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(700, 480, 40, 12)),
            ZLayer.ActivityBody);

    private static int CountStaged(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        Excel.Shapes shapes = scope.Track(sheet.Shapes);
        int count = 0;
        int total = shapes.Count;
        for (int index = 1; index <= total; index++)
        {
            Excel.Shape shape = scope.Track(shapes.Item(index));
            if (shape.Name.StartsWith(ShapeStagingTag.NamePrefix, StringComparison.Ordinal))
            {
                count++;
            }
        }

        return count;
    }

    private static string SelectionAddress(OfficeFixture.ComScope scope, Excel.Application application)
    {
        object selection = scope.Track(application.Selection);
        return selection is Excel.Range range
            ? Convert.ToString(range.Address, CultureInfo.InvariantCulture)
            : string.Empty;
    }

    private static Excel.ListObject ResolveTable(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        Excel.ListObjects objects = scope.Track(sheet.ListObjects);
        return scope.Track(objects[GanttTableSchema.TableName]);
    }

    private static void SetRowCell(
        OfficeFixture.ComScope scope,
        Excel.ListRow row,
        string headerName,
        object value)
    {
        Excel.Range range = scope.Track(row.Range);
        Excel.Range cell = scope.Track(range.Cells[1, ColumnIndex(headerName)]);
        cell.Value2 = value;
    }

    private static void TrimBlankBodyRows(
        OfficeFixture.ComScope scope,
        Excel.Application application,
        Excel.ListRows listRows)
    {
        for (int index = listRows.Count; index >= 1; index--)
        {
            Excel.ListRow row = scope.Track(listRows[index]);
            Excel.Range range = scope.Track(row.Range);
            double nonEmpty = Convert.ToDouble(
                application.WorksheetFunction.CountA(range),
                CultureInfo.InvariantCulture);
            if (nonEmpty == 0)
            {
                row.Delete();
            }
        }
    }

    private static int ColumnIndex(string headerName)
    {
        for (int index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            if (string.Equals(GanttTableSchema.Default.Columns[index].Name, headerName, StringComparison.Ordinal))
            {
                return index + 1;
            }
        }

        throw new InvalidOperationException($"'{headerName}' is not a schema column.");
    }
}
