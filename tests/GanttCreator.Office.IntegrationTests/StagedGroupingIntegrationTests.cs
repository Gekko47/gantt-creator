using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R6.4 Office gate: the staged editable-export set groups into one shape group with the
/// scene's child order, the group carries the staging identity, its envelope equals the
/// R6.3 staged bounds on the host, and ungrouping restores every child as an editable
/// shape whose geometry survives an edit.
/// </summary>
/// <remarks>
/// The fixture drives the production chain exactly as R6.3's gate does: initialise a real
/// table, write real rows, read and validate them, measure the real grid, read the
/// workbook's settings and styles, build the EXPORT request through
/// <see cref="ExcelSceneBuildRequestFactory.CreateExport"/>, resolve the bounds through
/// <see cref="ExportBoundsResolver"/>, stage through an <see cref="ExcelShapeWriter"/>
/// constructed with <see cref="ShapeOwnershipScheme.Staging"/>, and group through the real
/// <see cref="ExcelShapeGrouper"/> inside the R4.2 application-state scope. The
/// group/ungroup/edit sequence is the work item's Required host gate (D3).
/// </remarks>
public sealed class StagedGroupingIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_staged_set_groups_ungroups_and_stays_editable()
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

            ExcelShapeWriter stagingWriter = new(
                fixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);
            ExcelShapeWriter liveWriter = new(fixture.Excel);
            ExcelShapeGrouper grouper = new(fixture.Excel);
            Assert.True(liveWriter.Create(LiveProbe("live:probe-1")).Succeeded);
            Assert.True(liveWriter.Create(LiveProbe("live:probe-2")).Succeeded);

            // R4.8 sentinels and the view state the operation must not disturb.
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

            Excel.Window window = scope.Track(fixture.Excel.ActiveWindow);
            window.ScrollRow = 2;
            window.ScrollColumn = 1;
            window.Zoom = 90;
            int scrollRowBefore = window.ScrollRow;
            int scrollColumnBefore = window.ScrollColumn;
            double zoomBefore = window.Zoom;
            string selectionBefore = SelectionAddress(scope, fixture.Excel);

            StagingCompositionOutcome staged;
            StagedGroupOutcome grouped;
            using (ExcelApplicationStateScope state = new(fixture.Excel))
            {
                state.SuppressScreenUpdating();
                state.SuppressEvents();
                state.SuppressAlerts();
                state.CaptureSelection();
                staged = EditablePanelStagingComposer.ComposeFullScene(resolution, stagingWriter);
                grouped = StagedShapeGrouper.Group(staged, grouper);
            }

            string[] stagedNames =
            [
                .. staged.StagedPrimitiveIds.Select(id => ShapeStagingTag.NamePrefix + id),
            ];

            try
            {
                Assert.True(staged.Succeeded, $"Composition refused: {staged.Refusal}");
                Assert.True(grouped.Succeeded, $"Grouping refused: {grouped.Refusal}");

                RectD envelope = staged.StagedBounds!.Value;
                double hostTolerance = HostOwnedGeometry.GeometryMatchTolerancePt;

                // D1: the group exists under the deterministic staging identity, carries the
                // staging tag (so the live scheme is blind to it) and the temporary
                // placement class.
                Excel.Shape group = scope.Track(shapes.Item(ShapeStagingTag.GroupName));
                Assert.Equal(ShapeStagingTag.GroupName, group.Name);
                Assert.True(ShapeStagingTag.IsStagingTag(group.AlternativeText));
                Assert.False(ShapeOwnershipTag.IsOwnedTag(group.AlternativeText));
                Assert.Equal(Excel.XlPlacement.xlMove, group.Placement);

                Assert.Equal([ShapeStagingTag.GroupName], stagingWriter.ListOwned());
                Assert.Equal(["live:probe-1", "live:probe-2"], liveWriter.ListOwned());

                // D2 live pin (the Tier-A probe): GroupItems reports the scene's creation
                // order. A host that re-ordered is the work item's stop condition - this
                // assertion is the evidence, and it is never widened to accept a quirk.
                Excel.GroupShapes items = scope.Track(group.GroupItems);
                Assert.Equal(stagedNames.Length, items.Count);
                for (int index = 1; index <= items.Count; index++)
                {
                    Excel.Shape child = scope.Track(items.Item(index));
                    Assert.Equal(stagedNames[index - 1], child.Name);
                }

                Assert.Equal(stagedNames, grouped.ChildOrder);

                // The group envelope equals the R6.3 staged bounds within the documented
                // host tolerance (single-precision storage; measured 2026-11-18).
                RectD hostEnvelope = new(
                    Convert.ToDouble(group.Left, CultureInfo.InvariantCulture),
                    Convert.ToDouble(group.Top, CultureInfo.InvariantCulture),
                    Convert.ToDouble(group.Width, CultureInfo.InvariantCulture),
                    Convert.ToDouble(group.Height, CultureInfo.InvariantCulture));
                Assert.True(
                    Math.Abs(hostEnvelope.Left - envelope.Left) <= hostTolerance
                    && Math.Abs(hostEnvelope.Top - envelope.Top) <= hostTolerance
                    && Math.Abs(hostEnvelope.Right - envelope.Right) <= hostTolerance
                    && Math.Abs(hostEnvelope.Bottom - envelope.Bottom) <= hostTolerance,
                    $"Group envelope {hostEnvelope} does not equal {envelope} within {hostTolerance}pt.");

                // No cells, no sheet, no sentinels were touched.
                Assert.Equal(sheetCountBefore, sheets.Count);
                Assert.Equal(witnessBefore, witness.Value2);
                Assert.Equal("User note", sentinel.Name);
                Assert.Equal("Do not touch", sentinel.TextFrame2.TextRange.Text);

                // D3: ungroup (inspection, not a feature), then prove every expected child is
                // editable and its geometry survived.
                RectD[] before = CaptureChildEnvelopes(scope, group, stagedNames.Length);
                ShapeWriteOutcome ungrouped = grouper.UngroupStaged(ShapeStagingTag.GroupName);
                Assert.True(ungrouped.Succeeded, $"Ungroup refused: {ungrouped.Refusal}");

                for (int index = 0; index < stagedNames.Length; index++)
                {
                    Excel.Shape child = scope.Track(shapes.Item(stagedNames[index]));
                    RectD after = new(
                        Convert.ToDouble(child.Left, CultureInfo.InvariantCulture),
                        Convert.ToDouble(child.Top, CultureInfo.InvariantCulture),
                        Convert.ToDouble(child.Width, CultureInfo.InvariantCulture),
                        Convert.ToDouble(child.Height, CultureInfo.InvariantCulture));
                    Assert.True(
                        Math.Abs(after.Left - before[index].Left) <= hostTolerance
                        && Math.Abs(after.Top - before[index].Top) <= hostTolerance
                        && Math.Abs(after.Width - before[index].Width) <= hostTolerance
                        && Math.Abs(after.Height - before[index].Height) <= hostTolerance,
                        $"Child '{stagedNames[index]}' moved through ungroup: {before[index]} -> {after}");
                }

                // Edit one child and read the new position back: the editable premise,
                // proven rather than assumed.
                Excel.Shape edited = scope.Track(shapes.Item(stagedNames[0]));
                double originalLeft = Convert.ToDouble(edited.Left, CultureInfo.InvariantCulture);

                // The PIA types Shape.Left as float - the same single precision the host
                // stores geometry in - so the edit writes a float and reads it back.
                edited.Left = (float)(originalLeft + 5);
                Assert.Equal(
                    originalLeft + 5,
                    Convert.ToDouble(edited.Left, CultureInfo.InvariantCulture),
                    precision: 2);

                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"Excel build {fixture.Excel.Version}; staged={stagedNames.Length}"
                    + $" groupChildren={items.Count}"
                    + $" envelope={envelope.X:F1},{envelope.Y:F1},{envelope.Width:F1},{envelope.Height:F1}"
                    + $" host={hostEnvelope.X:F1},{hostEnvelope.Y:F1},{hostEnvelope.Width:F1},{hostEnvelope.Height:F1}"));
            }
            finally
            {
                // Cleanup through the staging scheme only: the group first (which would take
                // its children with it if it still existed), then any child already ungrouped.
                _ = stagingWriter.Delete(ShapeStagingTag.GroupPrimitiveId);
                for (int index = staged.StagedPrimitiveIds.Count - 1; index >= 0; index--)
                {
                    _ = stagingWriter.Delete(staged.StagedPrimitiveIds[index]);
                }
            }

            // Zero staged residue; the sentinel and the live chart survive cleanup.
            Assert.Equal(0, CountStaged(scope, sheet));
            Assert.Equal(sheetCountBefore, sheets.Count);
            Assert.Equal(witnessBefore, witness.Value2);
            Assert.Equal("User note", sentinel.Name);
            Assert.Equal(["live:probe-1", "live:probe-2"], liveWriter.ListOwned());

            // The view state the operation ran under is still the user's.
            Assert.Equal(scrollRowBefore, window.ScrollRow);
            Assert.Equal(scrollColumnBefore, window.ScrollColumn);
            Assert.Equal(zoomBefore, window.Zoom, precision: 6);
            Assert.Equal(selectionBefore, SelectionAddress(scope, fixture.Excel));

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

    private static RectD[] CaptureChildEnvelopes(
        OfficeFixture.ComScope scope,
        Excel.Shape group,
        int expectedCount)
    {
        Excel.GroupShapes items = scope.Track(group.GroupItems);
        Assert.Equal(expectedCount, items.Count);

        RectD[] envelopes = new RectD[expectedCount];
        for (int index = 1; index <= expectedCount; index++)
        {
            Excel.Shape child = scope.Track(items.Item(index));
            envelopes[index - 1] = new RectD(
                Convert.ToDouble(child.Left, CultureInfo.InvariantCulture),
                Convert.ToDouble(child.Top, CultureInfo.InvariantCulture),
                Convert.ToDouble(child.Width, CultureInfo.InvariantCulture),
                Convert.ToDouble(child.Height, CultureInfo.InvariantCulture));
        }

        return envelopes;
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
