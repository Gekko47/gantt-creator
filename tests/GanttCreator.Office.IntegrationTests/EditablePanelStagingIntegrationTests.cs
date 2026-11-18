using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R6.2 Office gate: a real staged table composition, inspected on the host - shape kinds,
/// texts, bounds, borders, z-order, nothing outside the staging envelope, no cell write
/// and no sheet addition - then cleaned up.
/// </summary>
/// <remarks>
/// The fixture drives the production chain: initialise a real table, write real rows, read
/// and validate them, measure the real grid, project the panel rows, build the panel with
/// <see cref="PanelBuilder"/>, and stage through an <see cref="ExcelShapeWriter"/>
/// constructed with <see cref="ShapeOwnershipScheme.Staging"/>. The composition profile is
/// the panel-only stage R6.2 owns; the full export scene (bars, markers, delineators) and
/// live theme resolution are R6.3's row, which is why the panel theme is stated here
/// rather than resolved from the workbook.
/// </remarks>
public sealed class EditablePanelStagingIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_staged_panel_composition_lands_inside_the_staging_envelope_and_cleans_up()
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

            // Clear any body rows a previous state left, so the fixture owns every row the
            // panel will reproduce.
            while (listRows.Count > 0)
            {
                scope.Track(listRows[1]).Delete();
            }

            Excel.ListRow spanRow = scope.Track(listRows.Add());
            Excel.ListRow milestoneRow = scope.Track(listRows.Add());

            // Two real rows: a span with a description and a milestone with none, so the
            // composed panel contains both a text shape and an omitted blank cell. Writes
            // go through each returned ListRow's own range: on this host the body can be
            // one row wider than ListRows.Count, and DataBodyRange offsets then address a
            // row the fixture did not add.
            SetRowCell(scope, spanRow, "Id", GanttRowId.New().Value);
            SetRowCell(scope, spanRow, "Type", "As-Planned Activity");
            SetRowCell(scope, spanRow, "Description", "Build phase");
            SetRowCell(scope, spanRow, "Start", 45658.0);
            SetRowCell(scope, spanRow, "Finish", 45670.0);
            SetRowCell(scope, milestoneRow, "Id", GanttRowId.New().Value);
            SetRowCell(scope, milestoneRow, "Type", "As-Planned Milestone");
            SetRowCell(scope, milestoneRow, "Start", 45675.0);

            // A blank body row produces no event, so it would leave the panel's row count
            // short of the measured grid's. The fixture composes only the rows it wrote.
            TrimBlankBodyRows(scope, fixture.Excel, listRows);

            GanttTableReadOutcome read = new ExcelGanttTableReader(fixture.Excel).Read();
            Assert.True(read.Succeeded, $"Read refused: {read.Refusal}");

            GanttValidationOutcome validation = GanttRowValidator.Validate(read.Rows);

            // A row that fails validation produces no event, and the panel rejects a row
            // count that disagrees with the measured grid (RowCountMismatch). Naming the
            // issue makes a fixture that writes an invalid row fail with its cause rather
            // than with the panel's generic refusal.
            Assert.True(
                read.Rows.Count == validation.Events.Count,
                $"The fixture rows did not all validate: rows={read.Rows.Count},"
                + $" events={validation.Events.Count}, issues="
                + string.Join(
                    ";",
                    validation.Issues.Select(issue =>
                        $"{issue.RowNumber}:{issue.Field}:{issue.Code}:{issue.Severity}")));
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
            Assert.Equal(validation.Events.Count, grid.RowHeightsPt.Count);

            PanelRowProjectionResult projection = PanelRowProjection.TryProject(
                validation.Events,
                grid,
                GanttDateDisplayFormat.DdMMyyyy);
            Assert.True(projection.Succeeded, $"Projection refused: {projection.Refusal}");

            // R6.2 stages the panel; live theme resolution belongs to R6.3/the export
            // command row, so the theme is stated here with the guide's default tokens.
            PanelTheme theme = new(
                new SceneStyle("DataPanelFill", fillColour: ColourHex.Parse("#FFFFFF")),
                new SceneStyle("DefaultText", fontSizePt: 9.0, textColour: ColourHex.Parse("#000000")),
                new SceneStyle("HeaderFill", fillColour: ColourHex.Parse("#F2F2F2")),
                new SceneStyle("HeaderFontSizePt", fontSizePt: 9.0, bold: false, textColour: ColourHex.Parse("#000000")),
                new SceneStyle("Border", strokeColour: ColourHex.Parse("#BFBFBF"), outlineWidthPt: 0.25));

            // The measured grid's OriginTopPt is the header row's bottom (the same edge
            // the empty-table measurement fact pins), so the panel hangs below it.
            RectD plotBounds = new(grid.PanelRightPt, grid.OriginTopPt, 400, 220);
            PanelBuildOutcome panel = PanelBuilder.TryBuild(
                new PanelBuildRequest(grid, projection.Rows, plotBounds, grid.OriginTopPt, theme));
            Assert.True(panel.Succeeded, $"Panel build refused: {panel.Refusal}");
            PanelBuildResult panelResult = panel.Result!;

            RectD chartBounds = new(
                panelResult.PanelBounds.X - 8,
                panelResult.PanelBounds.Y - 40,
                plotBounds.Right + 8 - (panelResult.PanelBounds.X - 8),
                plotBounds.Bottom + 8 - (panelResult.PanelBounds.Y - 40));
            SceneCreationOutcome scene = GanttScene.TryCreate(
                chartBounds,
                plotBounds,
                panelResult.Primitives,
                []);
            Assert.NotNull(scene.Scene);

            ExportBoundsOutcome resolution = new(
                new ExportBounds(chartBounds),
                null,
                null,
                scene.Scene);

            ExcelShapeWriter stagingWriter = new(
                fixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);

            // The cell-write witness: the composition must not touch a cell, so this exact
            // value must survive the whole operation.
            Excel.Range spanRange = scope.Track(spanRow.Range);
            Excel.Range witness = scope.Track(spanRange.Cells[1, ColumnIndex("Description")]);
            object? witnessBefore = witness.Value2;

            StagingCompositionOutcome staged = EditablePanelStagingComposer.Compose(resolution, stagingWriter);
            try
            {
                Assert.True(staged.Succeeded, $"Composition refused: {staged.Refusal}");
                Assert.True(staged.StagedBounds is not null);

                // No cell write and no sheet addition (work-item D1).
                Assert.Equal(sheetCountBefore, sheets.Count);
                Assert.Equal(witnessBefore, witness.Value2);

                Excel.Shapes shapes = scope.Track(sheet.Shapes);
                List<Excel.Shape> stagedShapes = [];
                List<string> stagedNames = [];
                int shapeCount = shapes.Count;
                for (int index = 1; index <= shapeCount; index++)
                {
                    Excel.Shape shape = scope.Track(shapes.Item(index));
                    string name = shape.Name;
                    if (name.StartsWith(ShapeStagingTag.NamePrefix, StringComparison.Ordinal))
                    {
                        stagedShapes.Add(shape);
                        stagedNames.Add(name);
                    }
                }

                string[] expectedNames =
                [
                    .. staged.StagedPrimitiveIds.Select(id => ShapeStagingTag.NamePrefix + id),
                ];
                Assert.Equal(expectedNames.Length, stagedShapes.Count);
                Assert.Equal(
                    expectedNames.Order(StringComparer.Ordinal).ToArray(),
                    stagedNames.Order(StringComparer.Ordinal).ToArray());

                Dictionary<string, Excel.Shape> stagedByName =
                    stagedShapes.ToDictionary(shape => shape.Name, StringComparer.Ordinal);

                RectD envelope = staged.StagedBounds!.Value;

                // Iterate in the SCENE's creation order, not the Shapes collection's
                // order: creation order is the z-order contract. The host reading is the
                // one PolygonRenderIntegrationTests already pins live - a HIGHER
                // ZOrderPosition is nearer the front - so shapes created back to front
                // must read a strictly increasing sequence.
                int zOrder = int.MinValue;
                foreach (string primitiveId in staged.StagedPrimitiveIds)
                {
                    Excel.Shape shape = stagedByName[ShapeStagingTag.NamePrefix + primitiveId];

                    // The staging identity, proved from the host's own alternative text.
                    Assert.True(
                        ShapeStagingTag.IsStagingTag(shape.AlternativeText),
                        $"A staged shape carries no staging tag: {shape.Name}");
                    Assert.False(
                        ShapeOwnershipTag.IsOwnedTag(shape.AlternativeText),
                        $"A staged shape carries the LIVE ownership tag: {shape.Name}");

                    // Nothing outside the staging origin: every staged shape lies inside
                    // the envelope the outcome reports.
                    double left = Convert.ToDouble(shape.Left, System.Globalization.CultureInfo.InvariantCulture);
                    double top = Convert.ToDouble(shape.Top, System.Globalization.CultureInfo.InvariantCulture);
                    double width = Convert.ToDouble(shape.Width, System.Globalization.CultureInfo.InvariantCulture);
                    double height = Convert.ToDouble(shape.Height, System.Globalization.CultureInfo.InvariantCulture);
                    Assert.True(
                        left >= envelope.X - HostOwnedGeometry.GeometryMatchTolerancePt
                        && top >= envelope.Y - HostOwnedGeometry.GeometryMatchTolerancePt
                        && left + width <= envelope.Right + HostOwnedGeometry.GeometryMatchTolerancePt
                        && top + height <= envelope.Bottom + HostOwnedGeometry.GeometryMatchTolerancePt,
                        $"Staged shape escapes the envelope: {shape.Name}");

                    // Paint order: created back-to-front, so each later shape must sit
                    // nearer the front than the one before it.
                    Assert.True(
                        shape.ZOrderPosition > zOrder,
                        $"Staged z-order is not the scene's order at {shape.Name}.");
                    zOrder = shape.ZOrderPosition;
                }

                // Shape kinds, texts, and borders for the families the panel emits.
                Assert.All(
                    StagedWithRole(stagedShapes, "panel-cell:"),
                    shape => Assert.Equal(MsoAutoShapeType.msoShapeRectangle, shape.AutoShapeType));
                Assert.All(
                    StagedWithRole(stagedShapes, "header-cell:"),
                    shape => Assert.Equal(MsoAutoShapeType.msoShapeRectangle, shape.AutoShapeType));
                Assert.All(
                    StagedWithRole(stagedShapes, "panel-border-v:"),
                    shape => Assert.Equal(MsoShapeType.msoLine, shape.Type));
                Assert.All(
                    StagedWithRole(stagedShapes, "panel-text:"),
                    shape => Assert.Equal(MsoShapeType.msoTextBox, shape.Type));

                string headerTextName = ShapeStagingTag.NamePrefix
                    + ScenePrimitive.CreateId(SceneOwnerId.Chart, "header-text:Type");
                Excel.Shape headerText = stagedShapes.Single(shape => shape.Name == headerTextName);
                Assert.Equal("Type", headerText.TextFrame2.TextRange.Text);

                Excel.Shape border = stagedShapes.First(shape =>
                    shape.Name.Contains("panel-border-v:", StringComparison.Ordinal));
                Assert.Equal(MsoTriState.msoTrue, border.Line.Visible);

                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"Excel build {fixture.Excel.Version}; staged={stagedShapes.Count}"
                    + $" envelope={envelope.X:F1},{envelope.Y:F1},{envelope.Width:F1},{envelope.Height:F1}"));
            }
            finally
            {
                // Cleanup through the staging scheme only, so the live chart (were one
                // present) could not be touched.
                for (int index = staged.StagedPrimitiveIds.Count - 1; index >= 0; index--)
                {
                    _ = stagingWriter.Delete(staged.StagedPrimitiveIds[index]);
                }
            }

            // No temporary shape survives, no sheet was added, and the witness cell still
            // holds what the user typed.
            Excel.Shapes afterShapes = scope.Track(sheet.Shapes);
            int afterCount = afterShapes.Count;
            for (int index = 1; index <= afterCount; index++)
            {
                Excel.Shape shape = scope.Track(afterShapes.Item(index));
                Assert.False(
                    shape.Name.StartsWith(ShapeStagingTag.NamePrefix, StringComparison.Ordinal),
                    $"A staged shape survived cleanup: {shape.Name}");
            }

            Assert.Equal(sheetCountBefore, sheets.Count);
            Assert.Equal(witnessBefore, witness.Value2);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// The staged shapes carrying one role marker, asserting the family is present so an
    /// <c>Assert.All</c> over an empty sequence cannot pass vacuously.
    /// </summary>
    /// <param name="shapes">The staged shapes.</param>
    /// <param name="roleMarker">The role marker the shape name must contain.</param>
    /// <returns>The matching shapes; at least one.</returns>
    private static Excel.Shape[] StagedWithRole(
        IReadOnlyList<Excel.Shape> shapes,
        string roleMarker)
    {
        Excel.Shape[] matching =
        [
            .. shapes.Where(candidate => candidate.Name.Contains(roleMarker, StringComparison.Ordinal)),
        ];
        Assert.True(matching.Length > 0, $"No staged shape carries the role {roleMarker}.");
        return matching;
    }

    private static Excel.ListObject ResolveTable(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        Excel.ListObjects objects = scope.Track(sheet.ListObjects);
        return scope.Track(objects[GanttTableSchema.TableName]);
    }

    /// <summary>
    /// Writes one cell on a fixture row through the row's own range, by schema display
    /// name, so a schema reorder cannot redirect the fixture at a different column.
    /// </summary>
    /// <param name="scope">The COM scope tracking every proxy this helper acquires.</param>
    /// <param name="row">The list row the fixture added.</param>
    /// <param name="headerName">The schema display name of the column.</param>
    /// <param name="value">The value to write.</param>
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

    /// <summary>
    /// Deletes body rows whose every cell is blank, newest first.
    /// </summary>
    /// <remarks>
    /// A fresh table on this host can report a body one row wider than the rows the
    /// fixture added. A blank row yields no validated event, so leaving it would make the
    /// panel's projected row count disagree with the measured grid - a fixture defect, not
    /// a product one. Only blank rows are removed, and only from this fixture's body.
    /// </remarks>
    /// <param name="scope">The COM scope tracking every proxy this helper acquires.</param>
    /// <param name="application">The Excel application, for the non-empty count.</param>
    /// <param name="listRows">The table's rows.</param>
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
