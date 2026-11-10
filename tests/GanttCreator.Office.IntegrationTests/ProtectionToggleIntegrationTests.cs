using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the two protection toggles (owner ruling 2026-11-09, ADR-0055 clause 7): the
/// CHART toggle governs the owned shapes, the CELL toggle governs everything outside the visible
/// table, the visible table is always editable, and the two switches do not interfere.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why live.</b> What the product promises here is host behaviour: <c>Protect(DrawingObjects,
/// Contents)</c> decides whether a drag lands and whether a write to a locked cell is refused, and no
/// contract test can observe that. The toggles are driven exactly as
/// <c>ToggleChartLockCommand</c> and <c>ToggleCellLockCommand</c> drive them — set the preference,
/// lock or release the owned shapes, then re-protect — so a failure names the host question rather
/// than a pipeline step. The command wiring itself (set rather than flip, one toggle leaving the
/// other alone) is proved without a host in <c>ProtectionToggleCommandTests</c>.
/// </para>
/// <para>
/// Deliberately no <c>RegisterXll</c>: loading the packed add-in starts a live scheduler in the same
/// Excel instance whose automatic refresh would protect the sheet underneath these assertions.
/// </para>
/// </remarks>
public class ProtectionToggleIntegrationTests(ITestOutputHelper output)
{
    private const string ShapeId = "probe:chart";

    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_chart_toggle_governs_the_shapes_and_the_cell_toggle_governs_the_rest_of_the_sheet()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

            using var scope = new OfficeFixture.ComScope();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            // The two toggles plus the ports the commands use. Defaults are the shipped state:
            // chart locked, cells locked.
            var preferences = new ProtectionPreferences();
            ProtectionOwnershipSource ownership = new();
            var writer = new ExcelShapeWriter(
                fixture.Excel,
                new ExcelWorksheetProtectionGuard(fixture.Excel, ownership));
            var port = new ExcelWorksheetProtectionPort(
                fixture.Excel,
                trace: _output.WriteLine,
                preferences: preferences);

            // The owned probe shape the chart toggle acts on. Ownership is what authorises the later
            // Lock/Unlock writes, exactly as the refresh's arming pass establishes it live.
            Assert.True(writer.Create(BarRequest(ShapeId)).Succeeded);
            ownership.Set(true);

            Excel.Shapes shapes = (Excel.Shapes)scope.Track(sheet.Shapes);
            Excel.Shape shape = (Excel.Shape)scope.Track(shapes.Item(ShapeId));
            float originalLeft = shape.Left;

            Excel.ListObject table = (Excel.ListObject)scope.Track(
                sheet.ListObjects[GanttTableSchema.TableName]);
            Excel.Range tableRange = (Excel.Range)scope.Track(table.Range);
            Excel.Range headerRow = (Excel.Range)scope.Track(table.HeaderRowRange);
            Excel.Range headerCell = (Excel.Range)scope.Track(headerRow.Cells[1, 1]);

            // A cell that is unambiguously OUTSIDE the table: two rows below it.
            int outsideRow = headerRow.Row + tableRange.Rows.Count + 2;
            Excel.Range outsideCell = (Excel.Range)scope.Track(sheet.Cells[outsideRow, 1]);

            // ---- CHART LOCKED (the default): the shape cannot be dragged. ----
            preferences.ChartLocked = true;
            Assert.True(writer.LockOwnedShapes().Succeeded);
            Assert.True(port.Protect().Succeeded);

            Assert.True(sheet.ProtectDrawingObjects);
            Assert.True(shape.Locked);
            _ = Assert.Throws<ArgumentException>(() => shape.Left = originalLeft + 20f);
            _output.WriteLine("chart locked: ProtectDrawingObjects=True, shape.Locked=True, drag refused");

            // ---- The fresh table is editable, and the header is not. ----
            // Initialise creates the table with a header and NO body rows, so DataBodyRange is null
            // and the port falls back to the row immediately below the header. That is the blank
            // entry row the user types the first activity into, and unlocking it is the ruling;
            // unlocking the HEADER instead (the previous fallback) would let a user rename a schema
            // column and break the name-keyed column map every adapter builds.
            Assert.True(sheet.ProtectContents);
            Assert.True(Convert.ToBoolean(headerCell.Locked));
            Excel.Range blankEntryCell = (Excel.Range)scope.Track(
                sheet.Cells[headerRow.Row + 1, tableRange.Column]);
            Assert.False(Convert.ToBoolean(blankEntryCell.Locked));
            _output.WriteLine("fresh table: header locked, the blank entry row is editable");

            // Give the table a real body row for the "table stays editable" assertions below.
            Assert.True(port.Unprotect().Succeeded);
            var inserter = new ExcelGanttRowInserter(
                fixture.Excel,
                new ExcelWorksheetProtectionGuard(fixture.Excel, ownership));
            Assert.True(inserter.Insert(GanttEntityType.AsPlannedActivity, () => FixedId()).Succeeded);
            Excel.Range? dataBody = table.DataBodyRange;
            Assert.NotNull(dataBody);
            Excel.Range body = (Excel.Range)scope.Track(dataBody!);
            Excel.Range bodyCell = (Excel.Range)scope.Track(body.Cells[1, 1]);

            // ---- CHART UNLOCKED: the same drag lands. ----
            preferences.ChartLocked = false;
            Assert.True(writer.UnlockShape(ShapeId).Succeeded);
            Assert.True(port.Protect().Succeeded);

            Assert.False(sheet.ProtectDrawingObjects);
            Assert.False(shape.Locked);
            shape.Left = originalLeft + 20f;
            Assert.Equal(originalLeft + 20f, shape.Left, 1);
            _output.WriteLine("chart unlocked: ProtectDrawingObjects=False, shape.Locked=False, drag ok");

            // ---- CELLS LOCKED: outside the table is read-only, the table is not. ----
            preferences.CellsUnlocked = false;
            Assert.True(port.Protect().Succeeded);

            // A UI-only protect exempts the COM write path — ADR-0055's measured table records
            // `Range.Value2` succeeding on a protected sheet — so an automated test cannot observe
            // the refusal a user sees. What the user experiences is `ProtectContents` plus the
            // cell's own `Locked` flag, and that is what these two assertions pin.
            Assert.True(sheet.ProtectContents);
            Assert.True(Convert.ToBoolean(outsideCell.Locked));

            // The owner's permanent rule: the visible table is editable whatever the toggles say.
            Assert.False(Convert.ToBoolean(bodyCell.Locked));
            bodyCell.Value2 = "editable";
            Assert.Equal("editable", bodyCell.Value2);
            _output.WriteLine("cells locked: outside write refused, table body write ok");

            // ---- CELLS UNLOCKED: the rest of the sheet becomes editable. ----
            preferences.CellsUnlocked = true;
            Assert.True(port.Protect().Succeeded);

            Assert.False(sheet.ProtectContents);
            outsideCell.Value2 = "editable now";
            Assert.Equal("editable now", outsideCell.Value2);
            _output.WriteLine("cells unlocked: ProtectContents=False, outside write ok");

            // ---- Independence: locking the chart again leaves the cells unlocked. ----
            preferences.ChartLocked = true;
            Assert.True(writer.LockOwnedShapes().Succeeded);
            Assert.True(port.Protect().Succeeded);

            Assert.True(sheet.ProtectDrawingObjects);
            Assert.False(sheet.ProtectContents);
            Assert.True(shape.Locked);
            _output.WriteLine("both toggles independent: drawings=True contents=False");

            _ = port.Unprotect();
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static GanttRowId FixedId() =>
        GanttRowId.Parse("G-0123456789abcdef0123456789abcdea");

    private static OfficeShapeRequest BarRequest(string primitiveId) => new(
        primitiveId,
        OfficeShapeKind.Rectangle,
        new OfficeShapeGeometry(Bounds: new RectD(200d, 120d, 90d, 12d)),
        ZLayer.ActivityBody,
        FillColour: ColourHex.Parse("#4A90D9"),
        StrokeColour: ColourHex.Parse("#1F4E79"),
        LineWidthPt: 0.75d);
}
