using System.Runtime.InteropServices;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for ADR-0055: what this host actually allows once the Gantt sheet is protected
/// with <c>DrawingObjects</c>, <c>Contents</c> and <c>UserInterfaceOnly</c>.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this test exists and what it changed.</b> Every other integration test initialises a
/// workbook and never arms protection, so the suite could not see that
/// <c>UserInterfaceOnly</c> exempts CELL writes but not shape or row operations. Measured live on
/// 2026-11-04: a cell write lands on a protected sheet, while <c>Shapes.AddShape</c> and
/// <c>Shape.Left</c> raise "The specified value is out of range" and a table-row insert is refused
/// outright — and every one of them succeeds again the moment the sheet is unprotected. The
/// assertions below pin the HOST's behaviour rather than the product's, so the day Excel changes it
/// this test says so instead of the add-in breaking in front of a user.
/// </para>
/// <para>
/// The arming is driven exactly as <c>GanttRefreshOrchestrator</c> drives it — the same
/// <see cref="WorksheetProtectionArming"/>, the same order — so a failure names the protection
/// question rather than a pipeline step.
/// </para>
/// </remarks>
public class ProtectionArmingIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Protection_exempts_cell_writes_but_not_shapes_or_row_inserts()
    {
        var fixture = new OfficeFixture();
        fixture.SuppressLeakSignal = true;
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            using var scope = new OfficeFixture.ComScope();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            ProtectionOwnershipSource ownership = new();
            ProtectionArmingState arming = new();
            var writer = new ExcelShapeWriter(
                fixture.Excel,
                new ExcelWorksheetProtectionGuard(fixture.Excel, ownership));
            var port = new ExcelWorksheetProtectionPort(fixture.Excel);
            var protection = new WorksheetProtectionArming(port, writer, arming, _output.WriteLine);
            var inserter = new ExcelGanttRowInserter(
                fixture.Excel,
                new ExcelWorksheetProtectionGuard(fixture.Excel, ownership));

            // Positive controls on the UNPROTECTED sheet, so a refusal below cannot be the probe.
            Assert.True(writer.Create(BarRequest("probe:control")).Succeeded);
            Assert.True(Insert(inserter, '1').Succeeded);
            sheet.Shapes.Item("probe:control").Left = 250f;
            _output.WriteLine("controls: create, insert and move all succeed unprotected");

            // Arming ends with the sheet RELEASED: a protected sheet refuses shape creation (measured
            // below), so protecting during arming would break the write it exists to enable. The
            // window re-applies it when the write phase ends.
            protection.Arm(protects: true);
            ownership.Set(arming.IsArmed);
            Assert.True(arming.IsArmed);
            Assert.False(sheet.ProtectContents);
            _output.WriteLine("armed: sheet released for the write phase, arming recorded");

            Excel.Range probeCell = (Excel.Range)scope.Track(sheet.Cells[4, 2]);

            // The refresh's write phase, inside the window.
            var window = new WorksheetProtectionWindow(port, arming, _output.WriteLine);
            using (window.Open())
            {
                Assert.True(writer.Create(BarRequest("probe:live")).Succeeded);
                sheet.Shapes.Item("probe:control").Left = 300f;
                Assert.True(Insert(inserter, '2').Succeeded);
                Assert.True(writer.LockOwnedShapes().Succeeded);
                _output.WriteLine("create, move, insert and lock all succeed inside the window");
            }

            // Closing the window restores the user-facing lock.
            Assert.True(sheet.ProtectContents);
            Assert.True(sheet.ProtectDrawingObjects);
            _output.WriteLine("window closed: contents=True drawings=True");

            // What the lock actually does, and the one write it still allows.
            COMException createRefusal = Assert.Throws<COMException>(
                () => writer.Create(BarRequest("probe:protected")));
            _output.WriteLine("shape create on the protected sheet: " + createRefusal.Message);

            ArgumentException moveRefusal = Assert.Throws<ArgumentException>(
                () => sheet.Shapes.Item("probe:control").Left = 320f);
            _output.WriteLine("shape move on the protected sheet: " + moveRefusal.Message);

            GanttRowInsertOutcome insert = Insert(inserter, '3');
            Assert.False(insert.Succeeded);
            Assert.Equal(GanttRowInsertRefusalReason.RowWriteRefused, insert.Refusal);
            _output.WriteLine("row insert on the protected sheet: refused " + insert.Refusal);

            probeCell.Value2 = "probe2";
            Assert.Equal("probe2", probeCell.Value2);
            _output.WriteLine("cell write on the protected sheet: ok");

            // A second window round-trips again, which is what every refresh relies on.
            using (window.Open())
            {
                Assert.True(writer.Create(BarRequest("probe:live2")).Succeeded);
            }

            Assert.True(sheet.ProtectContents);
            _output.WriteLine("second window: wrote and re-protected");

            _ = port.Unprotect();
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static GanttRowInsertOutcome Insert(ExcelGanttRowInserter inserter, char suffix) =>
        inserter.Insert(
            GanttEntityType.AsPlannedActivity,
            () => GanttRowId.Parse("G-0123456789abcdef0123456789abcde" + suffix));

    private static OfficeShapeRequest BarRequest(string primitiveId) => new(
        primitiveId,
        OfficeShapeKind.Rectangle,
        new OfficeShapeGeometry(Bounds: new RectD(200d, 120d, 90d, 12d)),
        ZLayer.ActivityBody,
        FillColour: ColourHex.Parse("#4A90D9"),
        StrokeColour: ColourHex.Parse("#1F4E79"),
        LineWidthPt: 0.75d);
}
