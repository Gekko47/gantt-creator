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

            protection.Arm(protects: true);
            ownership.Set(arming.IsArmed);
            Assert.True(sheet.ProtectContents);
            Assert.True(sheet.ProtectDrawingObjects);
            _output.WriteLine("armed: contents=True drawings=True");

            // 1. A cell write is exempt, which is what lets a command repair Duration or StyleKey.
            Excel.Range probeCell = (Excel.Range)scope.Track(sheet.Cells[4, 2]);
            probeCell.Value2 = "probe";
            Assert.Equal("probe", probeCell.Value2);
            _output.WriteLine("cell write on the protected sheet: ok");

            // 2. Creating a shape is NOT exempt, and the host's message names no cause.
            COMException createRefusal = Assert.Throws<COMException>(
                () => writer.Create(BarRequest("probe:protected")));
            _output.WriteLine("shape create on the protected sheet: " + createRefusal.Message);

            // 3. Moving a shape is NOT exempt either.
            ArgumentException moveRefusal = Assert.Throws<ArgumentException>(
                () => sheet.Shapes.Item("probe:control").Left = 300f);
            _output.WriteLine("shape move on the protected sheet: " + moveRefusal.Message);

            // 4. Inserting a table row is refused by the add-in's own write path.
            GanttRowInsertOutcome insert = Insert(inserter, '2');
            Assert.False(insert.Succeeded);
            Assert.Equal(GanttRowInsertRefusalReason.RowWriteRefused, insert.Refusal);
            _output.WriteLine("row insert on the protected sheet: refused " + insert.Refusal);

            // 5. Stamping Locked is exempt, which is why the arming sweeps before it protects.
            Assert.True(writer.LockOwnedShapes().Succeeded);
            _output.WriteLine("lock sweep on the protected sheet: ok");

            // 6. The release window: everything the refresh does works again.
            Assert.True(port.Unprotect().Succeeded);
            Assert.True(writer.Create(BarRequest("probe:released")).Succeeded);
            Assert.True(Insert(inserter, '3').Succeeded);
            _output.WriteLine("create and insert after the release: ok");

            // 7. Re-protecting restores the user-facing lock without blocking cell writes.
            Assert.True(port.Protect().Succeeded);
            probeCell.Value2 = "probe2";
            Assert.Equal("probe2", probeCell.Value2);
            _output.WriteLine("cell write after re-protecting: ok");

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
