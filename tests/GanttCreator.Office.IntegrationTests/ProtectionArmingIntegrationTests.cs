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

    // Deliberately no RegisterXll in this class: loading the packed add-in starts a live
    // scheduler in the same Excel instance, whose automatic refresh now ends by PROTECTING the
    // sheet (ADR-0055). That raced these two tests -- one run failed at Assert.False after
    // arming because the add-in had protected the sheet between the call and the assertion.
    // These tests drive the ports directly, so the add-in adds nothing but a competitor.
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Protection_exempts_cell_writes_but_not_shapes_or_row_inserts()
    {
        var fixture = new OfficeFixture();
        fixture.SuppressLeakSignal = true;
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

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
            // Deliberately no name-based shape move here: the host is shared between tests, and a
            // lookup by name that lands on the wrong sheet reports the protection message, which
            // would read as a protection finding. The window section below moves a shape by the
            // proxy it created, which is unambiguous.
            Assert.True(writer.Create(BarRequest("probe:control")).Succeeded);
            Assert.True(Insert(inserter, '1').Succeeded);

            // One proxy, resolved once and reused, so every move below is on the shape this test
            // created rather than on a name looked up again in a shared host.
            Excel.Shape control = (Excel.Shape)scope.Track(sheet.Shapes.Item("probe:control"));
            control.Left = 250f;

            // Positive control for the outlining assertion below: the host does not report the
            // property as set until this add-in sets it.
            Assert.False(sheet.EnableOutlining);
            _output.WriteLine("controls: create, insert and move succeed unprotected");

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
                control.Left = 300f;
                Assert.True(Insert(inserter, '2').Succeeded);
                Assert.True(writer.LockOwnedShapes().Succeeded);
                _output.WriteLine("create, move, insert and lock all succeed inside the window");
            }

            // Closing the window restores the user-facing lock.
            Assert.True(sheet.ProtectContents);
            Assert.True(sheet.ProtectDrawingObjects);

            // The native +/- group buttons stay usable, which is the whole point of the outline the
            // refresh builds. Without EnableOutlining the host greys them out on a protected sheet,
            // and the property is "not saved with the worksheet or session", so every protect has to
            // re-apply it. Pinned here because the contract tests substitute ProtectWorksheet.
            Assert.True(sheet.EnableOutlining);
            _output.WriteLine("window closed: contents=True drawings=True outlining=True");

            // What the lock actually does, and the one write it still allows.
            COMException createRefusal = Assert.Throws<COMException>(
                () => writer.Create(BarRequest("probe:protected")));
            _output.WriteLine("shape create on the protected sheet: " + createRefusal.Message);

            ArgumentException moveRefusal = Assert.Throws<ArgumentException>(
                () => control.Left = 320f);
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

    /// <summary>
    /// The state a reopen leaves behind — a sheet protected by an earlier session, with no
    /// <c>UserInterfaceOnly</c> in force — must be released by arming and re-locked by the window,
    /// twice over, because that is the state every existing customer workbook is in.
    /// </summary>
    /// <remarks>
    /// This is the scenario a refresh and every row command face on the first click after a reopen,
    /// so it is worth its own live test rather than being folded into the fresh-workbook one.
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_sheet_protected_by_an_earlier_session_is_released_by_arming_and_re_locked_by_the_window()
    {
        var fixture = new OfficeFixture();
        fixture.SuppressLeakSignal = true;
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

            using var scope = new OfficeFixture.ComScope();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            // Exactly what a reopen leaves: protected, three flags as an earlier session applied them,
            // and no UserInterfaceOnly because Excel does not persist it.
            sheet.Protect(Type.Missing, true, true, Type.Missing, false);
            Assert.True(sheet.ProtectContents);

            ProtectionOwnershipSource ownership = new();
            ProtectionArmingState arming = new();
            var writer = new ExcelShapeWriter(
                fixture.Excel,
                new ExcelWorksheetProtectionGuard(fixture.Excel, ownership));
            var port = new ExcelWorksheetProtectionPort(fixture.Excel);
            var protection = new WorksheetProtectionArming(port, writer, arming, _output.WriteLine);
            var window = new WorksheetProtectionWindow(port, arming, _output.WriteLine);
            var inserter = new ExcelGanttRowInserter(
                fixture.Excel,
                new ExcelWorksheetProtectionGuard(fixture.Excel, ownership));

            protection.Arm(protects: true);
            ownership.Set(arming.IsArmed);
            Assert.True(arming.IsArmed);
            Assert.False(sheet.ProtectContents);
            _output.WriteLine("arming released a sheet protected by an earlier session");

            using (window.Open())
            {
                Assert.True(writer.Create(BarRequest("probe:reopened")).Succeeded);
                Assert.True(Insert(inserter, '4').Succeeded);
            }

            Assert.True(sheet.ProtectContents);
            Assert.True(sheet.ProtectDrawingObjects);
            _output.WriteLine("the window wrote and re-locked");

            // The second write phase, which is every refresh after the first.
            using (window.Open())
            {
                Assert.True(writer.Create(BarRequest("probe:reopened2")).Succeeded);
            }

            Assert.True(sheet.ProtectContents);
            _output.WriteLine("the second window released and re-locked");

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
