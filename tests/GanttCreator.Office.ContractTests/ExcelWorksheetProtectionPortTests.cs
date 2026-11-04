using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the ADR-0055 protection port. No live Excel: every COM call is a
/// <c>internal virtual</c> seam, so the observable operation is what is asserted.
/// </summary>
/// <remarks>
/// The arguments <c>Protect</c> is given cannot be asserted here: they live inside the one seam
/// that the real host call occupies, and the seam exists precisely so the test does not have to
/// model a sixteen-parameter COM method. That contract — <c>DrawingObjects</c>,
/// <c>Contents</c>, <c>UserInterfaceOnly</c>, no password — is proved by the live Office gate,
/// which is also the only place that can prove the shapes really stop being draggable.
/// </remarks>
public class ExcelWorksheetProtectionPortTests
{
    /// <summary>
    /// The port with its COM seams substituted, so each operation is observable without a host.
    /// </summary>
    private sealed class TestablePort(
        object? application,
        bool ganttWorksheetPresent,
        bool contentsProtected = false,
        bool shapesProtected = false)
        : ExcelWorksheetProtectionPort(application)
    {
        /// <summary>Gets whether the protect seam was reached.</summary>
        public bool ProtectCalled { get; private set; }

        /// <summary>Gets whether the unprotect seam was reached.</summary>
        public bool UnprotectCalled { get; private set; }

        /// <summary>Whether the host raises on protect, as a refusing host would.</summary>
        public bool RefuseProtect { get; set; }

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            ganttWorksheetPresent ? new Mock<Excel._Worksheet>().Object : null;

        internal override bool ReadContentsProtected(Excel._Worksheet sheet) => contentsProtected;

        internal override bool ReadShapesProtected(Excel._Worksheet sheet) => shapesProtected;

        internal override void ProtectWorksheet(Excel._Worksheet sheet)
        {
            if (RefuseProtect)
            {
                // An InvalidOperationException rather than a COMException, because CA2201 reserves
                // the runtime's own exception types and this codebase's fakes model a refused host
                // this way consistently; the adapter catches both, since the interop boundary can
                // surface either.
                throw new InvalidOperationException("The host refused protection.");
            }

            ProtectCalled = true;
        }

        internal override void UnprotectWorksheet(Excel._Worksheet sheet) => UnprotectCalled = true;
    }

    private static Mock<Excel.Application> ActiveApplication()
    {
        var workbook = new Mock<Excel.Workbook>();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        var application = new Mock<Excel.Application>();
        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }

    [Fact]
    public void Protect_reaches_the_host_once()
    {
        var port = new TestablePort(ActiveApplication().Object, ganttWorksheetPresent: true);

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.True(port.ProtectCalled);
        Assert.False(port.UnprotectCalled);
    }

    [Fact]
    public void Unprotect_reaches_the_host_once()
    {
        var port = new TestablePort(
            ActiveApplication().Object, ganttWorksheetPresent: true, contentsProtected: true, shapesProtected: true);

        ProtectionWriteOutcome outcome = port.Unprotect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.True(port.UnprotectCalled);
        Assert.False(port.ProtectCalled);
    }

    /// <summary>
    /// The two flags are reported separately, because a sheet protected WITHOUT
    /// <c>DrawingObjects</c> reports itself protected while every shape stays draggable — the
    /// exact state ADR-0055 must be able to tell apart from a real shape lock.
    /// </summary>
    [Fact]
    public void Read_reports_contents_and_shapes_separately()
    {
        WorksheetProtectionState? contentsOnly = new TestablePort(
            ActiveApplication().Object,
            ganttWorksheetPresent: true,
            contentsProtected: true,
            shapesProtected: false).Read();

        WorksheetProtectionState? both = new TestablePort(
            ActiveApplication().Object,
            ganttWorksheetPresent: true,
            contentsProtected: true,
            shapesProtected: true).Read();

        Assert.NotNull(contentsOnly);
        Assert.True(contentsOnly.ContentsProtected);
        Assert.False(contentsOnly.ShapesProtected);

        Assert.NotNull(both);
        Assert.True(both.ContentsProtected);
        Assert.True(both.ShapesProtected);
    }

    /// <summary>Positive test for the missing-worksheet path: no Gantt table, no state.</summary>
    [Fact]
    public void Read_returns_null_when_no_gantt_worksheet_is_present()
    {
        var port = new TestablePort(ActiveApplication().Object, ganttWorksheetPresent: false);

        Assert.Null(port.Read());
    }

    /// <summary>Positive test for the no-workbook refusal.</summary>
    [Fact]
    public void Protect_refuses_when_there_is_no_active_workbook()
    {
        var port = new TestablePort(application: null, ganttWorksheetPresent: true);

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.Equal(ProtectionWriteRefusal.NoActiveWorkbook, outcome.Refusal);
        Assert.False(port.ProtectCalled);
    }

    /// <summary>Positive test for the missing-worksheet refusal on the write path.</summary>
    [Fact]
    public void Protect_refuses_when_no_gantt_worksheet_is_present()
    {
        var port = new TestablePort(ActiveApplication().Object, ganttWorksheetPresent: false);

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.Equal(ProtectionWriteRefusal.WorksheetMissing, outcome.Refusal);
        Assert.False(port.ProtectCalled);
    }

    /// <summary>
    /// Positive test for the host-refusal path: a host that raises is reported, not thrown, so a
    /// refresh can continue unprotected instead of failing the whole pass.
    /// </summary>
    [Fact]
    public void A_host_that_refuses_protection_is_reported_as_HostRejected()
    {
        var port = new TestablePort(ActiveApplication().Object, ganttWorksheetPresent: true)
        {
            RefuseProtect = true,
        };

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.Equal(ProtectionWriteRefusal.HostRejected, outcome.Refusal);
        Assert.False(port.ProtectCalled);
    }
}
