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

    /// <summary>
    /// A port whose Protect runs the REAL outline-symbol logic, with only the host calls beneath it
    /// substituted, so the diagnostic line and the refusal behaviour around it are observable.
    /// </summary>
    /// <remarks>
    /// <see cref="TestablePort"/> cannot prove any of this: it substitutes
    /// <c>ProtectWorksheet</c>, which is exactly the method under test here.
    /// </remarks>
    private sealed class OutliningPort(object? application, bool contentsProtected, List<string> trace)
        : ExcelWorksheetProtectionPort(application, trace.Add)
    {
        /// <summary>Gets or sets whether the sheet currently reports itself protected.</summary>
        public bool Protected { get; set; } = contentsProtected;

        /// <summary>Gets or sets the host's outline-symbol option.</summary>
        public bool OutlineEnabled { get; set; }

        /// <summary>Gets how many times the host refused the option.</summary>
        public int RefusedSets { get; private set; }

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            new Mock<Excel._Worksheet>().Object;

        internal override bool ReadContentsProtected(Excel._Worksheet sheet) => Protected;

        internal override void ApplyProtect(Excel._Worksheet sheet) => Protected = true;

        internal override bool? ReadOutlineSymbolsEnabled(Excel._Worksheet sheet) => OutlineEnabled;

        internal override bool SetOutlineSymbolsEnabled(Excel._Worksheet sheet)
        {
            // The measured host behaviour: the property is refused while the sheet is protected,
            // which is what turns a "set" into a silent no-op that leaves the option as Protect left it.
            if (Protected)
            {
                RefusedSets++;
                return false;
            }

            OutlineEnabled = true;
            return true;
        }
    }

    /// <summary>
    /// A protect on an unprotected sheet records what it found, what each attempt did, and what the
    /// host then reports — the facts the two live reports are argued from.
    /// </summary>
    [Fact]
    public void Protect_records_the_outline_symbol_state_it_found_and_left()
    {
        List<string> trace = [];
        var port = new OutliningPort(ActiveApplication().Object, contentsProtected: false, trace);

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        string line = Assert.Single(trace);
        Assert.Contains("gantt.protection protect outlining", line, StringComparison.Ordinal);
        Assert.Contains("protectedBefore=False", line, StringComparison.Ordinal);
        Assert.Contains("setBefore=ok", line, StringComparison.Ordinal);
        Assert.Contains("setAfter=refused", line, StringComparison.Ordinal);
        Assert.Contains("enabledAfter=true", line, StringComparison.Ordinal);
    }

    /// <summary>
    /// A protect re-applied to an ALREADY-protected sheet records both attempts as refused, because
    /// that is the state in which the option cannot be repaired and the symbols stay greyed.
    /// </summary>
    /// <remarks>
    /// This is the mechanism behind "full refresh keeps them unlocked but auto refresh locks them":
    /// the path that reaches Protect while the sheet is still protected cannot set the option at all,
    /// and until this line existed neither path said so.
    /// </remarks>
    [Fact]
    public void A_protect_on_an_already_protected_sheet_records_both_attempts_as_refused()
    {
        List<string> trace = [];
        var port = new OutliningPort(ActiveApplication().Object, contentsProtected: true, trace);

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        string line = Assert.Single(trace);
        Assert.Contains("protectedBefore=True", line, StringComparison.Ordinal);
        Assert.Contains("setBefore=refused", line, StringComparison.Ordinal);
        Assert.Contains("setAfter=refused", line, StringComparison.Ordinal);
        Assert.Contains("enabledAfter=false", line, StringComparison.Ordinal);
        Assert.Equal(2, port.RefusedSets);
    }
}
