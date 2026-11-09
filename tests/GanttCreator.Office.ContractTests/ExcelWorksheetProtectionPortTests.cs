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
    private sealed class OutliningPort(
        object? application,
        bool contentsProtected,
        List<string> trace,
        IProtectionPreferenceSource? preferences = null)
        : ExcelWorksheetProtectionPort(application, trace.Add, preferences)
    {
        /// <summary>Gets or sets whether the sheet currently reports itself protected.</summary>
        public bool Protected { get; set; } = contentsProtected;

        /// <summary>Gets or sets the host's outline-symbol option.</summary>
        public bool OutlineEnabled { get; set; }

        /// <summary>Gets how many times the host refused the option.</summary>
        public int RefusedSets { get; private set; }

        /// <summary>Gets the seam calls in the order the protect path made them.</summary>
        public List<string> Calls { get; } = [];

        /// <summary>Gets or sets whether the host refuses the cell-unlock format write.</summary>
        public bool RefuseUnlockCells { get; set; }

        /// <summary>Records the ruling-required unlock and reports whether it landed.</summary>
        internal override bool UnlockTableCells(Excel._Worksheet sheet)
        {
            Calls.Add("unlock-cells");
            return !RefuseUnlockCells;
        }

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            new Mock<Excel._Worksheet>().Object;

        internal override bool ReadContentsProtected(Excel._Worksheet sheet) => Protected;

        /// <summary>Gets the CHART toggle the protect was applied with.</summary>
        public bool? AppliedDrawingObjects { get; private set; }

        /// <summary>Gets the CELL toggle the protect was applied with.</summary>
        public bool? AppliedContents { get; private set; }

        internal override void ApplyProtect(
            Excel._Worksheet sheet,
            bool drawingObjectsProtected,
            bool contentsProtected)
        {
            Calls.Add("protect");
            AppliedDrawingObjects = drawingObjectsProtected;
            AppliedContents = contentsProtected;
            Protected = true;
        }

        /// <summary>Gets or sets whether the host refuses to release the protection (a password).</summary>
        public bool RefuseUnprotect { get; set; }

        internal override void UnprotectWorksheet(Excel._Worksheet sheet)
        {
            if (RefuseUnprotect)
            {
                // The same refusal shape the adapter already catches: a sheet protected with a
                // password this add-in does not have.
                throw new InvalidOperationException("The host refused to release the protection.");
            }

            Protected = false;
        }

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
    /// A protect on a sheet already protected by the PREVIOUS session cycles the protection so the
    /// outline option can be set, which is what the host +/- buttons need to be usable after a reopen.
    /// </summary>
    /// <remarks>
    /// `EnableOutlining` is refused while protected and is not saved with the worksheet, so a
    /// reopened workbook is protected with the symbols off and nothing in the add-in could repair it.
    /// Reported 2026-11-09: "first load from save the host +/- collapse expand buttons are locked"
    /// until a refresh ran.
    /// </remarks>
    [Fact]
    public void A_protect_on_an_already_protected_sheet_cycles_so_the_option_lands()
    {
        List<string> trace = [];
        var port = new OutliningPort(ActiveApplication().Object, contentsProtected: true, trace);

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        string line = Assert.Single(trace);
        Assert.Contains("protectedBefore=True", line, StringComparison.Ordinal);
        Assert.Contains("cycled=ok", line, StringComparison.Ordinal);
        Assert.Contains("setBefore=ok", line, StringComparison.Ordinal);
        Assert.Contains("enabledAfter=true", line, StringComparison.Ordinal);

        // Exactly one refusal: the after-set, because Protect has re-protected the sheet by then.
        // That is the sibling test's pinned behaviour and is why the before-set is the one that
        // has to land for the option to be repaired at all.
        Assert.Equal(1, port.RefusedSets);
    }

    /// <summary>
    /// When the sheet cannot be released — a password this add-in does not have — the protect still
    /// succeeds, records the refusal, and leaves the option as the host had it.
    /// </summary>
    [Fact]
    public void A_protect_that_cannot_release_the_sheet_records_the_refusal_and_still_protects()
    {
        List<string> trace = [];
        var port = new OutliningPort(ActiveApplication().Object, contentsProtected: true, trace)
        {
            RefuseUnprotect = true,
        };

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        string line = Assert.Single(trace);
        Assert.Contains("cycled=refused", line, StringComparison.Ordinal);
        Assert.Contains("setBefore=refused", line, StringComparison.Ordinal);
        Assert.Contains("enabledAfter=false", line, StringComparison.Ordinal);
    }

    /// <summary>
    /// The visible table's cells are unlocked BEFORE the protect, so the user can copy, paste and
    /// format schedule data while the chart is locked.
    /// </summary>
    /// <remarks>
    /// Owner ruling 2026-11-09: "the visible table area should be fully unlocked at all times for the
    /// user to copy paste change formatting etc". Excel locks every cell by default, so
    /// <c>Protect(Contents:=true)</c> made the table read-only. The order is the assertion: a
    /// <c>Locked</c> write is refused on a protected sheet, so an unlock after the protect could not
    /// land.
    /// </remarks>
    [Fact]
    public void Protect_unlocks_the_table_cells_before_it_protects()
    {
        List<string> trace = [];
        var port = new OutliningPort(ActiveApplication().Object, contentsProtected: false, trace);

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(["unlock-cells", "protect"], port.Calls);
        Assert.Contains("cellsUnlocked=ok", Assert.Single(trace), StringComparison.Ordinal);
    }

    /// <summary>
    /// A refused cell unlock is reported and does not abandon the protect.
    /// </summary>
    /// <remarks>
    /// The state that produces it is a sheet protected with a password this add-in does not have:
    /// the cycle cannot release it, so the format write is refused. Protecting anyway is still the
    /// behaviour the user relies on — the chart stays locked — and the log line says the table did
    /// not get unlocked, so the state is not silently half-applied.
    /// </remarks>
    /// <summary>
    /// With both toggles in their default state, the protect protects the chart and the cell area.
    /// </summary>
    /// <remarks>
    /// Owner ruling 2026-11-09: the ribbon carries one toggle for the chart and one for the sheet's
    /// cells. The default is "both locked" — the chart cannot be dragged and the area outside the
    /// visible table is read-only — which is the state a workbook should be left in.
    /// </remarks>
    [Fact]
    public void Default_toggles_protect_the_chart_and_the_cells_outside_the_table()
    {
        List<string> trace = [];
        var port = new OutliningPort(
            ActiveApplication().Object,
            contentsProtected: false,
            trace,
            new ProtectionPreferences());

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.True(port.AppliedDrawingObjects);
        Assert.True(port.AppliedContents);
        Assert.Contains("chartLocked=True", Assert.Single(trace), StringComparison.Ordinal);
        Assert.Contains("sheetCellsLocked=True", trace[0], StringComparison.Ordinal);
    }

    /// <summary>
    /// Switching both toggles off reaches the host as a protect that locks neither.
    /// </summary>
    /// <remarks>
    /// This is the assertion that the toggles are STATE the port reads rather than constants: before
    /// the ruling the two flags were hard-coded <c>true</c>, so an unlocked chart or an editable sheet
    /// was silently re-locked by the next protect — including the re-protect every release window
    /// performs at the end of a refresh.
    /// </remarks>
    [Fact]
    public void Neither_toggle_protects_neither_the_chart_nor_the_cells()
    {
        List<string> trace = [];
        var port = new OutliningPort(
            ActiveApplication().Object,
            contentsProtected: false,
            trace,
            new ProtectionPreferences { ChartLocked = false, CellsUnlocked = true });

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.False(port.AppliedDrawingObjects);
        Assert.False(port.AppliedContents);
        Assert.Contains("chartLocked=False", Assert.Single(trace), StringComparison.Ordinal);
        Assert.Contains("sheetCellsLocked=False", trace[0], StringComparison.Ordinal);
    }

    [Fact]
    public void A_refused_cell_unlock_still_protects_and_records_the_refusal()
    {
        List<string> trace = [];
        var port = new OutliningPort(ActiveApplication().Object, contentsProtected: true, trace)
        {
            RefuseUnprotect = true,
            RefuseUnlockCells = true,
        };

        ProtectionWriteOutcome outcome = port.Protect();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Contains("protect", port.Calls);
        Assert.Contains("cellsUnlocked=refused", Assert.Single(trace), StringComparison.Ordinal);
    }
}
