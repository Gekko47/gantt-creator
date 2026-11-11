using System.Runtime.InteropServices;
using GanttCreator.Core;

// Namespace alias rather than a wholesale using: Microsoft.Office.Interop.Excel declares its own
// Action type, which would collide with System.Action under a plain using.
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IWorksheetProtectionPort"/> over the Excel application object
/// (ADR-0055).
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none.
/// </param>
/// <param name="trace">
/// Optional diagnostic sink for the outline-symbol state a protect finds and leaves, or
/// <see langword="null"/> for no trace. The production sink must not throw.
/// </param>
/// <remarks>
/// <para>
/// <b>COM ownership.</b> Every proxy reached here is an Excel-owned shared root: it is held in a
/// local, used without chained member expressions, and never force-released, following
/// <see cref="ExcelApplicationAdapter"/>'s ownership policy.
/// </para>
/// <para>
/// <b>Why the arguments are spelled out rather than left to their defaults.</b>
/// Microsoft documents <c>DrawingObjects</c> as defaulting to <b>False</b>, so a
/// <c>Protect()</c> call with the arguments omitted leaves every shape fully draggable while
/// <c>ProtectContents</c> reports the sheet as protected. Shape locking therefore depends on
/// passing it explicitly, and that is the one argument this port must never lose.
/// <c>UserInterfaceOnly</c> is equally load-bearing: it is what lets this add-in keep writing
/// after it has protected the sheet, and it is not persisted — after a save and reopen the sheet
/// is fully protected until <c>Protect</c> is called again.
/// </para>
/// <para>
/// <b>No password.</b> ADR-0055 makes this accidental-edit prevention, not a security control: a
/// password this add-in does not store could never be re-entered, so the workbook would be
/// permanently unprotectable by the user. Both directions pass <see cref="Type.Missing"/>.
/// </para>
/// <param name="preferences">
/// The user's two protection toggles, or <see langword="null"/> for
/// <see cref="ProtectionPreferences.Shared"/>. Read-only here: the port applies whatever state the
/// ribbon has set, so a refresh re-applies the user's choice instead of the default.
/// </param>
/// </remarks>
public class ExcelWorksheetProtectionPort(
    object? application,
    Action<string>? trace = null,
    IProtectionPreferenceSource? preferences = null)
    : IWorksheetProtectionPort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly Action<string>? _trace = trace;

    /// <summary>
    /// The user's two protection toggles (owner ruling 2026-11-09). Read, never written: the port
    /// applies whatever the ribbon has set, so a refresh re-applies the user's choice instead of
    /// overwriting it with the default.
    /// </summary>
    /// <remarks>
    /// Defaults to <see cref="ProtectionPreferences.Shared"/> for the same reason the guard defaults
    /// its ownership source: every adapter in this assembly builds its own port, and the toggles are
    /// a property of the session, not of the port.
    /// </remarks>
    private readonly IProtectionPreferenceSource _preferences = preferences ?? ProtectionPreferences.Shared;

    /// <inheritdoc />
    public WorksheetProtectionState? Read()
    {
        Excel._Worksheet? worksheet = TryResolveWorksheet(out _);
        return worksheet is null
            ? null
            : new WorksheetProtectionState(ReadContentsProtected(worksheet), ReadShapesProtected(worksheet));
    }

    /// <inheritdoc />
    public ProtectionWriteOutcome Protect()
    {
        Excel._Worksheet? worksheet = TryResolveWorksheet(out ProtectionWriteRefusal? refusal);
        if (worksheet is null)
        {
            return ProtectionWriteOutcome.Refused(refusal!.Value);
        }

        try
        {
            ProtectWorksheet(worksheet);
            return ProtectionWriteOutcome.Ok();
        }
        catch (Exception exception) when (exception is COMException or InvalidOperationException)
        {
            // A host that refuses protection is a workbook condition, not a fault: the caller
            // reports it and carries on unprotected rather than failing the refresh.
            return ProtectionWriteOutcome.Refused(ProtectionWriteRefusal.HostRejected);
        }
    }

    /// <inheritdoc />
    public ProtectionWriteOutcome Unprotect()
    {
        Excel._Worksheet? worksheet = TryResolveWorksheet(out ProtectionWriteRefusal? refusal);
        if (worksheet is null)
        {
            return ProtectionWriteOutcome.Refused(refusal!.Value);
        }

        try
        {
            UnprotectWorksheet(worksheet);
            return ProtectionWriteOutcome.Ok();
        }
        catch (Exception exception) when (exception is COMException or InvalidOperationException)
        {
            return ProtectionWriteOutcome.Refused(ProtectionWriteRefusal.HostRejected);
        }
    }

    /// <summary>Resolves the worksheet carrying the Gantt table, or reports why it cannot.</summary>
    /// <param name="refusal">Why resolution failed, when it did.</param>
    /// <returns>The worksheet, or <see langword="null"/>.</returns>
    private Excel._Worksheet? TryResolveWorksheet(out ProtectionWriteRefusal? refusal)
    {
        refusal = null;

        Excel.Application? application = _application;
        if (application is null)
        {
            refusal = ProtectionWriteRefusal.NoActiveWorkbook;
            return null;
        }

        Excel.Workbook? workbook = GetActiveWorkbook(application);
        if (workbook is null)
        {
            refusal = ProtectionWriteRefusal.NoActiveWorkbook;
            return null;
        }

        Excel.Sheets? sheets = GetSheets(workbook);
        if (sheets is null)
        {
            refusal = ProtectionWriteRefusal.WorksheetMissing;
            return null;
        }

        Excel._Worksheet? worksheet = FindGanttWorksheet(sheets);
        if (worksheet is null)
        {
            refusal = ProtectionWriteRefusal.WorksheetMissing;
        }

        return worksheet;
    }

    // ---- Test seams (internal virtual, per the ExcelWorkbookInitialiser pattern) ----

    /// <summary>Reads the active workbook. Test seam over <c>Application.ActiveWorkbook</c>.</summary>
    /// <param name="application">The application object.</param>
    /// <returns>The workbook, or <see langword="null"/>.</returns>
    internal virtual Excel.Workbook? GetActiveWorkbook(Excel.Application application) =>
        application.ActiveWorkbook;

    /// <summary>Reads a workbook's sheets. Test seam over <c>Workbook.Sheets</c>.</summary>
    /// <param name="workbook">The workbook.</param>
    /// <returns>The sheet collection, or <see langword="null"/>.</returns>
    internal virtual Excel.Sheets? GetSheets(Excel.Workbook workbook) => workbook.Sheets;

    /// <summary>Reads the sheet count. Test seam over <c>Sheets.Count</c>.</summary>
    /// <param name="sheets">The sheet collection.</param>
    /// <returns>The number of sheets.</returns>
    internal virtual int GetSheetCount(Excel.Sheets sheets) => sheets.Count;

    /// <summary>Reads one sheet by its 1-based index. Test seam over <c>Sheets.Item</c>.</summary>
    /// <param name="sheets">The sheet collection.</param>
    /// <param name="index">The 1-based sheet index.</param>
    /// <returns>The sheet, or <see langword="null"/> when the index does not resolve.</returns>
    internal virtual Excel._Worksheet? GetSheetAt(Excel.Sheets sheets, int index) =>
        sheets[index] as Excel._Worksheet;

    /// <summary>
    /// Returns the worksheet carrying <c>tblGanttData</c>.
    /// </summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <returns>The Gantt worksheet, or <see langword="null"/> when absent.</returns>
    /// <remarks>
    /// Identified by its table rather than by its name: the user may rename the worksheet, and the
    /// table name is the schema's identity (as the shape writer and the table reader both assume).
    /// </remarks>
    internal virtual Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets)
    {
        ArgumentNullException.ThrowIfNull(sheets);

        int count = GetSheetCount(sheets);
        for (int index = 1; index <= count; index++)
        {
            Excel._Worksheet? sheet = GetSheetAt(sheets, index);
            if (sheet is not null && HasGanttTable(sheet))
            {
                return sheet;
            }
        }

        return null;
    }

    /// <summary>Whether a sheet carries <c>tblGanttData</c>. Test seam over the table collection.</summary>
    /// <param name="sheet">The candidate worksheet.</param>
    /// <returns><see langword="true"/> when the sheet carries the Gantt table.</returns>
    internal virtual bool HasGanttTable(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        Excel.ListObjects? objects = sheet.ListObjects;
        if (objects is null)
        {
            return false;
        }

        int count = GetListObjectCount(objects);
        for (int index = 1; index <= count; index++)
        {
            Excel.ListObject? table = GetListObjectAt(objects, index);
            if (table is not null && GanttChartContract.IsDataTableNameExact(table.Name))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>Reads the table count. Test seam over <c>ListObjects.Count</c>.</summary>
    /// <param name="objects">The table collection.</param>
    /// <returns>The number of tables on the sheet.</returns>
    internal virtual int GetListObjectCount(Excel.ListObjects objects) => objects.Count;

    /// <summary>Reads one table by its 1-based index. Test seam over <c>ListObjects.Item</c>.</summary>
    /// <param name="objects">The table collection.</param>
    /// <param name="index">The 1-based table index.</param>
    /// <returns>The table, or <see langword="null"/> when the index does not resolve.</returns>
    internal virtual Excel.ListObject? GetListObjectAt(Excel.ListObjects objects, int index) =>
        objects[index];

    /// <summary>
    /// Unlocks the visible table's data cells, so the user can copy, paste and format them while the
    /// chart stays locked. Test seam over <c>Range.Locked</c>.
    /// </summary>
    /// <param name="sheet">The worksheet carrying the table.</param>
    /// <returns><see langword="true"/> when the cells were unlocked.</returns>
    /// <remarks>
    /// <para>
    /// <b>Owner ruling 2026-11-09:</b> "the visible table area should be fully unlocked at all times
    /// for the user to copy paste change formatting etc". Excel's default is `Locked = true` on every
    /// cell, so `Protect(Contents:=true)` made the whole table read-only and blocked every one of
    /// those operations. The protection's remaining job is the CHART: `DrawingObjects` plus the
    /// `Shape.Locked` stamp is what stops an accidental shape drag, and that is unaffected by this.
    /// </para>
    /// <para>
    /// <b>Data body, not the whole range.</b> Unlocking the header would let a user rename a schema
    /// column, which silently breaks the name-keyed column map every adapter builds. The schedule
    /// data is the area the ruling is about. A table with no body rows yet falls back to the row
    /// immediately below the header: that is the blank entry row the user types the first activity
    /// into, and Excel grows the table into it, so the ruling holds without opening the header.
    /// </para>
    /// <para>
    /// <b>Must run while the sheet is released.</b> <c>Locked</c> is a format write and the host
    /// refuses it on a protected sheet, which is one more reason the protect path cycles protection
    /// first. A refusal here is reported and does not abandon the protect: a sheet protected with a
    /// password we do not have is still worth protecting.
    /// </para>
    /// </remarks>
    internal virtual bool UnlockTableCells(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

#pragma warning disable CA1031
        try
        {
            Excel.Range? range = FindGanttTableRange(sheet);
            if (range is null)
            {
                return false;
            }

            range.Locked = false;
            return true;
        }
        catch (Exception exception) when (exception is COMException or InvalidOperationException)
        {
            _ = exception;
            return false;
        }
#pragma warning restore CA1031
    }

    /// <summary>Resolves the Gantt table's data range, or the row below the header when the body is empty.</summary>
    /// <param name="sheet">The worksheet carrying the table.</param>
    /// <returns>The range to unlock, or <see langword="null"/> when the table is absent.</returns>
    private Excel.Range? FindGanttTableRange(Excel._Worksheet sheet)
    {
        Excel.ListObjects? objects = sheet.ListObjects;
        if (objects is null)
        {
            return null;
        }

        int count = GetListObjectCount(objects);
        for (int index = 1; index <= count; index++)
        {
            Excel.ListObject? table = GetListObjectAt(objects, index);
            if (table is null || !GanttChartContract.IsDataTableNameExact(table.Name))
            {
                continue;
            }

            // A freshly initialised table has a header and no body rows, so DataBodyRange is null.
            // The blank entry row the user types into is the one directly below the header; the
            // header itself must stay locked, because renaming a schema column breaks every
            // name-keyed column map. Measured live: asserting the old `?? table.Range` fallback left
            // the header cell unlocked on a new workbook.
            if (table.DataBodyRange is { } dataBody)
            {
                return dataBody;
            }

            Excel.Range? headerRow = table.HeaderRowRange;
            return headerRow?.Offset[1, 0];
        }

        return null;
    }

    /// <summary>Reads <c>ProtectContents</c>. Test seam over the COM property.</summary>
    /// <param name="sheet">The worksheet.</param>
    /// <returns><see langword="true"/> when locked cells are read-only.</returns>
    internal virtual bool ReadContentsProtected(Excel._Worksheet sheet) => sheet.ProtectContents;

    /// <summary>Reads <c>ProtectDrawingObjects</c>. Test seam over the COM property.</summary>
    /// <param name="sheet">The worksheet.</param>
    /// <returns><see langword="true"/> when <c>Shape.Locked</c> is enforced.</returns>
    internal virtual bool ReadShapesProtected(Excel._Worksheet sheet) => sheet.ProtectDrawingObjects;

    /// <summary>
    /// Applies protection with the three arguments ADR-0055 names. Test seam over
    /// <c>Worksheet.Protect</c>, whose sixteen optional parameters cannot be expressed in a
    /// Moq expression.
    /// </summary>
    /// <param name="sheet">The worksheet to protect.</param>
    /// <remarks>
    /// <b>Outlining is enabled first, on purpose.</b> Microsoft documents
    /// <c>Worksheet.EnableOutlining</c> as "True if outlining symbols are enabled when
    /// user-interface-only protection is turned on", and states that it "is not saved with the
    /// worksheet or session". Without it a protected sheet greys out the native +/- group buttons,
    /// which is exactly the outline the add-in spends a refresh building (ADR-0026 D5) and what the
    /// user reaches for to fold a section away. It is set before <c>Protect</c> because the host
    /// refuses the property once the sheet is protected, and it is set on every protect for the same
    /// "not saved" reason — the window's close is the only protect the product performs.
    /// </remarks>
    internal virtual void ProtectWorksheet(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        // The two toggles, read once per protect so every flag below comes from one consistent view
        // of the user's choice. Chart locked drives DrawingObjects; cells unlocked turns Contents
        // protection off, which is what makes the area outside the visible table editable.
        bool chartLocked = _preferences.ChartLocked;
        bool sheetCellsLocked = !_preferences.CellsUnlocked;

        // The two live reports disagree about which side of Protect the host needs, and a third
        // report says the automatic pass still greys the symbols out while a full refresh keeps
        // them. Rather than guess a fourth ordering, the state is READ on both sides and recorded:
        // the deciding fact is whether the sheet was already protected when this ran, because the
        // host refuses EnableOutlining on a protected sheet and a refused attempt is a silent
        // no-op that leaves the symbols exactly as Protect left them.
        bool protectedBefore = ReadContentsProtected(sheet);
        bool? enabledBefore = ReadOutlineSymbolsEnabled(sheet);

        // <b>A sheet protected by the PREVIOUS session has the option disabled and cannot have it
        // set.</b> `EnableOutlining` is refused while protected and is not saved with the
        // worksheet, so after a save and reopen the sheet is protected with the symbols off, and
        // nothing in the add-in re-applied the option until a protect happened on an UNPROTECTED
        // sheet. Reported 2026-11-09: "first load from save the host +/- collapse expand buttons are
        // locked" until a refresh. Cycling the protection is the only way to set a property the
        // host refuses in the current state: release, set, re-protect — with the same DrawingObjects
        // and Contents arguments, so the protection the user had is the protection they get back.
        bool cycled = false;
        if (protectedBefore)
        {
            try
            {
                UnprotectWorksheet(sheet);
                cycled = true;
            }
            catch (Exception exception) when (exception is COMException or InvalidOperationException)
            {
                // A password we do not have. The sheet stays as it was; protecting below is then a
                // no-op and the caller's guard reports what it can prove, which is the existing
                // behaviour rather than a new failure.
                _ = exception;
            }
        }

        // Owner ruling 2026-11-09: the visible table stays editable (copy/paste/format) while the
        // chart is locked. This must precede Protect — Locked is a format write the host refuses on
        // a protected sheet — and it is deliberately not fatal: a sheet we cannot release simply
        // stays as locked as it was.
        bool cellsUnlocked = UnlockTableCells(sheet);

        bool setBefore = SetOutlineSymbolsEnabled(sheet);

        ApplyProtect(sheet, chartLocked, sheetCellsLocked);

        bool setAfter = SetOutlineSymbolsEnabled(sheet);
        bool? enabledAfter = ReadOutlineSymbolsEnabled(sheet);

        // One line, emitted after the last read, so a live log answers the question the reports
        // keep re-opening: was the sheet already protected (which is what makes a set a no-op),
        // did each attempt land, and what does the host actually report afterwards.
        _trace?.Invoke(
            "gantt.protection protect outlining protectedBefore="
            + protectedBefore
            + " cellsUnlocked="
            + DescribeAttempt(cellsUnlocked)
            + " chartLocked="
            + chartLocked
            + " sheetCellsLocked="
            + sheetCellsLocked
            + " setBefore="
            + DescribeAttempt(setBefore)
            + " setAfter="
            + DescribeAttempt(setAfter)
            + " enabledBefore="
            + Describe(enabledBefore)
            + " enabledAfter="
            + Describe(enabledAfter)
            + " cycled="
            + DescribeAttempt(cycled));

        static string Describe(bool? value)
        {
            return value is { } flag ? (flag ? "true" : "false") : "unreadable";
        }

        static string DescribeAttempt(bool landed)
        {
            return landed ? "ok" : "refused";
        }
    }

    /// <summary>
    /// Applies the protection itself. Test seam over the sixteen-parameter
    /// <c>Worksheet.Protect</c> call, which Moq cannot express.
    /// </summary>
    /// <param name="sheet">The worksheet to protect.</param>
    /// <param name="drawingObjectsProtected">The CHART toggle: whether <c>Shape.Locked</c> is enforced.</param>
    /// <param name="contentsProtected">The CELL toggle: whether cells outside the visible table are read-only.</param>
    /// <remarks>
    /// <para>
    /// The <c>UserInterfaceOnly</c> and "no password" arguments are still proved only by the live
    /// Office gate, because they are constants here. The two flags are not: they come from the user's
    /// toggles, so the seam takes them as parameters and a contract test asserts the mapping from
    /// toggle state to <c>Protect</c> arguments without a host.
    /// </para>
    /// <para>
    /// Extracting the call is also what lets the OUTLINING diagnostic around it be tested, because a
    /// test that substitutes all of <see cref="ProtectWorksheet"/> never reaches the reads.
    /// </para>
    /// </remarks>
    internal virtual void ApplyProtect(
        Excel._Worksheet sheet,
        bool drawingObjectsProtected,
        bool contentsProtected)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        sheet.Protect(
            Type.Missing,   // Password: none, deliberately (ADR-0055).

            // Owner ruling 2026-11-09: the CHART toggle. True is what makes Shape.Locked meaningful
            // and stops a stray drag; False is the user asking to move chart elements.
            drawingObjectsProtected,

            // The CELL toggle. True makes cells OUTSIDE the visible table read-only; the table's own
            // cells are unlocked first, so they stay editable either way.
            contentsProtected,

            Type.Missing,   // Scenarios: host default.
            true,           // UserInterfaceOnly: the add-in may still write after protecting.

            // Owner ruling 2026-11-09: formatting is a SHEET permission, not a per-cell one, so
            // unlocking the table's cells alone would still leave Paste Special / Format Cells
            // refused. These three are what "change formatting" actually needs.
            true,           // AllowFormattingCells
            true,           // AllowFormattingColumns
            true,           // AllowFormattingRows
            Type.Missing,   // AllowInsertingColumns
            Type.Missing,   // AllowInsertingRows
            Type.Missing,   // AllowInsertingHyperlinks
            Type.Missing,   // AllowDeletingColumns
            Type.Missing,   // AllowDeletingRows
            Type.Missing,   // AllowSorting
            Type.Missing,   // AllowFiltering
            Type.Missing);  // AllowUsingPivotTables
    }

    /// <summary>
    /// Leaves the native outline symbols usable while the sheet is protected. Test seam over
    /// <c>Worksheet.EnableOutlining</c>.
    /// </summary>
    /// <param name="sheet">The worksheet about to be protected.</param>
    /// <returns><see langword="true"/> when the host accepted the value.</returns>
    /// <remarks>
    /// A refusal is expected rather than exceptional: Excel refuses the property on a protected
    /// sheet, and a refused attempt is a silent no-op that leaves the option exactly as
    /// <c>Protect</c> left it. Reporting the refusal as a bool rather than letting it escape keeps a
    /// harmless no-op — a second Lock shapes click, or a re-protect of an already-protected sheet —
    /// from becoming a <c>HostRejected</c> refusal that would leave the chart unlocked, and it lets
    /// the caller record which of its two attempts actually landed.
    /// </remarks>
    internal virtual bool SetOutlineSymbolsEnabled(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        try
        {
            sheet.EnableOutlining = true;
            return true;
        }
        catch (Exception exception) when (exception is COMException or InvalidOperationException)
        {
            return false;
        }
    }

    /// <summary>
    /// Reads the host's outline-symbol option so a protect can record what it found, rather than
    /// only what it attempted. Test seam over <c>Worksheet.EnableOutlining</c>.
    /// </summary>
    /// <param name="sheet">The worksheet being protected.</param>
    /// <returns>
    /// The host's value, or <see langword="null"/> when the host will not report it in the sheet's
    /// current protection state — which is itself the answer a diagnostic needs.
    /// </returns>
    internal virtual bool? ReadOutlineSymbolsEnabled(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        try
        {
            return sheet.EnableOutlining;
        }
        catch (Exception exception) when (exception is COMException or InvalidOperationException)
        {
            return null;
        }
    }

    /// <summary>Removes protection. Test seam over <c>Worksheet.Unprotect</c>.</summary>
    /// <param name="sheet">The worksheet to unprotect.</param>
    internal virtual void UnprotectWorksheet(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        sheet.Unprotect(Type.Missing);
    }
}
