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
/// </remarks>
public class ExcelWorksheetProtectionPort(object? application, Action<string>? trace = null)
    : IWorksheetProtectionPort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly Action<string>? _trace = trace;

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
            if (table is not null
                && string.Equals(table.Name, GanttTableSchema.TableName, StringComparison.Ordinal))
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

        // The two live reports disagree about which side of Protect the host needs, and a third
        // report says the automatic pass still greys the symbols out while a full refresh keeps
        // them. Rather than guess a fourth ordering, the state is READ on both sides and recorded:
        // the deciding fact is whether the sheet was already protected when this ran, because the
        // host refuses EnableOutlining on a protected sheet and a refused attempt is a silent
        // no-op that leaves the symbols exactly as Protect left them.
        bool protectedBefore = ReadContentsProtected(sheet);
        bool? enabledBefore = ReadOutlineSymbolsEnabled(sheet);
        bool setBefore = SetOutlineSymbolsEnabled(sheet);

        ApplyProtect(sheet);

        bool setAfter = SetOutlineSymbolsEnabled(sheet);
        bool? enabledAfter = ReadOutlineSymbolsEnabled(sheet);

        // One line, emitted after the last read, so a live log answers the question the reports
        // keep re-opening: was the sheet already protected (which is what makes a set a no-op),
        // did each attempt land, and what does the host actually report afterwards.
        _trace?.Invoke(
            "gantt.protection protect outlining protectedBefore="
            + protectedBefore
            + " setBefore="
            + DescribeAttempt(setBefore)
            + " setAfter="
            + DescribeAttempt(setAfter)
            + " enabledBefore="
            + Describe(enabledBefore)
            + " enabledAfter="
            + Describe(enabledAfter));

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
    /// <remarks>
    /// The arguments this seam is given are the whole ADR-0055 contract and cannot be asserted in a
    /// contract test — <c>DrawingObjects</c>, <c>Contents</c>, <c>UserInterfaceOnly</c>, no password —
    /// so they are proved by the live Office gate. Extracting the call is what lets the OUTLINING
    /// diagnostic around it be tested, because a test that substitutes all of
    /// <see cref="ProtectWorksheet"/> never reaches the reads.
    /// </remarks>
    internal virtual void ApplyProtect(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        sheet.Protect(
            Type.Missing,   // Password: none, deliberately (ADR-0055).
            true,           // DrawingObjects: REQUIRED. Defaults to False, which leaves Shape.Locked inert.
            true,           // Contents: locked cells become read-only, per the schema's IsLocked columns.
            Type.Missing,   // Scenarios: host default.
            true,           // UserInterfaceOnly: the add-in may still write after protecting.
            Type.Missing,   // AllowFormattingCells
            Type.Missing,   // AllowFormattingColumns
            Type.Missing,   // AllowFormattingRows
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
