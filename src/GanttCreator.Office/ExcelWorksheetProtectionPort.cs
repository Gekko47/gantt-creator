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
public class ExcelWorksheetProtectionPort(object? application) : IWorksheetProtectionPort
{
    private readonly Excel.Application? _application = application as Excel.Application;

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

        // Before AND after, because the two live reports disagree about which one the host needs:
        // a full refresh kept the symbols while an automatic pass greyed them out, and the only
        // difference between those paths is whether a protect is re-applied to an already-protected
        // sheet. The property is refused while protected, so the attempt is best-effort; the
        // unprotected attempt before Protect is the one that matters on a fresh cycle, and the one
        // after covers a host that resets the option when protection is re-applied.
        _ = TrySetOutlineSymbolsEnabled(sheet);

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

        _ = TrySetOutlineSymbolsEnabled(sheet);
    }

    /// <summary>
    /// Sets <c>EnableOutlining</c> if the host allows it in the sheet's current protection state.
    /// </summary>
    /// <param name="sheet">The worksheet being protected.</param>
    /// <returns><see langword="true"/> when the host accepted the value.</returns>
    /// <remarks>
    /// A refusal is expected rather than exceptional: Excel refuses the property on a protected
    /// sheet, and every window close after the first protects an unprotected sheet anyway, so the
    /// next cycle re-applies it. Swallowing the refusal keeps a harmless no-op (a second Lock shapes
    /// click) from becoming a <c>HostRejected</c> refusal that would leave the chart unlocked.
    /// </remarks>
    private bool TrySetOutlineSymbolsEnabled(Excel._Worksheet sheet)
    {
        try
        {
            SetOutlineSymbolsEnabled(sheet);
            return true;
        }
        catch (COMException)
        {
            return false;
        }
    }

    /// <summary>
    /// Leaves the native outline symbols usable while the sheet is protected. Test seam over
    /// <c>Worksheet.EnableOutlining</c>.
    /// </summary>
    /// <param name="sheet">The worksheet about to be protected.</param>
    internal virtual void SetOutlineSymbolsEnabled(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        sheet.EnableOutlining = true;
    }

    /// <summary>Removes protection. Test seam over <c>Worksheet.Unprotect</c>.</summary>
    /// <param name="sheet">The worksheet to unprotect.</param>
    internal virtual void UnprotectWorksheet(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        sheet.Unprotect(Type.Missing);
    }
}
