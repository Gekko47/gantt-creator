// Namespace alias rather than a wholesale using: Microsoft.Office.Interop.Excel
// declares its own Action type, which would collide with System.Action under a
// plain `using Microsoft.Office.Interop.Excel;`.
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IWorksheetProtectionGuard"/> over the Excel application
/// object. Queries the active worksheet's <c>ProtectContents</c> and the
/// active workbook's <c>ProtectStructure</c> and returns a typed outcome;
/// read-only by contract — never mutates the workbook.
/// </summary>
/// <param name="application">
/// The Excel application object (for example
/// <c>ExcelDnaUtil.Application</c>), or <see langword="null"/> when the host
/// supplied none. A foreign object fails the interface cast and degrades to the
/// <c>NoActiveWorkbook</c> outcome.
/// </param>
/// <param name="ownershipSource">
/// Supplies whether the protection on the Gantt worksheet is the add-in's own (ADR-0055), or
/// <see langword="null"/> when no ownership is configured — in which case a protected sheet
/// refuses exactly as it always has.
/// </param>
/// <remarks>
/// <para>
/// COM ownership: the <c>Application</c>, <c>Workbook</c>, and
/// <c>Worksheet</c> proxies reached here are Excel-owned shared roots. This
/// adapter takes no ownership of them, never calls
/// <c>FinalReleaseComObject</c>, and force-releases nothing (the ownership
/// policy of <see cref="ExcelApplicationAdapter"/>). Every proxy is held in a
/// local and used without chained member expressions; the worksheet is reached
/// through <see cref="GetActiveSheet(Excel.Workbook)"/> rather than a chained
/// <c>application.ActiveWorkbook.ActiveSheet</c> expression.
/// </para>
/// <para>
/// Probe properties (R2.7a evidence ledger): <c>Worksheet.ProtectContents</c>
/// (<c>System.Boolean</c>, read-only) and <c>Workbook.ProtectStructure</c>
/// (<c>System.Boolean</c>, read-only) were verified against the installed
/// <c>Microsoft.Office.Interop.Excel</c> 16.0.0 PIA and the Microsoft Learn
/// definition before this guard was written. Both are read-only; the guard
/// never writes protection state.
/// </para>
/// <para>
/// The <c>internal virtual</c> accessors isolate the COM parameterised
/// properties (<c>Workbook.ActiveSheet</c>, <c>Sheets.Item</c>) and the
/// <c>Worksheet</c>/<c>Workbook</c> interface casts so contract tests can
/// substitute them without a live Excel host (CS0855); the real behaviour is
/// exercised by the tagged live-Office integration test.
/// </para>
/// </remarks>
public class ExcelWorksheetProtectionGuard(
    object? application,
    IProtectionOwnershipSource? ownershipSource = null) : IWorksheetProtectionGuard
{
    private readonly Excel.Application? _application = application as Excel.Application;

    /// <summary>
    /// The ownership fact, or <see langword="null"/> when none is configured.
    /// </summary>
    /// <remarks>
    /// Read as a field, never fetched: ADR-0055 requires the ownership decision to add no host
    /// call to the refresh, so the answer is handed in once and read here as many times as the
    /// guard is asked.
    /// </remarks>
    private readonly IProtectionOwnershipSource? _ownershipSource = ownershipSource;

    /// <inheritdoc />
    public ProtectionGuardOutcome Query()
    {
        Excel.Application? application = _application;
        if (application is null)
        {
            return ProtectionGuardOutcome.NoActiveWorkbook;
        }

        Excel.Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return ProtectionGuardOutcome.NoActiveWorkbook;
        }

        if (IsWorkbookStructureProtected(workbook))
        {
            return ProtectionGuardOutcome.WorkbookStructureProtected;
        }

        Excel._Worksheet? sheet = GetActiveSheet(workbook);
        return sheet is null
            ? ProtectionGuardOutcome.NoActiveWorkbook
            : ProtectedOutcome(sheet);
    }

    /// <inheritdoc />
    public ProtectionGuardOutcome QueryTarget(object? target)
    {
        Excel.Application? application = _application;
        if (application is null)
        {
            return ProtectionGuardOutcome.NoActiveWorkbook;
        }

        Excel.Workbook? workbook = application.ActiveWorkbook;
        return workbook is null
            ? ProtectionGuardOutcome.NoActiveWorkbook
            : IsWorkbookStructureProtected(workbook)
                ? ProtectionGuardOutcome.WorkbookStructureProtected
                : target is Excel._Worksheet sheet
                    ? ProtectedOutcome(sheet)
                    : ProtectionGuardOutcome.NoActiveWorkbook;
    }

    /// <summary>
    /// Maps a worksheet's protection state to the guard's answer, honouring add-in ownership.
    /// </summary>
    /// <param name="worksheet">The worksheet the caller intends to write.</param>
    /// <returns>
    /// <see cref="ProtectionGuardOutcome.NotProtected"/> when the sheet is open or the add-in owns
    /// the protection, otherwise <see cref="ProtectionGuardOutcome.SheetProtected"/>.
    /// </returns>
    /// <remarks>
    /// <para>
    /// <b>Ownership (ADR-0055).</b> The add-in protects the visible Gantt worksheet so its shapes
    /// can be locked, which would otherwise make the sheet unwritable to the very code that
    /// protects it. When ownership is recorded, a protected sheet is authorised; when it is not,
    /// every protected sheet refuses exactly as before.
    /// </para>
    /// <para>
    /// <b>Why the flag names no worksheet.</b> It cannot: COM proxies for the same sheet are not
    /// reliably reference-equal, and asking "is this the sheet we protected" would cost a table
    /// walk on every call — the extra step ADR-0055 forbids. The invariant that makes the flag
    /// sufficient is that the add-in protects exactly one worksheet, the visible Gantt sheet, and
    /// every guard call site acts on that sheet or on the VeryHidden configuration sheet, which
    /// the add-in never protects.
    /// </para>
    /// </remarks>
    private ProtectionGuardOutcome ProtectedOutcome(Excel._Worksheet worksheet) =>
        !IsWorksheetProtected(worksheet) || _ownershipSource?.AddInOwnsProtection == true
            ? ProtectionGuardOutcome.NotProtected
            : ProtectionGuardOutcome.SheetProtected;

    /// <summary>
    /// Returns the active sheet of the workbook. Test seam over the COM
    /// parameterised <c>Workbook.ActiveSheet</c> property (CS0855).
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <returns>
    /// The active <c>Worksheet</c>, or <see langword="null"/> when the workbook
    /// has no active sheet.
    /// </returns>
    internal virtual Excel._Worksheet? GetActiveSheet(Excel.Workbook workbook) =>
        workbook.ActiveSheet as Excel._Worksheet;

    /// <summary>
    /// Returns <c>true</c> when the worksheet's contents are protected. Test
    /// seam over the <c>ProtectContents</c> read-only property.
    /// </summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <returns>
    /// <c>true</c> when <c>ProtectContents</c> is <c>true</c>.
    /// </returns>
    internal virtual bool IsWorksheetProtected(Excel._Worksheet worksheet) =>
        worksheet.ProtectContents;

    /// <summary>
    /// Returns <c>true</c> when the workbook structure is protected. Test seam
    /// over the <c>ProtectStructure</c> read-only property.
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <returns>
    /// <c>true</c> when <c>ProtectStructure</c> is <c>true</c>.
    /// </returns>
    internal virtual bool IsWorkbookStructureProtected(Excel.Workbook workbook) =>
        workbook.ProtectStructure;
}
