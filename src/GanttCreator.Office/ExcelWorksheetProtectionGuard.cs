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
public class ExcelWorksheetProtectionGuard(object? application) : IWorksheetProtectionGuard
{
    private readonly Excel.Application? _application = application as Excel.Application;

    /// <inheritdoc />
    public ProtectionGuardOutcome Query()
    {
        Excel.Application? application = _application;
        if (application is null)
        {
            return ProtectionGuardOutcome.NoActiveWorkbook;
        }

        // One proxy per local: no chained
        // `application.ActiveWorkbook.ActiveSheet` member expression
        // (docs/02-ARCHITECTURE.md COM ownership).
        Excel.Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return ProtectionGuardOutcome.NoActiveWorkbook;
        }

        // Read-only check 1 — workbook structure protection. ADR-0008 D4:
        // the protection guard is the first check in every mutating adapter.
        // Structure protection is checked before the sheet because it is the
        // workbook-wide gate; a structure-protected workbook must refuse even
        // if the active sheet is not content-protected.
        if (IsWorkbookStructureProtected(workbook))
        {
            return ProtectionGuardOutcome.WorkbookStructureProtected;
        }

        // Read-only check 2 — active sheet content protection. If the active
        // sheet cannot be resolved (the workbook has no active sheet), treat
        // that as "not determinable" and refuse rather than assume unprotected.
        Excel._Worksheet? sheet = GetActiveSheet(workbook);
        if (sheet is null)
        {
            return ProtectionGuardOutcome.NoActiveWorkbook;
        }

        // Read-only check 3 — sheet content protection. ProtectContents is a
        // read-only Boolean on _Worksheet; the guard queries it, never writes
        // it.
#pragma warning disable IDE0046 // 'if' statement can be simplified
        if (IsWorksheetProtected(sheet))
        {
            return ProtectionGuardOutcome.SheetProtected;
        }

        return ProtectionGuardOutcome.NotProtected;
    }

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
