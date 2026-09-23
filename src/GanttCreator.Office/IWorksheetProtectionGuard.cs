namespace GanttCreator.Office;

/// <summary>
/// A read-only port over the workbook-protection guard for destructive
/// commands. The guard is the first check in every mutating adapter per
/// ADR-0008 D4; it never mutates the workbook.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c> (the CS0433 duplicate-type hazard
/// documented on <see cref="IExcelApplicationAdapter"/>). Office-mutating
/// adapters depend on the port, not on the PIA.
/// </para>
/// <para>
/// The guard is read-only by contract: calling it must not change workbook
/// state. Every mutating adapter queries this port first and refuses on any
/// non-<c>NotProtected</c> outcome before it touches COM.
/// </para>
/// </remarks>
public interface IWorksheetProtectionGuard
{
    /// <summary>
    /// Queries the current workbook-protection state of the active
    /// worksheet's contents and the active workbook's structure. Read-only by
    /// contract — never mutates the workbook.
    /// </summary>
    /// <returns>
    /// The typed outcome. <see cref="ProtectionGuardOutcome.NotProtected"/>
    /// means a mutating command may proceed; any other value means the guard
    /// refuses and the command must not mutate.
    /// </returns>
    ProtectionGuardOutcome Query();
}
