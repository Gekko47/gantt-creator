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
    /// Queries the current workbook-protection state of the active worksheet and
    /// active workbook. This remains available for Ribbon/state probes; it is not
    /// sufficient authorization for a mutation against another worksheet.
    /// </summary>
    /// <returns>The typed active-sheet protection outcome.</returns>
    ProtectionGuardOutcome Query();

    /// <summary>
    /// Queries workbook structure protection and the contents protection of the
    /// supplied target worksheet. The target is supplied as an opaque object so
    /// Excel interop does not cross this port.
    /// </summary>
    /// <remarks>
    /// This is the authoritative check for a mutation against a worksheet other
    /// than the active one: <see cref="Query"/> reports the active sheet, whose
    /// protection says nothing about the target. It is deliberately abstract
    /// rather than defaulting to <see cref="Query"/>, because a default would
    /// silently substitute the active sheet for the target and authorise a
    /// mutation the caller meant to refuse.
    /// </remarks>
    /// <param name="target">The resolved target worksheet proxy.</param>
    /// <returns>The typed target protection outcome.</returns>
    ProtectionGuardOutcome QueryTarget(object? target);

}
