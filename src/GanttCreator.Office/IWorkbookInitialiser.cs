namespace GanttCreator.Office;

/// <summary>
/// Narrow port over the workbook-initialisation operation: brings the active
/// workbook into the supported state — one visible <c>tblGanttData</c> Gantt
/// worksheet and one <c>xlSheetVeryHidden</c> <c>_GanttCreatorConfig</c>
/// configuration worksheet — or refuses with a typed result.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c> (the CS0433 duplicate-type hazard
/// documented on <see cref="IExcelApplicationAdapter"/>).
/// </para>
/// <para>
/// The port is on the Excel main STA thread (the Ribbon <c>onAction</c>
/// thread). Expected conflicts return typed refusals and mutate nothing;
/// unexpected COM failures propagate to the command boundary, which is the
/// single translation point for unexpected exceptions.
/// </para>
/// </remarks>
public interface IWorkbookInitialiser
{
    /// <summary>
    /// Attempts one initialise operation on the active workbook.
    /// </summary>
    /// <returns>
    /// The typed outcome: the path taken and, on success, the final label of
    /// the Gantt worksheet. On a refusal the workbook is exactly as it was
    /// before the call.
    /// </returns>
    WorkbookInitialiseOutcome Initialise();
}
