namespace GanttCreator.Office;

/// <summary>
/// Restores every managed row of the visible Gantt table to the configured
/// <c>GanttRowHeightPt</c>, so the worksheet and the chart cannot disagree about
/// how tall a row is (R4.7D, ADR-0026 D3/D4).
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive -- no interop type crosses the port --
/// so the AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c>.
/// </para>
/// <para>
/// The port <strong>mutates</strong>: it writes row heights. The implementation
/// therefore consults the worksheet protection guard first (ADR-0008 D4) and
/// refuses a protected target, and normalises on Initialise, on hierarchy
/// mutation, and on Refresh. R4.8A owns the orchestration; this row owns the
/// operation.
/// </para>
/// <para>
/// A row already at the token is not written, so a correctly normalised sheet
/// reports zero writes and the workbook is not marked dirty by a Refresh.
/// </para>
/// </remarks>
public interface IRowHeightNormalisationPort
{
    /// <summary>Normalises the managed row heights of the Gantt table.</summary>
    /// <param name="managedHeightPt">The target height for an ordinary managed row.</param>
    /// <param name="splitterHeightPt">The target height for a <c>Splitter</c> row.</param>
    /// <param name="spacerHeightPt">The target height for a <c>Spacer</c> row.</param>
    /// <returns>How many rows were written, or the refusal reason.</returns>
    RowHeightNormalisationOutcome Normalise(double managedHeightPt, double splitterHeightPt, double spacerHeightPt);
}
