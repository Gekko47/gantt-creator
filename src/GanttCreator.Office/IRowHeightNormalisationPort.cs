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
    /// <param name="headerHeightPt">
    /// The target height for the table's header row, which is the period band's row
    /// (ADR-0030 D5, entity guide §4).
    /// </param>
    /// <param name="reservedRowHeightPt">
    /// The target height for the reserved row above the table, which carries the
    /// table title and the year band (ADR-0030 D4).
    /// </param>
    /// <param name="paddingRowHeightPt">
    /// The target height for the chart's top and bottom padding rows (ADR-0031 D2).
    /// These are empty rows rather than drawn content, so they take their own token
    /// instead of borrowing the year band's: a user who wants a taller margin
    /// should not have to resize a header to get it.
    /// </param>
    /// <param name="anchorRowHeightPt">
    /// The target height for the reserved anchor row below the body (ADR-0038 D1).
    /// It is a THIRD reserved row with its own sub-row token, distinct from the
    /// padding rows: writing the padding height into it would put the plot-spanning
    /// shapes' bottom anchor a whole margin-height down, which is a visible strip of
    /// sheet rather than the sub-row anchor the geometry depends on.
    /// </param>
    /// <param name="typeTexts">
    /// The <c>Type</c> cell text of each body row, in body order, as the table reader
    /// already read them in its single bulk <c>Value2</c> call; or <see langword="null"/>
    /// to have the adapter read the column itself.
    /// </param>
    /// <returns>How many rows were written, or the refusal reason.</returns>
    /// <remarks>
    /// <para>
    /// The layout rows are separate parameters rather than members of the body
    /// plan because they live <em>outside</em> <c>DataBodyRange</c>: the plan's
    /// <c>RowNumber</c> is a body index, and mixing a worksheet row into that list
    /// would make the number mean two different things depending on which list it is
    /// read from. They are written by the same adapter under the same protection
    /// check, and obey the same "already at the token, so no write" rule, so a
    /// normalised sheet still reports zero writes.
    /// </para>
    /// <para>
    /// <b>Why the Type text is passed in.</b> The adapter used to read every body row's
    /// <c>Type</c> cell through COM to decide the row's height policy, one parameterised
    /// read per row, even though the table reader had already read the whole body —
    /// including <c>Type</c> — in one bulk call moments earlier. Passing the texts
    /// removes those reads without changing the decision: it is the same data from the
    /// same refresh.
    /// </para>
    /// </remarks>
    RowHeightNormalisationOutcome Normalise(
        double managedHeightPt,
        double splitterHeightPt,
        double spacerHeightPt,
        double headerHeightPt,
        double reservedRowHeightPt,
        double paddingRowHeightPt,
        double anchorRowHeightPt,
        IReadOnlyList<string?>? typeTexts = null);
}
