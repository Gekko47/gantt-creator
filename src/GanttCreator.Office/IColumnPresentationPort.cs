using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>Why the managed columns could not be restored to their classification.</summary>
public enum ColumnPresentationRefusalReason
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The Gantt worksheet or its <c>tblGanttData</c> table was missing.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet is protected.</summary>
    TargetProtected = 2,

    /// <summary>
    /// The live table is missing a column that <see cref="GanttTableSchema"/>
    /// declares.
    /// </summary>
    /// <remarks>
    /// Distinct from <see cref="TableMissing"/> because the table WAS found and is
    /// usable; it is incomplete. Restoration is refused rather than continued,
    /// because writing the remaining columns would apply a correct classification to
    /// a table that is already wrong in a way this adapter cannot repair.
    /// </remarks>
    SchemaColumnMissing = 3,
}

/// <summary>The typed result of restoring the managed columns' hidden/locked state.</summary>
/// <param name="ColumnsRestored">How many columns had a hidden or locked flag written.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record ColumnPresentationOutcome(int ColumnsRestored, ColumnPresentationRefusalReason? Refusal)
{
    /// <summary>Gets whether the restore completed.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="restored">How many columns were written.</param>
    /// <returns>The successful outcome.</returns>
    public static ColumnPresentationOutcome Ok(int restored) => new(restored, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static ColumnPresentationOutcome Refused(ColumnPresentationRefusalReason refusal) => new(0, refusal);
}

/// <summary>
/// Restores every managed column of <c>tblGanttData</c> to the code-owned
/// <see cref="GanttColumnAccess"/> classification — engine columns hidden, every
/// non-authoring column's cells locked (R4.7C D1/D2, ADR-0029 D7/D8).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is a port and not a call into the initialiser.</b> R4.7C applies the
/// classification on Initialise only, and its own code comment records that the
/// "un-hide one and refresh restores it" case is a <em>repair</em> belonging to
/// R4.8A's orchestration. The initialiser's routine is private to that adapter and
/// takes a live <c>ListObject</c>, so it cannot be reused here without either a
/// second copy of the rules or a COM type crossing the port. This port states the
/// operation; the adapter reads the same code-owned
/// <see cref="GanttTableSchema.Default"/> the initialiser reads, so there is one
/// classification and not two.
/// </para>
/// <para>
/// <b>Why it is idempotent.</b> A column already carrying the schema's visibility is
/// not rewritten, so a correct sheet produces an empty write set and a Refresh does
/// not mark the workbook dirty — the same discipline the row-height normaliser uses.
/// </para>
/// <para>
/// <b>This port mutates</b>, so the implementation consults the worksheet protection
/// guard first (ADR-0008 D4) and refuses a protected target rather than writing.
/// </para>
/// </remarks>
public interface IColumnPresentationPort
{
    /// <summary>Restores the hidden and locked state of every managed column.</summary>
    /// <param name="repair">
    /// <see langword="true"/> to read the live flags and rewrite any that drifted; <see langword="false"/>
    /// to trust what the initialiser and the last explicit Refresh left behind, making no host call.
    /// </param>
    /// <returns>How many columns were written, or the refusal reason.</returns>
    /// <remarks>
    /// <para>
    /// <b>Why an automatic pass may decline.</b> The flag verification measured 19-45 ms per pass on
    /// the owner's workbook (2026-11-20: `columns=19`, `22.4`, `27.3`, `45.1`), and the only things
    /// that can change the state are a user unhiding an engine column or unlocking its cells. Owner
    /// ruling 2026-11-20: an automatic pass keeps the chart's geometry and content correct and does
    /// not re-assert presentation, because the explicit Refresh repairs it — the same division of
    /// labour the drag ruling already applies to shape geometry.
    /// </para>
    /// <para>
    /// <b>What that costs.</b> A column the user unhid stays visible until the next explicit Refresh.
    /// The chart still renders from the columns it measures, so the result is "largely correct" and
    /// one click repairs it — which is the trade the owner accepted rather than the 19-45 ms.
    /// </para>
    /// </remarks>
    ColumnPresentationOutcome EnsureClassification(bool repair = true);
}
