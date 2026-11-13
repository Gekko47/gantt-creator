namespace GanttCreator.Office;

/// <summary>Why a column-visibility operation could not run (R5.18h).</summary>
public enum ColumnVisibilityRefusalReason
{
    /// <summary>The application object or the active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The Gantt worksheet or its <c>tblGanttData</c> table was missing.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet is protected, so the host would refuse the write.</summary>
    TargetProtected = 2,

    /// <summary>The set named a column the table does not carry, or the dropdown does not offer.</summary>
    UnknownColumn = 3,

    /// <summary>The host refused the write.</summary>
    WriteFailed = 4,
}

/// <summary>The typed result of reading or applying the column-visibility set.</summary>
/// <param name="Hidden">
/// The hidden set as the worksheet reports it, or <see langword="null"/> on refusal.
/// </param>
/// <param name="ColumnsWritten">How many whole columns this call actually wrote.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record ColumnVisibilityPortOutcome(
    IReadOnlySet<string>? Hidden,
    int ColumnsWritten,
    ColumnVisibilityRefusalReason? Refusal)
{
    /// <summary>Gets whether the operation completed.</summary>
    public bool Succeeded => Refusal is null && Hidden is not null;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="hidden">The resulting hidden set.</param>
    /// <param name="columnsWritten">How many columns were written.</param>
    /// <returns>The successful outcome.</returns>
    public static ColumnVisibilityPortOutcome Ok(IReadOnlySet<string> hidden, int columnsWritten = 0) =>
        new(hidden, columnsWritten, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static ColumnVisibilityPortOutcome Refused(ColumnVisibilityRefusalReason refusal) =>
        new(null, 0, refusal);
}

/// <summary>
/// Reads and applies the visibility of the columns the ribbon's dropdown offers
/// (R5.18h), over the live <c>tblGanttData</c> table.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port — so the
/// AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c> (the CS0433
/// duplicate-type hazard documented on <see cref="IExcelApplicationAdapter"/>).
/// </para>
/// <para>
/// <b>Only the five dropdown columns.</b> Which columns those are comes from
/// <see cref="Core.GanttColumnVisibility.ToggleableColumnNames"/>, not from a list here:
/// the eleven engine columns are hidden by the schema's own classification
/// (<see cref="Core.GanttColumnAccess.EngineHidden"/>) and the presentation restorer
/// owns them. A port that could also unhide an engine column would let the ribbon
/// produce a table the reader refuses.
/// </para>
/// <para>
/// <b>Excel has no per-table hidden flag.</b> A table column is hidden by hiding the
/// whole worksheet column behind it — what the host's own Hide command does — so both
/// methods reach <c>ListColumn.Range.EntireColumn</c> and read or write
/// <c>Range.Hidden</c>, which this PIA declares as <see cref="object"/> and therefore
/// must be unboxed before comparison.
/// </para>
/// <para>
/// <b><see cref="Apply"/> mutates, so it resolves first and writes second.</b> Every
/// column is found before any flag is written: a missing column must refuse before the
/// worksheet carries a partial change, which is the ADR-0008 D4 ordering the
/// presentation restorer also follows. A column already carrying the requested state is
/// not rewritten, so a repeat of the same choice leaves the workbook clean.
/// </para>
/// </remarks>
public interface IColumnVisibilityPort
{
    /// <summary>Reads the hidden columns of the live table.</summary>
    /// <returns>The hidden set, or a refusal.</returns>
    ColumnVisibilityPortOutcome ReadHidden();

    /// <summary>Applies a hidden set to the live table.</summary>
    /// <param name="hidden">The columns to hide; every other offered column is shown.</param>
    /// <returns>How many columns were written, or a refusal.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="hidden"/> is null.</exception>
    ColumnVisibilityPortOutcome Apply(IReadOnlySet<string> hidden);
}
