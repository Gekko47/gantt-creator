namespace GanttCreator.Office;

/// <summary>Why the <c>Duration</c> column could not be written.</summary>
public enum DurationWriteRefusalReason
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The Gantt worksheet or its <c>tblGanttData</c> table was missing.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet is protected.</summary>
    TargetProtected = 2,

    /// <summary>The host refused the ranged write.</summary>
    WriteFailed = 3,
}

/// <summary>The typed result of writing the <c>Duration</c> column.</summary>
/// <param name="CellsWritten">How many cells the bulk write touched.</param>
/// <param name="WriteCount">How many cells the plan asked for.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record DurationWriteOutcome(
    int CellsWritten,
    int WriteCount,
    DurationWriteRefusalReason? Refusal)
{
    /// <summary>Gets whether the write completed.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="cellsWritten">How many cells were written.</param>
    /// <param name="writeCount">How many the plan asked for.</param>
    /// <returns>The successful outcome.</returns>
    public static DurationWriteOutcome Ok(int cellsWritten, int writeCount) =>
        new(cellsWritten, writeCount, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static DurationWriteOutcome Refused(DurationWriteRefusalReason refusal) =>
        new(0, 0, refusal);
}

/// <summary>
/// Writes the engine-owned <c>Duration</c> column in one bulk pass (R4.7F D5).
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive -- no interop type crosses the port -- so
/// the AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c>.
/// </para>
/// <para>
/// <b>Why one ranged write rather than N cell writes.</b> D5 requires it, and the
/// reason is observable rather than aesthetic: N COM round-trips over a 1,000-event
/// table dominate Refresh time (REV5 §22 instruction 12). The port therefore takes
/// the whole <see cref="Core.DurationWritePlan"/> and performs one write over the
/// <c>Duration</c> column range, never a per-cell loop. An empty plan must perform no
/// write at all, so a correctly written sheet is not marked dirty.
/// </para>
/// <para>
/// The port <strong>mutates</strong>: it writes cells. The implementation consults
/// the worksheet protection guard first (ADR-0008 D4) and refuses a protected
/// target. It never writes a value the Core plan did not produce -- calculation is
/// Core-owned and this adapter only writes.
/// </para>
/// </remarks>
public interface IDurationWritePort
{
    /// <summary>Writes the planned <c>Duration</c> values.</summary>
    /// <param name="plan">The plan produced by <see cref="Core.DurationCalculator.Plan"/>.</param>
    /// <returns>How many cells were written, or the refusal reason.</returns>
    DurationWriteOutcome Write(Core.DurationWritePlan plan);
}
