namespace GanttCreator.Office;

/// <summary>Why the <c>StyleKey</c> column could not be written.</summary>
public enum StyleKeyWriteRefusalReason
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

/// <summary>The typed result of writing the <c>StyleKey</c> column.</summary>
/// <param name="CellsWritten">How many cells the bulk write touched.</param>
/// <param name="WriteCount">How many cells the plan asked for.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record StyleKeyWriteOutcome(
    int CellsWritten,
    int WriteCount,
    StyleKeyWriteRefusalReason? Refusal)
{
    /// <summary>Gets whether the write completed.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="cellsWritten">How many cells were written.</param>
    /// <param name="writeCount">How many the plan asked for.</param>
    /// <returns>The successful outcome.</returns>
    public static StyleKeyWriteOutcome Ok(int cellsWritten, int writeCount) =>
        new(cellsWritten, writeCount, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static StyleKeyWriteOutcome Refused(StyleKeyWriteRefusalReason refusal) =>
        new(0, 0, refusal);
}

/// <summary>
/// Writes the engine-owned <c>StyleKey</c> column (column K of <c>tblGanttData</c>)
/// so every row's <c>StyleKey</c> matches its current <c>Type</c>.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive -- no interop type crosses the port -- so
/// the AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c>.
/// </para>
/// <para>
/// <b>Why one ranged write rather than N cell writes.</b> The same cost argument
/// as <see cref="IDurationWritePort"/> applies: the plan names only the rows whose
/// value changes, so the range is read once, patched at those rows, and written
/// back once. An empty plan performs no write at all, so a correctly written sheet
/// is not marked dirty on every Refresh.
/// </para>
/// <para>
/// The port <strong>mutates</strong>: it writes cells. The implementation consults
/// the worksheet protection guard first (ADR-0008 D4) and refuses a protected
/// target. It never writes a value the Core plan did not produce -- calculation is
/// Core-owned and this adapter only writes.
/// </para>
/// </remarks>
public interface IStyleKeyWritePort
{
    /// <summary>Writes the planned <c>StyleKey</c> values.</summary>
    /// <param name="plan">The plan produced by <see cref="Core.StyleKeyWritePlanner.Plan"/>.</param>
    /// <returns>How many cells were written, or the refusal reason.</returns>
    StyleKeyWriteOutcome Write(Core.StyleKeyWritePlan plan);
}
