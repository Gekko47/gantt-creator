namespace GanttCreator.Office;

/// <summary>
/// Why a table-read attempt refused to return rows.
/// </summary>
/// <remarks>
/// Every refusal is a typed, expected outcome — routine, not an exception
/// (docs/02-ARCHITECTURE.md "Error handling"). Unexpected COM failures are
/// not refusals and propagate to the command boundary.
/// </remarks>
public enum GanttTableReadRefusalReason
{
    /// <summary>The Excel application object or the active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>No table named <c>tblGanttData</c> was found anywhere in the workbook.</summary>
    TableMissing = 1,

    /// <summary>The workbook uses the 1904 date system, which R2.4 explicitly rejects (see ADR-0006).</summary>
    DateSystemUnsupported = 2,

    /// <summary>
    /// A <c>tblGanttData</c> table exists but its required headers or columns do
    /// not resolve. This is distinct from <see cref="TableMissing"/> because the
    /// table object is present and can be repaired in place (R5.17a).
    /// </summary>
    TableUnreadable = 3,
}

/// <summary>
/// The typed outcome of one table-read attempt: the rows read, or the
/// refusal reason. On a refusal nothing was mutated.
/// </summary>
/// <param name="Succeeded">Whether rows were returned.</param>
/// <param name="Rows">The rows in body order; empty on a refusal.</param>
/// <param name="Refusal">The refusal reason when <paramref name="Succeeded"/> is <see langword="false"/>; otherwise <see langword="null"/>.</param>
/// <param name="ConfigurationPresent">
/// Whether the workbook carries a <c>_GanttCreatorConfig</c> worksheet. It is
/// computed on every read: on a <see cref="GanttTableReadRefusalReason.TableMissing"/>
/// refusal it separates a first run from a damaged workbook (R5.16, ADR-0056),
/// and on success the Ribbon uses it to keep <c>Ready</c> honest when the
/// configuration is absent (R5.17b).
/// </param>
/// <param name="TablePresent">
/// Whether a <c>tblGanttData</c> table exists. It is true on success and on
/// <see cref="GanttTableReadRefusalReason.TableUnreadable"/>, false for
/// <see cref="GanttTableReadRefusalReason.TableMissing"/>.
/// </param>
public sealed record GanttTableReadOutcome(
    bool Succeeded,
    IReadOnlyList<Core.GanttRowDto> Rows,
    GanttTableReadRefusalReason? Refusal,
    bool ConfigurationPresent = false,
    bool TablePresent = false)
{
    /// <summary>
    /// Whether a worksheet carries <see cref="Core.GanttWorkbookContract.GanttSheetLabel"/>.
    /// The Ribbon uses it to distinguish a recoverable damaged workbook from a
    /// genuine first run when both containers are absent.
    /// </summary>
    public bool GanttSheetPresent { get; init; }

    /// <summary>
    /// Whether the workbook's active sheet is the worksheet that carries
    /// <c>tblGanttData</c>. The Ribbon greys every non-exempt control when this
    /// is <see langword="false"/>, so a chart command cannot be run against a
    /// different visible sheet (R5.16b, ADR-0058). On a
    /// <see cref="GanttTableReadRefusalReason.TableMissing"/> refusal it reports
    /// whether the active sheet carries the Gantt label instead.
    /// </summary>
    public bool ActiveSheetIsGanttSheet { get; init; }

    /// <summary>
    /// Creates a success outcome.
    /// </summary>
    /// <param name="rows">The rows in body order.</param>
    /// <param name="configurationPresent">Whether the workbook carries a <c>_GanttCreatorConfig</c> worksheet.</param>
    /// <param name="tablePresent">Whether the contract table exists. True on every success.</param>
    /// <param name="ganttSheetPresent">Whether a standard-labelled Gantt worksheet exists.</param>
    /// <param name="activeSheetIsGanttSheet">
    /// Whether the active sheet carries the Gantt table. It defaults to
    /// <see langword="true"/> — the fail-open direction for a snapshot built from
    /// a successful read — because a hand-built outcome in a test must not grey
    /// every control by omission. Every live read supplies the measured value.
    /// </param>
    /// <returns>The success outcome.</returns>
    public static GanttTableReadOutcome Ok(
        IReadOnlyList<Core.GanttRowDto> rows,
        bool configurationPresent = false,
        bool tablePresent = true,
        bool ganttSheetPresent = false,
        bool activeSheetIsGanttSheet = true) =>
        new(true, rows, null, configurationPresent, tablePresent)
        {
            GanttSheetPresent = ganttSheetPresent,
            ActiveSheetIsGanttSheet = activeSheetIsGanttSheet,
        };

    /// <summary>
    /// Creates a refusal outcome. Nothing was mutated.
    /// </summary>
    /// <param name="refusal">Why no rows were returned.</param>
    /// <param name="configurationPresent">
    /// Whether the workbook carries a <c>_GanttCreatorConfig</c> worksheet.
    /// </param>
    /// <param name="tablePresent">
    /// Whether a contract table exists. True for <see cref="GanttTableReadRefusalReason.TableUnreadable"/>.
    /// </param>
    /// <param name="ganttSheetPresent">Whether a standard-labelled Gantt worksheet exists.</param>
    /// <param name="activeSheetIsGanttSheet">Whether the active sheet carries the Gantt table.</param>
    /// <returns>The refusal outcome.</returns>
    public static GanttTableReadOutcome Refused(
        GanttTableReadRefusalReason refusal,
        bool configurationPresent = false,
        bool tablePresent = false,
        bool ganttSheetPresent = false,
        bool activeSheetIsGanttSheet = false) =>
        new(false, [], refusal, configurationPresent, tablePresent)
        {
            GanttSheetPresent = ganttSheetPresent,
            ActiveSheetIsGanttSheet = activeSheetIsGanttSheet,
        };
}
