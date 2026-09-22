namespace GanttCreator.Office;

/// <summary>
/// Why a catalogue-read attempt refused to return the workbook state.
/// </summary>
/// <remarks>
/// Every refusal is a typed, expected outcome — routine, not an exception
/// (docs/02-ARCHITECTURE.md "Error handling"). Unexpected COM failures are
/// not refusals and propagate to the command boundary.
/// </remarks>
public enum ConfigReadRefusalReason
{
    /// <summary>The Excel application object or the active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>No <c>_GanttCreatorConfig</c> worksheet exists (run Initialise sheet first).</summary>
    ConfigSheetMissing = 1,

    /// <summary>An expected catalogue table was not found on the configuration worksheet.</summary>
    TableMissing = 2,

    /// <summary>A catalogue table's headers do not match the contract (ADR-0007 D7).</summary>
    HeaderMismatch = 3,

    /// <summary>A catalogue table's body-row count does not match the contract.</summary>
    RowCountMismatch = 4,

    /// <summary>
    /// A regenerated table's content drifts from the code-owned catalogue
    /// (types, styles, metrics, settings, or config metadata).
    /// </summary>
    CatalogueMismatch = 5,

    /// <summary>
    /// The stored schema version or catalogue hash drifts from the
    /// code-owned catalogue (ADR-0007 D6).
    /// </summary>
    CatalogueHashMismatch = 6,

    /// <summary>
    /// A value is outside its valid range or not a permitted setting value.
    /// </summary>
    ValueOutOfRange = 7,

    /// <summary>A style colour is not uppercase <c>#RRGGBB</c>.</summary>
    BadColourFormat = 8,
}

/// <summary>
/// The typed outcome of one catalogue-read attempt: the workbook ID and the
/// effective settings, or the refusal reason. On a refusal nothing was
/// mutated.
/// </summary>
/// <param name="Succeeded">Whether the catalogues were read and validated.</param>
/// <param name="WorkbookId">
/// The workbook's stable ID from <c>tblGanttConfig</c> on success;
/// <see langword="null"/> on a refusal.
/// </param>
/// <param name="Settings">
/// The effective <c>tblGanttSettings</c> key/value map on success; empty on
/// a refusal.
/// </param>
/// <param name="Refusal">
/// The refusal reason when <paramref name="Succeeded"/> is
/// <see langword="false"/>; otherwise <see langword="null"/>.
/// </param>
public sealed record ConfigReadOutcome(
    bool Succeeded,
    string? WorkbookId,
    IReadOnlyDictionary<string, string> Settings,
    ConfigReadRefusalReason? Refusal)
{
    private static readonly IReadOnlyDictionary<string, string> _emptySettings =
        new Dictionary<string, string>();

    /// <summary>
    /// Creates a success outcome.
    /// </summary>
    /// <param name="workbookId">The workbook's stable ID.</param>
    /// <param name="settings">The effective settings key/value map.</param>
    /// <returns>The success outcome.</returns>
    public static ConfigReadOutcome Ok(string workbookId, IReadOnlyDictionary<string, string> settings) =>
        new(true, workbookId, settings, null);

    /// <summary>
    /// Creates a refusal outcome. Nothing was mutated.
    /// </summary>
    /// <param name="refusal">Why nothing was returned.</param>
    /// <returns>The refusal outcome.</returns>
    public static ConfigReadOutcome Refused(ConfigReadRefusalReason refusal) =>
        new(false, null, _emptySettings, refusal);
}

/// <summary>
/// Narrow port over the catalogue-read operation: bulk-reads the five
/// catalogue tables on the <c>_GanttCreatorConfig</c> worksheet via
/// <c>Value2</c> and validates them against the code-owned Core catalogue
/// (ADR-0007 D6). Read-only: never writes cells or changes application
/// state.
/// </summary>
/// <remarks>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c> (the CS0433 duplicate-type hazard
/// documented on <see cref="IExcelApplicationAdapter"/>).
/// </remarks>
public interface IConfigCatalogueReader
{
    /// <summary>
    /// Reads and validates every catalogue table on the active workbook's
    /// <c>_GanttCreatorConfig</c> worksheet.
    /// </summary>
    /// <returns>
    /// The typed outcome: the workbook ID and settings on success, or the
    /// refusal reason. On a refusal nothing was mutated.
    /// </returns>
    ConfigReadOutcome Read();
}
