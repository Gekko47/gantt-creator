namespace GanttCreator.Office;

/// <summary>
/// Why a catalogue-write attempt refused to touch the workbook.
/// </summary>
/// <remarks>
/// Every refusal is a typed, expected outcome — routine, not an exception
/// (docs/02-ARCHITECTURE.md "Error handling"). Unexpected COM failures are
/// not refusals and propagate to the command boundary.
/// </remarks>
public enum ConfigWriteRefusalReason
{
    /// <summary>The Excel application object or the active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>
    /// No <c>_GanttCreatorConfig</c> worksheet exists; catalogues are
    /// materialised during initialise (run Initialise sheet first).
    /// </summary>
    ConfigSheetMissing = 1,

    /// <summary>The configuration worksheet or the workbook structure is protected;
    /// nothing was written.
    /// </summary>
    TargetProtected = 2,
    /// <summary>Existing configuration content could not be preserved safely.</summary>
    CataloguePreservationInvalid = 3,
}

/// <summary>
/// The typed outcome of one catalogue-write attempt. On a refusal nothing
/// was mutated.
/// </summary>
/// <param name="Succeeded">Whether the catalogues were written.</param>
/// <param name="Refusal">The refusal reason when <paramref name="Succeeded"/> is <see langword="false"/>; otherwise <see langword="null"/>.</param>
public sealed record ConfigWriteOutcome(
    bool Succeeded,
    ConfigWriteRefusalReason? Refusal)
{
    /// <summary>
    /// Creates a success outcome.
    /// </summary>
    /// <returns>The success outcome.</returns>
    public static ConfigWriteOutcome Ok() => new(true, null);

    /// <summary>
    /// Creates a refusal outcome. Nothing was mutated.
    /// </summary>
    /// <param name="refusal">Why nothing was written.</param>
    /// <returns>The refusal outcome.</returns>
    public static ConfigWriteOutcome Refused(ConfigWriteRefusalReason refusal) =>
        new(false, refusal);
}

/// <summary>
/// Narrow port over the catalogue-write operation: materialises the
/// code-owned configuration catalogues (ADR-0007 D2) as Excel tables on the
/// <c>_GanttCreatorConfig</c> worksheet, preserving user-authored rows
/// (ADR-0007 D4). Called by the workbook initialiser and by the R2.10
/// repair workflow.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c> (the CS0433 duplicate-type hazard
/// documented on <see cref="IExcelApplicationAdapter"/>).
/// </para>
/// <para>
/// The port writes only catalogue rows on the configuration worksheet. It
/// never writes schedule data, never creates a second helper sheet, and
/// never touches the visible worksheet.
/// </para>
/// </remarks>
public interface IConfigCatalogueWriter
{
    /// <summary>
    /// Materialises (or regenerates) the five catalogue tables on the active
    /// workbook's <c>_GanttCreatorConfig</c> worksheet.
    /// </summary>
    /// <returns>
    /// The typed outcome. On a refusal nothing was mutated.
    /// </returns>
    ConfigWriteOutcome Write();
}
