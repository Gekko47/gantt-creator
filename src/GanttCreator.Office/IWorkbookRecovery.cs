using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>Why a workbook recovery attempt refused to mutate.</summary>
public enum WorkbookRecoveryRefusalReason
{
    /// <summary>No active workbook was supplied.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The target worksheet or workbook structure is protected.</summary>
    TargetProtected = 1,

    /// <summary>No worksheet carries the contract Gantt label.</summary>
    TargetMissing = 2,

    /// <summary>More than one worksheet carries the contract Gantt label.</summary>
    TargetAmbiguous = 3,

    /// <summary>The target worksheet carries a table other than the contract table.</summary>
    TargetHasForeignTable = 4,

    /// <summary>The contract header row does not align with the schema.</summary>
    HeaderMisaligned = 5,

    /// <summary>The table structure cannot be restored without guessing.</summary>
    TableStructureRefused = 6,

    /// <summary>The host refused to create the table.</summary>
    TableCreateFailed = 7,

    /// <summary>The host refused a header or column write.</summary>
    HeaderWriteFailed = 8,

    /// <summary>The rebuilt table could not be read back.</summary>
    VerificationFailed = 9,

    /// <summary>The workbook uses the 1904 date system, which the table reader rejects.</summary>
    DateSystemUnsupported = 10,

    /// <summary>The user declined the confirm-required recovery.</summary>
    ConfirmationDeclined = 11,

    /// <summary>A visible worksheet carries the configuration name but no contract table.</summary>
    ConfigurationNameCollision = 12,

    /// <summary>More than one worksheet could be the configuration.</summary>
    ConfigurationAmbiguous = 13,

    /// <summary>A visible worksheet carries a configuration contract table.</summary>
    ConfigurationCandidateAmbiguous = 14,

    /// <summary>A non-canonical VeryHidden worksheet carries no contract table.</summary>
    ConfigurationUnrecognised = 15,

    /// <summary>The host refused the configuration worksheet rename or its visibility write.</summary>
    ConfigurationRenameFailed = 16,

    /// <summary>The host refused the misnamed table rename.</summary>
    TableRenameFailed = 17,

    /// <summary>The host refused to create the configuration worksheet.</summary>
    ConfigurationCreateFailed = 18,

    /// <summary>The catalogue writer refused to materialise the configuration.</summary>
    ConfigurationWriteFailed = 19,

    /// <summary>The host refused to create the Gantt worksheet.</summary>
    GanttSheetCreateFailed = 20,
}

/// <summary>
/// A read-only inspection of the visible table recovery state.
/// </summary>
/// <param name="TableReadable">Whether the contract table is readable.</param>
/// <param name="TablePresent">Whether a contract table exists.</param>
/// <param name="ConfigurationPresent">Whether the configuration worksheet exists.</param>
/// <param name="NeedsRecovery">Whether this inspection has a table recovery to perform.</param>
/// <param name="TargetSheetName">The worksheet the recovery targets, when resolved.</param>
/// <param name="BodyRowCount">The body rows that would be adopted by a rebuild.</param>
/// <param name="HeaderPlan">The header alignment plan for a rebuild.</param>
/// <param name="StructurePlan">The column repair plan for an unreadable table.</param>
/// <param name="RequiresConfirmation">Whether adopting body rows needs consent.</param>
/// <param name="ConfirmationMessage">The confirmation text when consent is required.</param>
/// <param name="Refusal">The typed refusal, or <see langword="null"/>.</param>
public sealed record WorkbookRecoveryInspection(
    bool TableReadable,
    bool TablePresent,
    bool ConfigurationPresent,
    bool NeedsRecovery,
    string? TargetSheetName,
    int BodyRowCount,
    GanttTableHeaderAlignmentPlan? HeaderPlan,
    GanttTableStructureRepairPlan? StructurePlan,
    bool RequiresConfirmation,
    string? ConfirmationMessage,
    WorkbookRecoveryRefusalReason? Refusal)
{
    /// <summary>
    /// Whether the recovery also creates the missing configuration worksheet
    /// before repairing the table.
    /// </summary>
    public bool NeedsConfigurationRecovery { get; init; }

    /// <summary>
    /// Whether the recovery renames one misnamed VeryHidden worksheet back to the
    /// contract configuration name. Never combined with
    /// <see cref="NeedsConfigurationRecovery"/>.
    /// </summary>
    public bool NeedsConfigurationRename { get; init; }

    /// <summary>
    /// Whether the recovery renames one table on the Gantt worksheet to the
    /// contract table name after proving the sheet carries the chart.
    /// </summary>
    public bool NeedsTableRename { get; init; }

    /// <summary>The live table name when <see cref="NeedsTableRename"/> is set.</summary>
    public string? TargetTableName { get; init; }

    /// <summary>
    /// Whether the recovery creates a missing Gantt worksheet before writing the
    /// table. The confirmation must state that old schedule rows are not restored.
    /// </summary>
    public bool NeedsGanttSheet { get; init; }
}

/// <summary>The typed result of one recovery attempt.</summary>
/// <param name="Changed">Whether the workbook was mutated.</param>
/// <param name="TableRebuilt">Whether the contract table was rebuilt.</param>
/// <param name="TableStructureRepaired">Whether headers or columns were repaired in place.</param>
/// <param name="AdoptedRowCount">How many existing body rows became table rows.</param>
/// <param name="Refusal">The typed refusal, or <see langword="null"/>.</param>
public sealed record WorkbookRecoveryOutcome(
    bool Changed,
    bool TableRebuilt,
    bool TableStructureRepaired,
    int AdoptedRowCount,
    WorkbookRecoveryRefusalReason? Refusal)
{
    /// <summary>Whether the configuration worksheet was created.</summary>
    public bool ConfigurationCreated { get; init; }

    /// <summary>Whether a misnamed configuration worksheet was renamed back.</summary>
    public bool ConfigurationRenamed { get; init; }

    /// <summary>Whether a proven misnamed Gantt table was renamed to the contract name.</summary>
    public bool TableRenamed { get; init; }

    /// <summary>Whether a new Gantt worksheet was created.</summary>
    public bool GanttSheetCreated { get; init; }

    /// <summary>Creates a no-change outcome.</summary>
    /// <returns>The outcome.</returns>
    public static WorkbookRecoveryOutcome NoChange() => new(false, false, false, 0, null);

    /// <summary>Creates a refusal outcome; nothing was mutated.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The outcome.</returns>
    public static WorkbookRecoveryOutcome Refused(WorkbookRecoveryRefusalReason refusal) =>
        new(false, false, false, 0, refusal);
}

/// <summary>Read-only inspection and constrained recovery of the visible table.</summary>
public interface IWorkbookRecovery
{
    /// <summary>Inspects the active workbook without mutating it.</summary>
    /// <returns>The typed inspection.</returns>
    WorkbookRecoveryInspection Inspect();

    /// <summary>Performs the recovery described by <paramref name="inspection"/>.</summary>
    /// <param name="inspection">The inspection returned by <see cref="Inspect"/>.</param>
    /// <param name="confirmed">Whether the user approved the confirm-required changes.</param>
    /// <returns>The typed outcome.</returns>
    WorkbookRecoveryOutcome Recover(WorkbookRecoveryInspection inspection, bool confirmed);
}
