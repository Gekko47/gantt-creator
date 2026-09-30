using System.Globalization;

namespace GanttCreator.Core.ConfigIntegrity;
/// <summary>The deterministic integrity conditions detected during configuration repair.</summary>
public enum ConfigIntegrityFindingKind
{
    /// <summary>The configuration worksheet is absent.</summary>
    ConfigSheetMissing = 0,

    /// <summary>The configuration worksheet is not VeryHidden.</summary>
    WrongVisibility = 1,

    /// <summary>A required code-owned catalogue table is absent.</summary>
    CatalogueTableMissing = 2,

    /// <summary>A catalogue table has invalid structure or content.</summary>
    CatalogueTableCorrupt = 3,

    /// <summary>The stored catalogue hash does not match the current catalogue.</summary>
    CatalogueHashMismatch = 4,

    /// <summary>The workbook schema is older than the running add-in.</summary>
    SchemaVersionOlder = 5,

    /// <summary>The workbook schema is newer than the running add-in.</summary>
    SchemaVersionNewer = 6,

    /// <summary>The stored plot anchor differs from the live table-derived anchor.</summary>
    PlotAnchorDisagreement = 7,

    /// <summary>The TypeOptions name or its target is missing or invalid.</summary>
    TypeOptionsMissing = 8,

    /// <summary>Visible row IDs need repair.</summary>
    IdentityDamage = 9,

    /// <summary>A second helper worksheet exists.</summary>
    SecondHelperSheet = 10,

    /// <summary>Schedule data exists on the configuration worksheet.</summary>
    ScheduleDataOnConfig = 11,

    /// <summary>A user style row could not be parsed and was quarantined.</summary>
    UserStyleQuarantined = 12,

    /// <summary>The damage class is not recognised by the running add-in.</summary>
    Unknown = 13,
}

/// <summary>The safe-repair classification for one integrity finding.</summary>
public enum ConfigRepairClassification
{
    /// <summary>The finding is add-in-owned and loss-free to repair automatically.</summary>
    AutoRepair = 0,

    /// <summary>The finding may touch user-authored data and requires confirmation.</summary>
    ConfirmRequired = 1,

    /// <summary>The finding is unsafe, unsupported, or unknown and must not be changed.</summary>
    Refuse = 2,
}

/// <summary>
/// Builds the Ribbon-visible detail for a schema-version mismatch (R4.7C D6,
/// ADR-0029 D6).
/// </summary>
/// <remarks>
/// <para>
/// A version mismatch is the one integrity finding that is <em>reported and never
/// converted</em>: there is no migration, so the message has to tell the analyst
/// what to do rather than imply that something will be fixed. A generic
/// "catalogue mismatch" string is not actionable — it does not say which version
/// the workbook carries, which the add-in expects, or that re-initialising is the
/// remedy.
/// </para>
/// <para>
/// The remedy is version-agnostic on purpose: a workbook from any earlier version
/// lacks this schema, so every one of them needs Initialise. Naming the numbers
/// keeps the message honest about what actually drifted.
/// </para>
/// </remarks>
public static class SchemaVersionMismatchMessage
{
    /// <summary>
    /// Builds the detail text for a workbook whose stored schema version differs
    /// from the running add-in's.
    /// </summary>
    /// <param name="workbookVersion">The version stored on the workbook.</param>
    /// <param name="expectedVersion">The version the running add-in requires.</param>
    /// <returns>The actionable, Ribbon-visible detail.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when either version is not positive. A non-positive version is not a
    /// version this builder can describe, and rendering "version 0" would be less
    /// honest than refusing.
    /// </exception>
    public static string Build(int workbookVersion, int expectedVersion)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(workbookVersion);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(expectedVersion);

        // Invariant culture: the version numbers are the only interpolations, but
        // a rendered message must not vary with the analyst's locale digits.
        return string.Create(
            CultureInfo.InvariantCulture,
            $"This workbook uses Gantt schema version {workbookVersion}, but this version of Gantt Creator requires version {expectedVersion}. There is no automatic upgrade: run Initialise Gantt Sheet on a copy of this workbook to create a current one, then enter your schedule again.");
    }
}

/// <summary>One immutable configuration-integrity finding.</summary>
/// <param name="Kind">The detected damage class.</param>
/// <param name="Detail">A non-sensitive diagnostic detail; it must not contain schedule content.</param>
/// <param name="TableName">The affected configuration table, when applicable.</param>
/// <param name="RowNumber">The affected visible-table row, when applicable.</param>
public sealed record ConfigIntegrityFinding(
    ConfigIntegrityFindingKind Kind,
    string Detail,
    string? TableName = null,
    int? RowNumber = null
);

/// <summary>The immutable, deterministic repair plan derived from findings.</summary>
/// <param name="Findings">All findings in deterministic order.</param>
public sealed record ConfigIntegrityPlan(IReadOnlyList<ConfigIntegrityFinding> Findings)
{
    private static readonly IReadOnlyList<ConfigIntegrityFinding> _emptyFindings = [];

    /// <summary>Gets findings classified for automatic repair.</summary>
    public IReadOnlyList<ConfigIntegrityFinding> AutoRepairFindings => Classify(ConfigRepairClassification.AutoRepair);

    /// <summary>Gets findings that require confirmation.</summary>
    public IReadOnlyList<ConfigIntegrityFinding> ConfirmRequiredFindings => Classify(ConfigRepairClassification.ConfirmRequired);

    /// <summary>Gets findings that must be refused.</summary>
    public IReadOnlyList<ConfigIntegrityFinding> RefusedFindings => Classify(ConfigRepairClassification.Refuse);

    /// <summary>Gets findings representing user style rows that were quarantined.</summary>
    public IReadOnlyList<ConfigIntegrityFinding> QuarantinedFindings =>
        [.. Findings.Where(finding => finding.Kind == ConfigIntegrityFindingKind.UserStyleQuarantined)];

    /// <summary>Creates an empty plan.</summary>
    /// <returns>The empty deterministic plan.</returns>
    public static ConfigIntegrityPlan Empty() => new(_emptyFindings);

    /// <summary>Builds a stable plan from findings, sorting by the explicit stable key.</summary>
    /// <param name="findings">The findings to classify.</param>
    /// <returns>The immutable plan.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="findings"/> is null.</exception>
    public static ConfigIntegrityPlan Build(IEnumerable<ConfigIntegrityFinding> findings)
    {
        ArgumentNullException.ThrowIfNull(findings);
        ConfigIntegrityFinding[] ordered =
        [
            .. findings
                .OrderBy(finding => finding.Kind)
                .ThenBy(finding => finding.TableName, StringComparer.Ordinal)
                .ThenBy(finding => finding.RowNumber ?? int.MaxValue)
                .ThenBy(finding => finding.Detail, StringComparer.Ordinal),
        ];
        return new ConfigIntegrityPlan(ordered);
    }

    /// <summary>Classifies one finding using ADR-0011's safe-repair boundary.</summary>
    /// <param name="finding">The finding to classify.</param>
    /// <returns>The classification for <paramref name="finding"/>.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="finding"/> is null.</exception>
    public static ConfigRepairClassification Classify(ConfigIntegrityFinding finding)
    {
        ArgumentNullException.ThrowIfNull(finding);
        return finding.Kind switch
        {
            ConfigIntegrityFindingKind.WrongVisibility => ConfigRepairClassification.AutoRepair,
            ConfigIntegrityFindingKind.CatalogueTableMissing
                when finding.TableName is GanttCatalogues.TypesTableName or GanttCatalogues.MetricsTableName =>
                ConfigRepairClassification.AutoRepair,
            ConfigIntegrityFindingKind.CatalogueTableCorrupt
                when finding.TableName is GanttCatalogues.TypesTableName or GanttCatalogues.MetricsTableName =>
                ConfigRepairClassification.AutoRepair,
            ConfigIntegrityFindingKind.CatalogueHashMismatch
            or ConfigIntegrityFindingKind.PlotAnchorDisagreement
            or ConfigIntegrityFindingKind.TypeOptionsMissing => ConfigRepairClassification.AutoRepair,
            ConfigIntegrityFindingKind.SchemaVersionOlder or ConfigIntegrityFindingKind.IdentityDamage =>
                ConfigRepairClassification.ConfirmRequired,
            ConfigIntegrityFindingKind.CatalogueTableMissing when finding.TableName == GanttCatalogues.StylesTableName =>
                ConfigRepairClassification.ConfirmRequired,
            ConfigIntegrityFindingKind.CatalogueTableCorrupt when finding.TableName == GanttCatalogues.StylesTableName =>
                ConfigRepairClassification.ConfirmRequired,
            ConfigIntegrityFindingKind.SchemaVersionNewer
            or ConfigIntegrityFindingKind.ConfigSheetMissing
            or ConfigIntegrityFindingKind.SecondHelperSheet
            or ConfigIntegrityFindingKind.ScheduleDataOnConfig
            or ConfigIntegrityFindingKind.Unknown => ConfigRepairClassification.Refuse,
            ConfigIntegrityFindingKind.UserStyleQuarantined => ConfigRepairClassification.ConfirmRequired,
            _ => ConfigRepairClassification.Refuse,
        };
    }

    private ConfigIntegrityFinding[] Classify(ConfigRepairClassification classification) =>
        [.. Findings.Where(finding => Classify(finding) == classification)];
}
