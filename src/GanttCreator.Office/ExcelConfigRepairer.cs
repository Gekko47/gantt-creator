using GanttCreator.Core.ConfigIntegrity;

namespace GanttCreator.Office;

/// <summary>Excel adapter for approved configuration repair plans.</summary>
/// <remarks>
/// <para>
/// The adapter composes the existing catalogue writer, TypeOptions materialiser, and deferred repair
/// ports rather than duplicating their mutation logic.
/// </para>
/// <para>
/// <b>Protection (ADR-0011 amendment 2026-11-21).</b> A protected Gantt worksheet is released for the
/// duration of the repair and re-protected afterwards. The protection a reopened workbook carries is
/// not the current session's arming — <c>UserInterfaceOnly</c> is not persisted — so the window is
/// asked to release foreign protection, and a protection it cannot release (a password the add-in does
/// not have) still refuses before any write.
/// </para>
/// </remarks>
public class ExcelConfigRepairer(
    object? application,
    IConfigCatalogueWriter? catalogueWriter = null,
    ITypeOptionsMaterialiser? typeOptionsMaterialiser = null,
    IWorksheetProtectionGuard? protectionGuard = null,
    IPlotAnchorRepairer? plotAnchorRepairer = null,
    IGanttRowIdentityRepairer? identityRepairer = null,
    IConfigSheetVisibilityRepairer? visibilityRepairer = null,
    IWorksheetProtectionPort? protectionPort = null
) : IConfigRepairer
{
    private readonly IConfigCatalogueWriter _catalogueWriter =
        catalogueWriter ?? new ExcelConfigCatalogueWriter(application, protectionGuard);
    private readonly ITypeOptionsMaterialiser _typeOptionsMaterialiser =
        typeOptionsMaterialiser ?? new ExcelTypeOptionsMaterialiser(application, catalogueReader: null, protectionGuard);
    private readonly IPlotAnchorRepairer _plotAnchorRepairer =
        plotAnchorRepairer ?? new ExcelPlotAnchorRepairer(application, protectionGuard);
    private readonly IGanttRowIdentityRepairer _identityRepairer =
        identityRepairer ?? new ExcelGanttRowIdentityRepairer(application, protectionGuard);
    private readonly IConfigSheetVisibilityRepairer _visibilityRepairer =
        visibilityRepairer ?? new ExcelConfigSheetVisibilityRepairer(application, protectionGuard);
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>
    /// The release window this repair runs inside when the target is protected.
    /// </summary>
    /// <remarks>
    /// Its arming state is deliberately a private, never-armed instance: the repair must release
    /// whatever protection it finds and put back exactly what it found, so the session's "is this ours"
    /// fact is the wrong question here and asking it would make the release depend on whether a
    /// refresh had already run in this session.
    /// </remarks>
    private readonly WorksheetProtectionWindow _protectionWindow = new(
        protectionPort ?? new ExcelWorksheetProtectionPort(application),
        new ProtectionArmingState());

    /// <inheritdoc />
    public ConfigRepairOutcome Repair(ConfigIntegrityPlan plan, bool confirmRequiredChanges)
    {
        ArgumentNullException.ThrowIfNull(plan);

        if (plan.RefusedFindings.Count > 0)
        {
            return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.PlanRefused);
        }

        int confirmCount = plan.ConfirmRequiredFindings.Count;
        if (confirmCount > 0 && !confirmRequiredChanges)
        {
            return ConfigRepairOutcome.Ok(0, confirmCount, plan.QuarantinedFindings.Count);
        }

        // The shared protection guard is intentionally the first operation in
        // this mutating entry point. Delegated adapters repeat it for direct callers.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection is ProtectionGuardOutcome.NoActiveWorkbook
            or ProtectionGuardOutcome.WorkbookStructureProtected)
        {
            return ConfigRepairOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ConfigRepairRefusalReason.NoActiveWorkbook
                    : ConfigRepairRefusalReason.TargetProtected
            );
        }

        // A protected worksheet is released around the repair itself, not refused: the workbook Fix
        // exists to repair is the one a previous session left protected, and demanding a manual
        // unprotect first made the command unusable on it (owner report 2026-11-21). The second guard
        // query is what keeps a protection the window cannot release (a password we do not have) a
        // refusal, and it happens before any write.
        using IDisposable? release = protection == ProtectionGuardOutcome.SheetProtected
            ? _protectionWindow.Open(repairCells: true, releaseForeignProtection: true)
            : null;
        if (release is not null && _protectionGuard.Query() != ProtectionGuardOutcome.NotProtected)
        {
            return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.TargetProtected);
        }

        int repaired = 0;
        bool catalogueRepair = plan.Findings.Any(finding =>
            finding.Kind
                is ConfigIntegrityFindingKind.CatalogueTableMissing
                    or ConfigIntegrityFindingKind.CatalogueTableCorrupt
                    or ConfigIntegrityFindingKind.CatalogueHashMismatch
                    or ConfigIntegrityFindingKind.SchemaVersionOlder
                    or ConfigIntegrityFindingKind.UserStyleQuarantined
        );
        if (catalogueRepair)
        {
            ConfigWriteOutcome catalogueOutcome = _catalogueWriter.Write();
            if (!catalogueOutcome.Succeeded)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.CatalogueWriteRefused);
            }

            repaired++;
        }

        if (plan.Findings.Any(finding => finding.Kind == ConfigIntegrityFindingKind.WrongVisibility))
        {
            ConfigSheetVisibilityRepairOutcome visibility = _visibilityRepairer.Repair();
            if (!visibility.Succeeded)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.VisibilityRepairFailed);
            }

            repaired++;
        }

        if (plan.Findings.Any(finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement))
        {
            PlotAnchorRepairOutcome anchor = _plotAnchorRepairer.Repair();
            if (!anchor.Succeeded)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.PlotAnchorRepairFailed);
            }

            repaired++;
        }

        if (plan.Findings.Any(finding => finding.Kind == ConfigIntegrityFindingKind.IdentityDamage))
        {
            GanttRowIdentityRepairOutcome identity = _identityRepairer.Repair();
            if (identity.Refusal is not null)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.IdentityRepairFailed);
            }

            repaired += identity.RepairedCount;
        }

        if (plan.Findings.Any(finding => finding.Kind == ConfigIntegrityFindingKind.TypeOptionsMissing))
        {
            // The repair path owns the catalogue rewrite that can leave the
            // TypeOptions name pointing at the previous range, so it is the one
            // caller allowed to replace that stale target.
            TypeOptionsMaterialiseOutcome typeOptions = _typeOptionsMaterialiser.MaterialiseForRepair();
            if (!typeOptions.Succeeded)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.TypeOptionsUnavailable);
            }
        }

        return ConfigRepairOutcome.Ok(repaired, 0, plan.QuarantinedFindings.Count);
    }
}
