using GanttCreator.Core.ConfigIntegrity;

namespace GanttCreator.Office;

/// <summary>Excel adapter for approved configuration repair plans.</summary>
/// <remarks>The adapter composes the existing catalogue writer, TypeOptions materialiser, and deferred repair ports rather than duplicating their mutation logic.</remarks>
public class ExcelConfigRepairer(
    object? application,
    IConfigCatalogueWriter? catalogueWriter = null,
    ITypeOptionsMaterialiser? typeOptionsMaterialiser = null,
    IWorksheetProtectionGuard? protectionGuard = null,
    IPlotAnchorRepairer? plotAnchorRepairer = null,
    IGanttRowIdentityRepairer? identityRepairer = null,
    IConfigSheetVisibilityRepairer? visibilityRepairer = null
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

    /// <inheritdoc />
    public ConfigRepairOutcome Repair(ConfigIntegrityPlan plan, bool confirmRequiredChanges)
    {
        ArgumentNullException.ThrowIfNull(plan);

        if (plan.RefusedFindings.Count > 0)
        {
            return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.PlanRefused);
        }

        var confirmCount = plan.ConfirmRequiredFindings.Count;
        if (confirmCount > 0 && !confirmRequiredChanges)
        {
            return ConfigRepairOutcome.Ok(0, confirmCount, plan.QuarantinedFindings.Count);
        }

        // The shared protection guard is intentionally the first operation in
        // this mutating entry point. Delegated adapters repeat it for direct callers.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ConfigRepairOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ConfigRepairRefusalReason.NoActiveWorkbook
                    : ConfigRepairRefusalReason.TargetProtected
            );
        }

        var repaired = 0;
        var catalogueRepair = plan.Findings.Any(finding =>
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

            repaired += plan.Findings.Count(finding => finding.Kind != ConfigIntegrityFindingKind.TypeOptionsMissing);
        }

        if (plan.Findings.Any(finding => finding.Kind == ConfigIntegrityFindingKind.WrongVisibility))
        {
            ConfigSheetVisibilityRepairOutcome visibility = _visibilityRepairer.Repair();
            if (!visibility.Succeeded)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.CatalogueWriteRefused);
            }

            repaired++;
        }

        if (plan.Findings.Any(finding => finding.Kind == ConfigIntegrityFindingKind.PlotAnchorDisagreement))
        {
            PlotAnchorRepairOutcome anchor = _plotAnchorRepairer.Repair();
            if (!anchor.Succeeded)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.CatalogueWriteRefused);
            }

            repaired++;
        }

        if (plan.Findings.Any(finding => finding.Kind == ConfigIntegrityFindingKind.IdentityDamage))
        {
            GanttRowIdentityRepairOutcome identity = _identityRepairer.Repair();
            if (identity.Refusal is not null)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.CatalogueWriteRefused);
            }

            repaired += identity.RepairedCount;
        }

        if (plan.Findings.Any(finding => finding.Kind == ConfigIntegrityFindingKind.TypeOptionsMissing))
        {
            TypeOptionsMaterialiseOutcome typeOptions = _typeOptionsMaterialiser.Materialise();
            if (!typeOptions.Succeeded)
            {
                return ConfigRepairOutcome.Refused(ConfigRepairRefusalReason.TypeOptionsUnavailable);
            }
        }

        return ConfigRepairOutcome.Ok(repaired, 0, plan.QuarantinedFindings.Count);
    }
}
