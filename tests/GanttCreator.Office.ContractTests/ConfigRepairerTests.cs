using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.Office.ContractTests;

public sealed class ConfigRepairerTests
{
    private static ConfigIntegrityPlan Plan(ConfigIntegrityFindingKind kind, string table = GanttCatalogues.TypesTableName) =>
        ConfigIntegrityPlan.Build([new ConfigIntegrityFinding(kind, "test", table)]);

    [Fact]
    public void Repair_refuses_a_refused_plan_before_calling_the_guard_or_writer()
    {
        var guard = new Mock<IWorksheetProtectionGuard>(MockBehavior.Strict);
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        var plan = Plan(ConfigIntegrityFindingKind.SchemaVersionNewer);

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(null, writer.Object, null, guard.Object).Repair(plan, true);

        Assert.Equal(ConfigRepairRefusalReason.PlanRefused, outcome.Refusal);
        guard.Verify(g => g.Query(), Times.Never);
        writer.Verify(w => w.Write(), Times.Never);
    }

    [Fact]
    public void Repair_skips_all_mutation_when_confirmation_is_declined()
    {
        var guard = new Mock<IWorksheetProtectionGuard>(MockBehavior.Strict);
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        ConfigIntegrityPlan plan = Plan(ConfigIntegrityFindingKind.IdentityDamage);

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(null, writer.Object, null, guard.Object).Repair(plan, false);

        Assert.Null(outcome.Refusal);
        Assert.Equal(1, outcome.ConfirmationSkippedCount);
        guard.Verify(g => g.Query(), Times.Never);
        writer.Verify(w => w.Write(), Times.Never);
    }

    [Fact]
    public void Repair_refuses_protection_before_calling_the_writer()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object).Repair(Plan(ConfigIntegrityFindingKind.WrongVisibility), true);

        Assert.Equal(ConfigRepairRefusalReason.TargetProtected, outcome.Refusal);
        guard.Verify(g => g.Query(), Times.Once);
        writer.Verify(w => w.Write(), Times.Never);
    }

    [Fact]
    public void Repair_writes_the_existing_catalogue_for_an_auto_finding()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.Write()).Returns(ConfigWriteOutcome.Ok());

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object).Repair(Plan(ConfigIntegrityFindingKind.CatalogueHashMismatch), true);

        Assert.Null(outcome.Refusal);
        Assert.Equal(1, outcome.RepairedCount);
        writer.Verify(w => w.Write(), Times.Once);
    }

    /// <summary>
    /// The repair path is the one caller allowed to replace a stale TypeOptions
    /// name target, so it must go through the repair-specific entry point and
    /// never through ordinary materialisation.
    /// </summary>
    [Fact]
    public void Repair_uses_the_repair_specific_type_options_entry_point()
    {
        var guard = new Mock<IWorksheetProtectionGuard>(MockBehavior.Strict);
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var typeOptions = new Mock<ITypeOptionsMaterialiser>(MockBehavior.Strict);
        _ = typeOptions.Setup(t => t.MaterialiseForRepair()).Returns(TypeOptionsMaterialiseOutcome.Ok());

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            typeOptionsMaterialiser: typeOptions.Object,
            protectionGuard: guard.Object).Repair(
                Plan(ConfigIntegrityFindingKind.TypeOptionsMissing),
                true);

        Assert.Null(outcome.Refusal);
        typeOptions.Verify(t => t.MaterialiseForRepair(), Times.Once);
        typeOptions.Verify(t => t.Materialise(), Times.Never);
    }

    [Fact]
    public void Repair_refuses_when_the_repair_path_cannot_materialise_type_options()
    {
        var guard = new Mock<IWorksheetProtectionGuard>(MockBehavior.Strict);
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var typeOptions = new Mock<ITypeOptionsMaterialiser>(MockBehavior.Strict);
        _ = typeOptions.Setup(t => t.MaterialiseForRepair())
            .Returns(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NameTargetInvalid));

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            typeOptionsMaterialiser: typeOptions.Object,
            protectionGuard: guard.Object).Repair(
                Plan(ConfigIntegrityFindingKind.TypeOptionsMissing),
                true);

        Assert.Equal(ConfigRepairRefusalReason.TypeOptionsUnavailable, outcome.Refusal);
    }

    [Fact]
    public void Repair_dispatches_visibility_anchor_and_identity_ports_without_catalogue_rewrite()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        var typeOptions = new Mock<ITypeOptionsMaterialiser>(MockBehavior.Strict);
        var visibility = new Mock<IConfigSheetVisibilityRepairer>();
        _ = visibility.Setup(r => r.Repair()).Returns(ConfigSheetVisibilityRepairOutcome.Ok());
        var anchor = new Mock<IPlotAnchorRepairer>();
        _ = anchor.Setup(r => r.Repair()).Returns(PlotAnchorRepairOutcome.Ok());
        var identity = new Mock<IGanttRowIdentityRepairer>();
        _ = identity.Setup(r => r.Repair()).Returns(GanttRowIdentityRepairOutcome.Ok(2));
        ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build(
        [
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.WrongVisibility, "visibility"),
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.PlotAnchorDisagreement, "anchor"),
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.IdentityDamage, "identity"),
        ]);

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            typeOptions.Object,
            guard.Object,
            anchor.Object,
            identity.Object,
            visibility.Object).Repair(plan, true);

        Assert.Null(outcome.Refusal);
        Assert.Equal(4, outcome.RepairedCount);
        writer.Verify(w => w.Write(), Times.Never);
        typeOptions.Verify(t => t.Materialise(), Times.Never);
        visibility.Verify(r => r.Repair(), Times.Once);
        anchor.Verify(r => r.Repair(), Times.Once);
        identity.Verify(r => r.Repair(), Times.Once);
    }

    [Fact]
    public void Repair_distinguishes_visibility_failure_from_catalogue_failure()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        var visibility = new Mock<IConfigSheetVisibilityRepairer>();
        _ = visibility.Setup(r => r.Repair())
            .Returns(ConfigSheetVisibilityRepairOutcome.Refused(ConfigSheetVisibilityRepairRefusalReason.TargetProtected));

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object,
            visibilityRepairer: visibility.Object).Repair(
                Plan(ConfigIntegrityFindingKind.WrongVisibility),
                true);

        Assert.Equal(ConfigRepairRefusalReason.VisibilityRepairFailed, outcome.Refusal);
        writer.Verify(w => w.Write(), Times.Never);
    }

    [Fact]
    public void Repair_distinguishes_anchor_and_identity_failures()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);
        var anchor = new Mock<IPlotAnchorRepairer>();
        _ = anchor.Setup(r => r.Repair())
            .Returns(PlotAnchorRepairOutcome.Refused(PlotAnchorRepairRefusalReason.NameWriteFailed));
        var identity = new Mock<IGanttRowIdentityRepairer>();
        _ = identity.Setup(r => r.Repair())
            .Returns(GanttRowIdentityRepairOutcome.Refused(GanttRowIdentityRepairRefusalReason.WriteFailed));

        ConfigRepairOutcome anchorOutcome = new ExcelConfigRepairer(
            null, writer.Object, null, guard.Object, plotAnchorRepairer: anchor.Object)
            .Repair(Plan(ConfigIntegrityFindingKind.PlotAnchorDisagreement), true);
        Assert.Equal(ConfigRepairRefusalReason.PlotAnchorRepairFailed, anchorOutcome.Refusal);

        ConfigRepairOutcome identityOutcome = new ExcelConfigRepairer(
            null, writer.Object, null, guard.Object, identityRepairer: identity.Object)
            .Repair(Plan(ConfigIntegrityFindingKind.IdentityDamage), true);
        Assert.Equal(ConfigRepairRefusalReason.IdentityRepairFailed, identityOutcome.Refusal);
    }

    [Fact]
    public void Repair_rejects_a_null_plan()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        var writer = new Mock<IConfigCatalogueWriter>();

        Assert.Throws<ArgumentNullException>(() =>
            new ExcelConfigRepairer(null, writer.Object, null, guard.Object).Repair(null!, true));
    }
}
