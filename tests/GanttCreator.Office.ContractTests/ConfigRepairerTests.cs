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

    /// <summary>
    /// A protected worksheet is released for the repair and re-protected afterwards, because the
    /// workbook Fix exists to repair is the one a previous session left protected (ADR-0011
    /// amendment 2026-11-21).
    /// </summary>
    /// <remarks>
    /// The owner's 2026-11-21 report is the reason: Fix refused a protected sheet outright, so the
    /// user had to unlock the cells by hand before the one command that exists to repair that workbook
    /// would run. The write is what proves the ordering — at the moment it runs, the release has
    /// happened and the re-protect has not.
    /// </remarks>
    [Fact]
    public void Repair_releases_a_foreign_protection_and_re_protects_around_the_catalogue_write()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.SetupSequence(g => g.Query())
            .Returns(ProtectionGuardOutcome.SheetProtected)
            .Returns(ProtectionGuardOutcome.NotProtected);
        FakeWorksheetProtectionPort port = new() { State = new WorksheetProtectionState(true, true) };
        string[] callsAtWrite = [];
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.Write())
            .Callback(() => callsAtWrite = [.. port.Calls])
            .Returns(ConfigWriteOutcome.Ok());

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object,
            protectionPort: port).Repair(Plan(ConfigIntegrityFindingKind.CatalogueHashMismatch), true);

        Assert.Null(outcome.Refusal);
        Assert.Equal(1, outcome.RepairedCount);
        Assert.Equal(["Read", "Unprotect"], callsAtWrite);
        Assert.Equal(["Read", "Unprotect", "Protect"], port.Calls);
        Assert.True(port.State.ContentsProtected);
        Assert.True(port.RepairCellsRequested);
        guard.Verify(g => g.Query(), Times.Exactly(2));
    }

    /// <summary>
    /// A protection the window cannot release — a password this add-in does not have — still refuses,
    /// and refuses before any write.
    /// </summary>
    [Fact]
    public void Repair_refuses_a_protection_it_cannot_release_before_calling_the_writer()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        FakeWorksheetProtectionPort port = new()
        {
            State = new WorksheetProtectionState(true, true),
            UnprotectRefusal = ProtectionWriteRefusal.HostRejected,
        };
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object,
            protectionPort: port).Repair(Plan(ConfigIntegrityFindingKind.CatalogueHashMismatch), true);

        Assert.Equal(ConfigRepairRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Equal(["Read", "Unprotect"], port.Calls);
        Assert.True(port.State.ContentsProtected);
        writer.Verify(w => w.Write(), Times.Never);
    }

    /// <summary>
    /// Workbook structure protection is not a worksheet protection the window can release, so it
    /// keeps the old refusal and costs no protection write.
    /// </summary>
    [Fact]
    public void Repair_refuses_workbook_structure_protection_before_calling_the_writer()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.WorkbookStructureProtected);
        FakeWorksheetProtectionPort port = new();
        var writer = new Mock<IConfigCatalogueWriter>(MockBehavior.Strict);

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object,
            protectionPort: port).Repair(Plan(ConfigIntegrityFindingKind.CatalogueHashMismatch), true);

        Assert.Equal(ConfigRepairRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(port.Calls);
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
    /// A workbook whose stored schema is a version behind is repaired by rewriting the
    /// catalogues, which is how a new metric token reaches an existing file.
    /// </summary>
    /// <remarks>
    /// The ribbon no longer carries an Initialise command, so a schema bump cannot tell the
    /// user to re-initialise: Repair is the only command they have. <c>SchemaVersionOlder</c>
    /// is a confirmation-required finding, so the rewrite happens only when the caller
    /// confirms - and that rewrite regenerates the catalogue rows from the code-owned
    /// definitions, which is what materialises a token a version-26 workbook has never
    /// seen. Both arms are asserted: skipping must not write, confirming must.
    /// </remarks>
    [Fact]
    public void Repair_rewrites_the_catalogues_for_an_older_schema_after_confirmation()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.Write()).Returns(ConfigWriteOutcome.Ok());

        ConfigRepairOutcome skipped = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object).Repair(Plan(ConfigIntegrityFindingKind.SchemaVersionOlder), confirmRequiredChanges: false);
        Assert.Null(skipped.Refusal);
        Assert.Equal(1, skipped.ConfirmationSkippedCount);
        writer.Verify(w => w.Write(), Times.Never);

        ConfigRepairOutcome confirmed = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object).Repair(Plan(ConfigIntegrityFindingKind.SchemaVersionOlder), confirmRequiredChanges: true);
        Assert.Null(confirmed.Refusal);
        Assert.Equal(1, confirmed.RepairedCount);
        writer.Verify(w => w.Write(), Times.Once);
    }

    /// <summary>
    /// ADR-0073: a one-step settings-only bump uses the narrow writer, not the
    /// full catalogue rewrite, so metric edits survive.
    /// </summary>
    [Fact]
    public void Repair_uses_the_narrow_writer_for_an_additive_schema_upgrade()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.UpgradeAdditiveSchema()).Returns(ConfigWriteOutcome.Ok());

        ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build(
        [
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.SchemaVersionOlderAdditive, "additive"),
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.CatalogueHashMismatch, "hash"),
        ]);

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object).Repair(plan, confirmRequiredChanges: false);

        Assert.Null(outcome.Refusal);
        Assert.Equal(1, outcome.RepairedCount);
        writer.Verify(w => w.UpgradeAdditiveSchema(), Times.Once);
        writer.Verify(w => w.Write(), Times.Never);
    }

    /// <summary>
    /// An additive upgrade that finds the workbook already current reports zero repairs.
    /// </summary>
    /// <remarks>
    /// The narrow upgrade is idempotent, so it succeeds without writing when the version has
    /// already moved on between the check and the repair. Counting that success told the user
    /// "1 repaired" for a no-op (owner report 2026-11-27), which is the same class of
    /// misleading sentence as the repair loop itself.
    /// </remarks>
    [Fact]
    public void Repair_counts_no_repair_for_an_additive_upgrade_that_wrote_nothing()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.UpgradeAdditiveSchema()).Returns(ConfigWriteOutcome.OkWithoutChanges());

        ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build(
        [
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.SchemaVersionOlderAdditive, "additive"),
            new ConfigIntegrityFinding(ConfigIntegrityFindingKind.CatalogueHashMismatch, "hash"),
        ]);

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object).Repair(plan, confirmRequiredChanges: false);

        Assert.Null(outcome.Refusal);
        Assert.Equal(0, outcome.RepairedCount);
        writer.Verify(w => w.Write(), Times.Never);
    }

    /// <summary>
    /// An unattributed catalogue fault is repairable, because refusing it made Fix a no-op.
    /// </summary>
    /// <remarks>
    /// The reader reports a table failure as a reason only ("RowCountMismatch"), and the
    /// integrity checker converts that to a <c>CatalogueTableCorrupt</c> finding with no
    /// table name. The plan classified every such finding as refused, so Repair returned a
    /// refused plan without writing anything - the owner's 2026-11-21 log shows five
    /// consecutive <c>config refused=RowCountMismatch</c> and a recreated sheet. This is that
    /// exact finding, nameless, and it must reach the catalogue rewrite.
    /// </remarks>
    [Fact]
    public void Repair_rewrites_the_catalogues_for_an_unattributed_catalogue_fault()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.Write()).Returns(ConfigWriteOutcome.Ok());

        ConfigRepairOutcome outcome = new ExcelConfigRepairer(
            null,
            writer.Object,
            null,
            guard.Object).Repair(Plan(ConfigIntegrityFindingKind.CatalogueTableCorrupt), confirmRequiredChanges: true);

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
