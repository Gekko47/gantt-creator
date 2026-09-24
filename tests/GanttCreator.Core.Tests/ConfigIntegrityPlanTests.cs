using GanttCreator.Core.ConfigIntegrity;

namespace GanttCreator.Core.Tests;

public sealed class ConfigIntegrityPlanTests
{
    [Theory]
    [InlineData(ConfigIntegrityFindingKind.WrongVisibility, ConfigRepairClassification.AutoRepair)]
    [InlineData(ConfigIntegrityFindingKind.CatalogueHashMismatch, ConfigRepairClassification.AutoRepair)]
    [InlineData(ConfigIntegrityFindingKind.PlotAnchorDisagreement, ConfigRepairClassification.AutoRepair)]
    [InlineData(ConfigIntegrityFindingKind.TypeOptionsMissing, ConfigRepairClassification.AutoRepair)]
    [InlineData(ConfigIntegrityFindingKind.SchemaVersionOlder, ConfigRepairClassification.ConfirmRequired)]
    [InlineData(ConfigIntegrityFindingKind.IdentityDamage, ConfigRepairClassification.ConfirmRequired)]
    [InlineData(ConfigIntegrityFindingKind.SchemaVersionNewer, ConfigRepairClassification.Refuse)]
    [InlineData(ConfigIntegrityFindingKind.SecondHelperSheet, ConfigRepairClassification.Refuse)]
    [InlineData(ConfigIntegrityFindingKind.ScheduleDataOnConfig, ConfigRepairClassification.Refuse)]
    [InlineData(ConfigIntegrityFindingKind.Unknown, ConfigRepairClassification.Refuse)]
    public void Classify_maps_each_damage_class(
        ConfigIntegrityFindingKind kind,
        ConfigRepairClassification expected)
    {
        var finding = new ConfigIntegrityFinding(kind, "test", "tblGanttTypes");

        Assert.Equal(expected, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void Classify_refuses_a_second_helper_sheet_even_with_no_table_context()
    {
        var finding = new ConfigIntegrityFinding(
            ConfigIntegrityFindingKind.SecondHelperSheet,
            "extra helper");

        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void Classify_refuses_unknown_configuration_table_damage()
    {
        var finding = new ConfigIntegrityFinding(
            ConfigIntegrityFindingKind.CatalogueTableCorrupt,
            "unknown table",
            "tblUnknown");

        Assert.Equal(ConfigRepairClassification.Refuse, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void Classify_marks_code_owned_missing_table_auto_repair_and_style_missing_confirm()
    {
        var types = new ConfigIntegrityFinding(
            ConfigIntegrityFindingKind.CatalogueTableMissing,
            "types missing",
            GanttCatalogues.TypesTableName);
        var styles = new ConfigIntegrityFinding(
            ConfigIntegrityFindingKind.CatalogueTableMissing,
            "styles missing",
            GanttCatalogues.StylesTableName);

        Assert.Equal(ConfigRepairClassification.AutoRepair, ConfigIntegrityPlan.Classify(types));
        Assert.Equal(ConfigRepairClassification.ConfirmRequired, ConfigIntegrityPlan.Classify(styles));
    }

    [Fact]
    public void Build_orders_findings_deterministically()
    {
        ConfigIntegrityFinding[] input =
        [
            new(ConfigIntegrityFindingKind.TypeOptionsMissing, "type options"),
            new(ConfigIntegrityFindingKind.WrongVisibility, "visibility"),
            new(ConfigIntegrityFindingKind.CatalogueHashMismatch, "hash"),
        ];

        ConfigIntegrityPlan first = ConfigIntegrityPlan.Build(input);
        ConfigIntegrityPlan second = ConfigIntegrityPlan.Build(input.Reverse());

        Assert.Equal(first.Findings, second.Findings);
        Assert.Equal(
            [
                ConfigIntegrityFindingKind.WrongVisibility,
                ConfigIntegrityFindingKind.CatalogueHashMismatch,
                ConfigIntegrityFindingKind.TypeOptionsMissing,
            ],
            first.Findings.Select(finding => finding.Kind));
    }

    [Fact]
    public void Build_keeps_quarantine_findings_visible()
    {
        var quarantined = new ConfigIntegrityFinding(
            ConfigIntegrityFindingKind.UserStyleQuarantined,
            "row 2",
            GanttCatalogues.StylesTableName,
            2);

        ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build([quarantined]);

        Assert.Equal(quarantined, Assert.Single(plan.QuarantinedFindings));
        Assert.Equal(ConfigRepairClassification.ConfirmRequired, ConfigIntegrityPlan.Classify(quarantined));
    }

    [Fact]
    public void Build_rejects_null_findings()
    {
        Assert.Throws<ArgumentNullException>(() => ConfigIntegrityPlan.Build(null!));
    }

    [Fact]
    public void Classify_rejects_null_finding()
    {
        Assert.Throws<ArgumentNullException>(() => ConfigIntegrityPlan.Classify(null!));
    }
}
