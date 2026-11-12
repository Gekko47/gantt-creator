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
    [InlineData(ConfigIntegrityFindingKind.DataTableOnUnexpectedSheet, ConfigRepairClassification.Refuse)]
    [InlineData(ConfigIntegrityFindingKind.ExtraTableColumns, ConfigRepairClassification.Refuse)]
    [InlineData(ConfigIntegrityFindingKind.TableOffContractOrigin, ConfigRepairClassification.Refuse)]
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

    [Fact]
    public void Empty_is_a_plan_with_no_findings_in_any_classification()
    {
        // The plan a clean workbook produces. It must be empty per classification
        // rather than merely empty overall, or "nothing to repair" and "nothing I
        // could classify" would look the same to a caller.
        ConfigIntegrityPlan plan = ConfigIntegrityPlan.Empty();

        Assert.Empty(plan.Findings);
        Assert.Empty(plan.AutoRepairFindings);
        Assert.Empty(plan.ConfirmRequiredFindings);
        Assert.Empty(plan.RefusedFindings);
        Assert.Empty(plan.QuarantinedFindings);
    }

    [Theory]
    [InlineData(GanttCatalogues.TypesTableName, ConfigRepairClassification.AutoRepair)]
    [InlineData(GanttCatalogues.MetricsTableName, ConfigRepairClassification.AutoRepair)]
    [InlineData(GanttCatalogues.StylesTableName, ConfigRepairClassification.ConfirmRequired)]
    public void Corrupt_catalogue_damage_is_classified_by_the_table_it_names(
        string tableName,
        ConfigRepairClassification expected)
    {
        // The code-owned catalogues can be rebuilt from Core, so they auto-repair;
        // the styles table carries user rows, so rebuilding it needs confirmation.
        var finding = new ConfigIntegrityFinding(
            ConfigIntegrityFindingKind.CatalogueTableCorrupt,
            "corrupt",
            tableName);

        Assert.Equal(expected, ConfigIntegrityPlan.Classify(finding));
    }

    [Fact]
    public void Build_orders_equal_findings_by_table_then_row_then_detail()
    {
        // The tie-breaks are why two runs over the same damage produce the same
        // Ribbon list. Ascending order and "a null row sorts last" are asserted
        // together, because reversing either one changes the order without changing
        // how many findings there are.
        ConfigIntegrityFinding[] input =
        [
            new(ConfigIntegrityFindingKind.IdentityDamage, "z-detail", "tblB"),
            new(ConfigIntegrityFindingKind.IdentityDamage, "a-detail", "tblA"),
            new(ConfigIntegrityFindingKind.IdentityDamage, "row-5", "tblA", 5),
            new(ConfigIntegrityFindingKind.IdentityDamage, "row-2", "tblA", 2),
            new(ConfigIntegrityFindingKind.IdentityDamage, "no-row-z", "tblA"),
        ];

        ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build(input);

        Assert.Equal(
            ["row-2", "row-5", "a-detail", "no-row-z", "z-detail"],
            plan.Findings.Select(finding => finding.Detail));
    }

    [Fact]
    public void Each_classification_property_returns_only_its_own_findings()
    {
        // These three properties are what the repairer and the Ribbon act on, so a
        // filter returning the wrong subset would repair damage the analyst has not
        // consented to. An empty plan exercises none of them; see
        // Empty_is_a_plan_with_no_findings_in_any_classification.
        ConfigIntegrityFinding auto = new(ConfigIntegrityFindingKind.WrongVisibility, "visibility");
        ConfigIntegrityFinding confirm = new(ConfigIntegrityFindingKind.IdentityDamage, "identity");
        ConfigIntegrityFinding refuse = new(ConfigIntegrityFindingKind.SecondHelperSheet, "helper");

        ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build([auto, confirm, refuse]);

        Assert.Same(auto, Assert.Single(plan.AutoRepairFindings));
        Assert.Same(confirm, Assert.Single(plan.ConfirmRequiredFindings));
        Assert.Same(refuse, Assert.Single(plan.RefusedFindings));
    }
}
