namespace GanttCreator.Core.Tests;

public class ConfigSheetLocatorTests
{
    private static ConfigSheetObservation Sheet(
        string name,
        bool veryHidden = false,
        int contractTables = 0) => new(name, veryHidden, contractTables);

    [Fact]
    public void A_workbook_with_no_configuration_candidate_is_missing()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate([Sheet("Gantt Data")]);

        Assert.Equal(ConfigSheetState.Missing, location.State);
        Assert.Null(location.SheetName);
    }

    [Fact]
    public void A_stripped_canonical_very_hidden_sheet_is_present_and_is_never_recreated()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [Sheet(GanttWorkbookContract.ConfigSheetName, veryHidden: true)]);

        Assert.Equal(ConfigSheetState.Present, location.State);
        Assert.Equal(GanttWorkbookContract.ConfigSheetName, location.SheetName);
    }

    [Fact]
    public void A_canonical_visible_sheet_carrying_contract_tables_is_present()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [Sheet(GanttWorkbookContract.ConfigSheetName, contractTables: 2)]);

        Assert.Equal(ConfigSheetState.Present, location.State);
    }

    [Fact]
    public void A_canonical_visible_sheet_with_no_contract_table_is_a_name_collision()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [Sheet(GanttWorkbookContract.ConfigSheetName)]);

        Assert.Equal(ConfigSheetState.NameCollision, location.State);
        Assert.False(location.NeedsRename);
    }

    [Fact]
    public void A_case_only_rename_is_accepted_without_needing_a_rename()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [Sheet("_ganttcreatorconfig", veryHidden: true, contractTables: 5)]);

        Assert.Equal(ConfigSheetState.Present, location.State);
        Assert.False(location.NeedsRename);
    }

    [Fact]
    public void Two_canonical_sheets_are_ambiguous()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [
                Sheet(GanttWorkbookContract.ConfigSheetName, veryHidden: true),
                Sheet(GanttWorkbookContract.ConfigSheetName.ToUpperInvariant(), contractTables: 1),
            ]);

        Assert.Equal(ConfigSheetState.Ambiguous, location.State);
        Assert.Null(location.SheetName);
    }

    [Fact]
    public void A_renamed_very_hidden_sheet_carrying_contract_tables_needs_a_rename()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [Sheet("ConfigBackup", veryHidden: true, contractTables: 1)]);

        Assert.Equal(ConfigSheetState.Renamed, location.State);
        Assert.Equal("ConfigBackup", location.SheetName);
        Assert.True(location.NeedsRename);
    }

    [Fact]
    public void Two_renamed_candidates_are_ambiguous()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [
                Sheet("ConfigBackup", veryHidden: true, contractTables: 1),
                Sheet("ConfigOld", veryHidden: true, contractTables: 2),
            ]);

        Assert.Equal(ConfigSheetState.Ambiguous, location.State);
    }

    [Fact]
    public void A_visible_sheet_carrying_a_contract_table_is_an_ambiguous_candidate()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [Sheet("Gantt Data", contractTables: 1)]);

        Assert.Equal(ConfigSheetState.AmbiguousCandidate, location.State);
    }

    [Fact]
    public void A_non_canonical_very_hidden_sheet_without_tables_is_unrecognised()
    {
        ConfigSheetLocation location = ConfigSheetLocator.Locate(
            [Sheet("Sheet2", veryHidden: true), Sheet("Gantt Data")]);

        Assert.Equal(ConfigSheetState.Unrecognised, location.State);
    }

    [Fact]
    public void The_contract_table_set_is_the_five_materialised_tables()
    {
        Assert.Equal(
            [
                GanttCatalogues.TypesTableName,
                GanttCatalogues.StylesTableName,
                GanttCatalogues.MetricsTableName,
                GanttCatalogues.SettingsTableName,
                GanttCatalogues.ConfigTableName,
            ],
            GanttCatalogues.ContractTableNames);
    }

    [Theory]
    [InlineData(GanttWorkbookContract.PlotAnchorDefinedName)]
    [InlineData("'Gantt Data'!GanttCreator.PlotAnchor")]
    [InlineData("ganttcreator.plotanchor")]
    public void The_plot_anchor_defined_name_matches_both_host_forms(string nameText) =>
        Assert.True(GanttWorkbookContract.IsPlotAnchorName(nameText));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("GanttCreator.TypeOptions")]
    [InlineData("'Gantt Data'!GanttCreator.PlotAnchorExtra")]
    public void Any_other_defined_name_is_not_the_plot_anchor(string? nameText) =>
        Assert.False(GanttWorkbookContract.IsPlotAnchorName(nameText));
}
