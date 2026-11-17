namespace GanttCreator.Core.Tests;

/// <summary>
/// R5.17e slice 2: the workbook's own names are compared in one place.
/// </summary>
/// <remarks>
/// Every case is derived from the contract constant, so renaming the sheet or the
/// label fails these tests instead of leaving them green against a name the
/// product no longer writes. The plot-anchor predicate's own tests remain in
/// <see cref="ConfigSheetLocatorTests"/>, where they were written.
/// </remarks>
public class GanttWorkbookContractTests
{
    [Fact]
    public void The_contract_name_is_the_configuration_sheet()
    {
        Assert.True(GanttWorkbookContract.IsConfigSheetName(GanttWorkbookContract.ConfigSheetName));
    }

    [Fact]
    public void An_upper_case_configuration_name_is_recognised()
    {
        string variant = GanttWorkbookContract.ConfigSheetName.ToUpperInvariant();

        Assert.NotEqual(GanttWorkbookContract.ConfigSheetName, variant, StringComparer.Ordinal);
        Assert.True(GanttWorkbookContract.IsConfigSheetName(variant));
    }

    [Fact]
    public void The_gantt_label_is_the_gantt_sheet_label()
    {
        Assert.True(GanttWorkbookContract.IsGanttSheetLabel(GanttWorkbookContract.GanttSheetLabel));
    }

    [Fact]
    public void An_upper_case_gantt_label_is_recognised()
    {
        string variant = GanttWorkbookContract.GanttSheetLabel.ToUpperInvariant();

        Assert.NotEqual(GanttWorkbookContract.GanttSheetLabel, variant, StringComparer.Ordinal);
        Assert.True(GanttWorkbookContract.IsGanttSheetLabel(variant));
    }

    [Fact]
    public void A_collision_suffixed_label_is_not_the_standard_label()
    {
        // Excel's own collision convention appends " (2)". Such a sheet is a Gantt
        // sheet for the reader's purposes only if it also carries the table, so the
        // label predicate must not claim it.
        Assert.False(GanttWorkbookContract.IsGanttSheetLabel(GanttWorkbookContract.GanttSheetLabel + " (2)"));
    }

    [Fact]
    public void A_renamed_configuration_sheet_is_not_the_configuration_name()
    {
        Assert.False(GanttWorkbookContract.IsConfigSheetName(GanttWorkbookContract.ConfigSheetName + "_old"));
    }

    [Fact]
    public void The_two_sheet_names_are_not_each_other()
    {
        // A guard against a copy-paste that points both predicates at one constant.
        Assert.False(GanttWorkbookContract.IsConfigSheetName(GanttWorkbookContract.GanttSheetLabel));
        Assert.False(GanttWorkbookContract.IsGanttSheetLabel(GanttWorkbookContract.ConfigSheetName));
    }

    [Fact]
    public void A_null_name_is_neither_sheet()
    {
        Assert.False(GanttWorkbookContract.IsConfigSheetName(null));
        Assert.False(GanttWorkbookContract.IsGanttSheetLabel(null));
    }

    [Fact]
    public void An_empty_name_is_neither_sheet()
    {
        Assert.False(GanttWorkbookContract.IsConfigSheetName(string.Empty));
        Assert.False(GanttWorkbookContract.IsGanttSheetLabel(string.Empty));
    }
}
