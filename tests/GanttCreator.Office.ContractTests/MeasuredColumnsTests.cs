using GanttCreator.Core;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The live panel measures what the sheet SHOWS, so the column set it measures is the
/// schema's visible columns minus the ones the user hid from the ribbon.
/// </summary>
/// <remarks>
/// Reported from the host on 2026-11-14: hiding a column made every Refresh fail with "The
/// worksheet columns could not be measured." The measured set came from the schema alone, a
/// hidden worksheet column measures as width 0, and <see cref="PanelCellGrid"/> correctly
/// refuses a non-positive width — so the refusal was right and the input was wrong. The
/// three states are pinned here because the third one is a silent fallback and would
/// otherwise never be exercised.
/// </remarks>
public class MeasuredColumnsTests
{
    private static IReadOnlyList<string> SchemaVisibleColumns =>
        [.. GanttTableSchema.Default.Columns.Where(column => !column.IsHidden).Select(column => column.Name)];

    [Fact]
    public void With_no_stored_choice_the_schema_visible_columns_are_measured()
    {
        IReadOnlyList<string> measured =
            ExcelSceneBuildRequestFactory.MeasuredColumnsFor(new Dictionary<string, string>());

        Assert.Equal(SchemaVisibleColumns, measured);
    }

    [Fact]
    public void A_column_the_user_hid_is_not_measured()
    {
        Dictionary<string, string> settings = new(StringComparer.Ordinal)
        {
            [GanttColumnVisibility.SettingKey] = "Duration",
        };

        IReadOnlyList<string> measured = ExcelSceneBuildRequestFactory.MeasuredColumnsFor(settings);

        Assert.DoesNotContain("Duration", measured);
        Assert.Contains("Type", measured);

        // Schema order survives, because PanelCellGrid treats column order as significant.
        Assert.Equal(
            [.. SchemaVisibleColumns.Where(name => !string.Equals(name, "Duration", StringComparison.Ordinal))],
            measured);
    }

    [Fact]
    public void Two_hidden_columns_are_both_excluded_and_only_those_two()
    {
        Dictionary<string, string> settings = new(StringComparer.Ordinal)
        {
            [GanttColumnVisibility.SettingKey] = "Start,Finish",
        };

        IReadOnlyList<string> measured = ExcelSceneBuildRequestFactory.MeasuredColumnsFor(settings);

        Assert.DoesNotContain("Start", measured);
        Assert.DoesNotContain("Finish", measured);
        Assert.Equal(SchemaVisibleColumns.Count - 2, measured.Count);
    }

    [Fact]
    public void A_malformed_stored_list_falls_back_to_the_schema_visible_columns()
    {
        // GanttColumnVisibility refuses a name it does not know rather than reading the
        // value as "nothing hidden" — that refusal belongs to the dropdown that wrote it.
        // A render only needs a width, so it measures the schema's set instead of failing.
        Dictionary<string, string> settings = new(StringComparer.Ordinal)
        {
            [GanttColumnVisibility.SettingKey] = "NotAColumn",
        };

        IReadOnlyList<string> measured = ExcelSceneBuildRequestFactory.MeasuredColumnsFor(settings);

        Assert.Equal(SchemaVisibleColumns, measured);
    }
}
