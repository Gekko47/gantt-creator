namespace GanttCreator.Core.Tests;

public class GanttTableHeaderAlignmentTests
{
    private static string[] SchemaNames() =>
        [.. GanttTableSchema.Default.Columns.Select(column => column.Name)];

    [Fact]
    public void Plan_reports_an_aligned_header_row_without_writes()
    {
        GanttTableHeaderAlignmentPlan plan = GanttTableHeaderAlignment.Plan(SchemaNames());

        Assert.Equal(GanttTableHeaderAlignmentOutcome.Aligned, plan.Outcome);
        Assert.Empty(plan.CellsToWrite);
        Assert.True(plan.Succeeded);
    }

    [Fact]
    public void Plan_writes_every_cell_when_the_header_row_is_blank()
    {
        string?[] live = new string?[GanttTableSchema.Default.Columns.Count];

        GanttTableHeaderAlignmentPlan plan = GanttTableHeaderAlignment.Plan(live);

        Assert.Equal(GanttTableHeaderAlignmentOutcome.Writable, plan.Outcome);
        Assert.Equal(live.Length, plan.CellsToWrite.Count);
    }

    /// <summary>
    /// The schema position of a column by name. The tests below are about how a plan
    /// reacts to a damaged header row, not about where the schema puts a column: R5.18k
    /// reordered the columns, and the literal indices these tests used to carry turned a
    /// schema decision into a red test with no behavioural question behind it.
    /// </summary>
    /// <param name="name">The exact column name.</param>
    /// <returns>The zero-based schema position.</returns>
    private static int SchemaIndexOf(string name) => Array.IndexOf(SchemaNames(), name);

    [Fact]
    public void Plan_writes_only_the_blank_cell_when_the_rest_aligns()
    {
        int description = SchemaIndexOf("Description");
        string?[] live = [.. SchemaNames()];
        live[description] = null;

        GanttTableHeaderAlignmentPlan plan = GanttTableHeaderAlignment.Plan(live);

        Assert.Equal(GanttTableHeaderAlignmentOutcome.Writable, plan.Outcome);
        Assert.Equal([description], plan.CellsToWrite);
    }

    [Fact]
    public void Plan_refuses_a_renamed_header()
    {
        string?[] live = [.. SchemaNames()];
        live[SchemaIndexOf("Description")] = "Notes";

        GanttTableHeaderAlignmentPlan plan = GanttTableHeaderAlignment.Plan(live);

        Assert.Equal(GanttTableHeaderAlignmentOutcome.Misaligned, plan.Outcome);
        Assert.False(plan.Succeeded);
        Assert.Equal("Description", plan.MismatchedColumn);
        Assert.Empty(plan.CellsToWrite);
    }

    [Fact]
    public void Structure_plan_inserts_a_deleted_required_column_at_its_schema_position()
    {
        int type = SchemaIndexOf("Type");
        List<string> live = [.. SchemaNames()];
        live.RemoveAt(type);

        GanttTableStructureRepairPlan plan = GanttTableStructureRepair.Plan(live);

        Assert.True(plan.Succeeded);
        GanttTableColumnRepair repair = Assert.Single(plan.Repairs);
        Assert.Equal(GanttTableColumnRepairKind.Insert, repair.Kind);
        Assert.Equal(type, repair.SchemaIndex);
        Assert.Equal("Type", repair.Name);
    }

    [Fact]
    public void Structure_plan_renames_a_single_unknown_header_in_place()
    {
        int description = SchemaIndexOf("Description");
        List<string> live = [.. SchemaNames()];
        live[description] = "Notes";

        GanttTableStructureRepairPlan plan = GanttTableStructureRepair.Plan(live);

        Assert.True(plan.Succeeded);
        GanttTableColumnRepair repair = Assert.Single(plan.Repairs);
        Assert.Equal(GanttTableColumnRepairKind.Rename, repair.Kind);
        Assert.Equal(description, repair.SchemaIndex);
        Assert.Equal("Description", repair.Name);
    }

    [Fact]
    public void Structure_plan_refuses_a_reorder_combined_with_a_deletion()
    {
        int type = SchemaIndexOf("Type");
        List<string> live = [.. SchemaNames()];
        live.RemoveAt(type);
        (live[type], live[type + 1]) = (live[type + 1], live[type]);

        GanttTableStructureRepairPlan plan = GanttTableStructureRepair.Plan(live);

        Assert.False(plan.Succeeded);
        Assert.Equal(GanttTableStructureRefusal.ColumnOrderShifted, plan.Refusal);
    }

    [Fact]
    public void Structure_plan_refuses_extra_columns_with_a_missing_required_column()
    {
        List<string> live = [.. SchemaNames()];
        live.RemoveAt(3);
        live.Add("UserColumn");
        live.Add("AnotherUserColumn");

        GanttTableStructureRepairPlan plan = GanttTableStructureRepair.Plan(live);

        Assert.False(plan.Succeeded);
        Assert.Equal(GanttTableStructureRefusal.ExtraColumnsPresent, plan.Refusal);
    }

    [Fact]
    public void Structure_plan_reports_a_readable_table_as_inconsistent()
    {
        GanttTableStructureRepairPlan plan = GanttTableStructureRepair.Plan(SchemaNames());

        Assert.False(plan.Succeeded);
        Assert.Equal(GanttTableStructureRefusal.Inconsistent, plan.Refusal);
    }
}
