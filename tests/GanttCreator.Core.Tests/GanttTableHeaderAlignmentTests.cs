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

    [Fact]
    public void Plan_writes_only_the_blank_cell_when_the_rest_aligns()
    {
        string?[] live = [.. SchemaNames()];
        live[4] = null;

        GanttTableHeaderAlignmentPlan plan = GanttTableHeaderAlignment.Plan(live);

        Assert.Equal(GanttTableHeaderAlignmentOutcome.Writable, plan.Outcome);
        Assert.Equal([4], plan.CellsToWrite);
    }

    [Fact]
    public void Plan_refuses_a_renamed_header()
    {
        string?[] live = [.. SchemaNames()];
        live[4] = "Notes";

        GanttTableHeaderAlignmentPlan plan = GanttTableHeaderAlignment.Plan(live);

        Assert.Equal(GanttTableHeaderAlignmentOutcome.Misaligned, plan.Outcome);
        Assert.False(plan.Succeeded);
        Assert.Equal("Description", plan.MismatchedColumn);
        Assert.Empty(plan.CellsToWrite);
    }

    [Fact]
    public void Structure_plan_inserts_a_deleted_required_column_at_its_schema_position()
    {
        List<string> live = [.. SchemaNames()];
        live.RemoveAt(3);

        GanttTableStructureRepairPlan plan = GanttTableStructureRepair.Plan(live);

        Assert.True(plan.Succeeded);
        GanttTableColumnRepair repair = Assert.Single(plan.Repairs);
        Assert.Equal(GanttTableColumnRepairKind.Insert, repair.Kind);
        Assert.Equal(3, repair.SchemaIndex);
        Assert.Equal("Type", repair.Name);
    }

    [Fact]
    public void Structure_plan_renames_a_single_unknown_header_in_place()
    {
        List<string> live = [.. SchemaNames()];
        live[4] = "Notes";

        GanttTableStructureRepairPlan plan = GanttTableStructureRepair.Plan(live);

        Assert.True(plan.Succeeded);
        GanttTableColumnRepair repair = Assert.Single(plan.Repairs);
        Assert.Equal(GanttTableColumnRepairKind.Rename, repair.Kind);
        Assert.Equal(4, repair.SchemaIndex);
        Assert.Equal("Description", repair.Name);
    }

    [Fact]
    public void Structure_plan_refuses_a_reorder_combined_with_a_deletion()
    {
        List<string> live = [.. SchemaNames()];
        live.RemoveAt(3);
        (live[3], live[4]) = (live[4], live[3]);

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
