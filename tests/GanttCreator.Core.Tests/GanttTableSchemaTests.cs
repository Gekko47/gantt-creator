namespace GanttCreator.Core.Tests;

public class GanttTableSchemaTests
{
    private static readonly string[] s_requiredColumnNames =
    [
        "Id", "LaneId", "StackIndex", "Type", "Description", "Start", "Finish",
        "ParentId", "StyleKey", "LabelPosition", "FillColour", "StrokeColour", "Visible",
    ];

    [Fact]
    public void Default_schema_uses_the_contract_table_name()
    {
        Assert.Equal("tblGanttData", GanttTableSchema.TableName);
    }

    [Fact]
    public void Default_schema_lists_the_13_required_columns_in_contract_order()
    {
        var columns = GanttTableSchema.Default.Columns;

        Assert.Equal(14, columns.Count);
        for (var i = 0; i < s_requiredColumnNames.Length; i++)
        {
            Assert.Equal(s_requiredColumnNames[i], columns[i].Name);
            Assert.True(columns[i].IsRequired);
        }

        var last = columns[^1];
        Assert.Equal("SortOrder", last.Name);
        Assert.False(last.IsRequired);
    }

    [Fact]
    public void Column_descriptors_round_trip_through_the_schema()
    {
        var rebuilt = new GanttTableSchema(
            GanttTableSchema.Default.Columns
                .Select(c => new GanttTableColumn(c.Name, c.IsRequired))
                .ToList());

        Assert.Equal(
            GanttTableSchema.Default.Columns.Select(c => (c.Name, c.IsRequired)),
            rebuilt.Columns.Select(c => (c.Name, c.IsRequired)));
    }

    [Fact]
    public void TryGetColumn_resolves_exact_names_only()
    {
        Assert.True(GanttTableSchema.Default.TryGetColumn("Type", out var typeColumn));
        Assert.NotNull(typeColumn);
        Assert.True(typeColumn.IsRequired);

        Assert.False(GanttTableSchema.Default.TryGetColumn("type", out _));
        Assert.False(GanttTableSchema.Default.TryGetColumn("ColumnType", out _));
        Assert.False(GanttTableSchema.Default.TryGetColumn("", out _));
        Assert.False(GanttTableSchema.Default.TryGetColumn(null, out _));
    }

    [Fact]
    public void Schema_constructor_throws_for_duplicate_column_names()
    {
        Assert.Throws<ArgumentException>(() => Schema(
            new GanttTableColumn("Id", isRequired: true),
            new GanttTableColumn("Id", isRequired: true)));
    }

    [Fact]
    public void Schema_constructor_throws_for_an_empty_column_name()
    {
        Assert.Throws<ArgumentException>(() => Schema(new GanttTableColumn("", isRequired: true)));
    }

    [Fact]
    public void Schema_constructor_throws_for_a_whitespace_column_name()
    {
        Assert.Throws<ArgumentException>(() => Schema(new GanttTableColumn("   ", isRequired: true)));
    }

    [Fact]
    public void Schema_constructor_throws_for_an_empty_column_list()
    {
        Assert.Throws<ArgumentException>(() => Schema());
    }

    [Fact]
    public void Schema_constructor_throws_for_a_null_column_list()
    {
        Assert.Throws<ArgumentNullException>(() => new GanttTableSchema(null!));
    }

    [Fact]
    public void Schema_constructor_throws_for_a_null_column()
    {
        Assert.Throws<ArgumentNullException>(() => new GanttTableSchema([null!]));
    }

    private static GanttTableSchema Schema(params GanttTableColumn[] columns) => new(columns);
}
