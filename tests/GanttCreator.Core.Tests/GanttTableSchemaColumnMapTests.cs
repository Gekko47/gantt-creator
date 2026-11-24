namespace GanttCreator.Core.Tests;

public class GanttTableSchemaColumnMapTests
{
    private static string[] DefaultHeaders() =>
        [.. GanttTableSchema.Default.Columns.Select(column => column.Name)];

    [Fact]
    public void Exact_headers_map_every_schema_column_to_its_one_based_index()
    {
        string[] headers = DefaultHeaders();

        Assert.True(GanttTableSchema.TryBuildColumnMap(headers, out int[]? map));
        int[] resolved = map!;
        Assert.Equal(headers.Length, resolved.Length);
        for (int index = 0; index < resolved.Length; index++)
        {
            Assert.Equal(index + 1, resolved[index]);
        }
    }

    [Fact]
    public void Reordered_headers_map_by_name()
    {
        string[] headers = [.. DefaultHeaders().Reverse()];
        int count = headers.Length;

        Assert.True(GanttTableSchema.TryBuildColumnMap(headers, out int[]? map));
        int[] resolved = map!;
        Assert.Equal(GanttTableSchema.Default.Columns.Count, resolved.Length);
        for (int index = 0; index < resolved.Length; index++)
        {
            Assert.Equal(count - index, resolved[index]);
        }
    }

    [Fact]
    public void Missing_required_column_refuses_with_a_null_map()
    {
        string[] headers = [.. DefaultHeaders().Where(name => name != "Type")];

        Assert.False(GanttTableSchema.TryBuildColumnMap(headers, out int[]? map));
        Assert.Null(map);
    }

    [Fact]
    public void Missing_optional_column_maps_to_minus_one()
    {
        string[] headers = [.. DefaultHeaders().Where(name => name != "DateLabelPosition")];

        Assert.True(GanttTableSchema.TryBuildColumnMap(headers, out int[]? map));
        int[] resolved = map!;
        Assert.Equal(-1, resolved[SchemaIndexOf("DateLabelPosition")]);
        Assert.Equal(1, resolved[SchemaIndexOf("Type")]);
    }

    [Fact]
    public void Unknown_header_is_ignored()
    {
        string[] headers = [.. DefaultHeaders(), "Notes"];

        Assert.True(GanttTableSchema.TryBuildColumnMap(headers, out int[]? map));
        int[] resolved = map!;
        Assert.Equal(1, resolved[SchemaIndexOf("Type")]);
        Assert.Equal(SchemaIndexOf("SortOrder") + 1, resolved[SchemaIndexOf("SortOrder")]);
    }

    [Fact]
    public void Duplicate_header_keeps_the_last_index()
    {
        string[] headers = [.. DefaultHeaders(), "Id"];

        Assert.True(GanttTableSchema.TryBuildColumnMap(headers, out int[]? map));
        int[] resolved = map!;
        Assert.Equal(headers.Length, resolved[SchemaIndexOf("Id")]);
    }

    [Fact]
    public void Empty_header_list_refuses()
    {
        Assert.False(GanttTableSchema.TryBuildColumnMap([], out int[]? map));
        Assert.Null(map);
    }

    [Fact]
    public void Null_header_list_throws()
    {
        Assert.Throws<ArgumentNullException>(() => GanttTableSchema.TryBuildColumnMap(null!, out _));
    }

    [Fact]
    public void Case_mismatch_does_not_resolve()
    {
        string[] headers = DefaultHeaders();
        headers[Array.IndexOf(headers, "Type")] = "type";

        Assert.False(GanttTableSchema.TryBuildColumnMap(headers, out int[]? map));
        Assert.Null(map);
    }

    [Fact]
    public void Null_and_empty_header_entries_are_skipped()
    {
        string[] headers = [.. DefaultHeaders(), null!, string.Empty];

        Assert.True(GanttTableSchema.TryBuildColumnMap(headers, out int[]? map));
        int[] resolved = map!;
        Assert.Equal(1, resolved[SchemaIndexOf("Type")]);
    }

    private static int SchemaIndexOf(string name) =>
        GanttTableSchema.Default.Columns
            .Select((column, index) => (column.Name, index))
            .Single(entry => string.Equals(entry.Name, name, StringComparison.Ordinal))
            .index;
}
