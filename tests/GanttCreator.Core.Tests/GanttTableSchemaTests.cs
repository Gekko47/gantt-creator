namespace GanttCreator.Core.Tests;

public class GanttTableSchemaTests
{
    private static readonly (string Name, bool IsRequired)[] s_columnContract =
    [
        // R5.18k: the visible columns lead, and the engine columns follow.
        // R5.21 (ADR-0064) adds the user-authored Activity ID directly after Type.
        ("Type", true), ("Activity ID", true), ("Description", true), ("Start", true),
        ("Finish", true), ("Duration", true),
        ("Id", true), ("LaneId", true), ("StackIndex", true),
        ("ParentId", true), ("SiblingOrder", true), ("StyleKey", true),
        ("LabelPosition", true), ("DateLabelPosition", false),
        ("FillColour", true), ("StrokeColour", true), ("Visible", true),
        ("SortOrder", false),
    ];

    [Fact]
    public void Default_schema_uses_the_contract_table_name()
    {
        Assert.Equal("tblGanttData", GanttTableSchema.TableName);
    }

    [Fact]
    public void Default_schema_lists_the_columns_in_contract_order()
    {
        var columns = GanttTableSchema.Default.Columns;

        // 16 required + two optional. R4.7A added SiblingOrder, taking
        // the required count from 13 to 14; R4.7C added Duration, taking
        // it to 15; R5.21 added the required Activity ID, taking it to 16.
        // R5.14 added DateLabelPosition as the second OPTIONAL
        // column (the SortOrder precedent): a workbook saved before
        // R5.14 legitimately lacks it, and a missing optional column
        // reads as a blank cell (= Auto) instead of failing the table
        // read, which would strand every pre-R5.14 workbook behind a
        // TableMissing refusal. R5.18k reordered the table — the visible
        // columns first, the engine columns after — without adding or
        // removing one, which is why the counts are unchanged and the
        // sequence is not.
        Assert.Equal(s_columnContract.Length, columns.Count);
        for (var i = 0; i < s_columnContract.Length; i++)
        {
            Assert.Equal(s_columnContract[i].Name, columns[i].Name);
            Assert.Equal(s_columnContract[i].IsRequired, columns[i].IsRequired);
        }
    }

    [Fact]
    public void Column_descriptors_round_trip_through_the_schema()
    {
        var rebuilt = new GanttTableSchema(
            GanttTableSchema.Default.Columns
                .Select(c => new GanttTableColumn(c.Name, c.IsRequired, c.Access))
                .ToList());

        Assert.Equal(
            GanttTableSchema.Default.Columns.Select(c => (c.Name, c.IsRequired)),
            rebuilt.Columns.Select(c => (c.Name, c.IsRequired)));
    }

    [Fact]
    public void Every_column_has_exactly_one_classification()
    {
        // ADR-0029 D8 makes the classification a contract rather than a comment:
        // exhaustiveness is what stops a new column from shipping unclassified,
        // and the enum's closed value set is what stops it shipping with a
        // classification that nothing acts on.
        var all = Enum.GetValues<GanttColumnAccess>();

        foreach (GanttTableColumn column in GanttTableSchema.Default.Columns)
        {
            Assert.Contains(column.Access, all);
        }

        // All three classes are actually used; an unused class would mean the
        // classification has a state no column can reach, which is dead contract.
        Assert.Equal(3, all.Length);
        Assert.Equal(
            all.OrderBy(access => access),
            GanttTableSchema.Default.Columns.Select(column => column.Access).Distinct().OrderBy(access => access));
    }

    [Theory]
    [InlineData("Type", GanttColumnAccess.Authoring)]
    [InlineData("Activity ID", GanttColumnAccess.Authoring)]
    [InlineData("Description", GanttColumnAccess.Authoring)]
    [InlineData("Start", GanttColumnAccess.Authoring)]
    [InlineData("Finish", GanttColumnAccess.Authoring)]
    [InlineData("Duration", GanttColumnAccess.ReadOnlyVisible)]
    [InlineData("Id", GanttColumnAccess.EngineHidden)]
    [InlineData("ParentId", GanttColumnAccess.EngineHidden)]
    [InlineData("SiblingOrder", GanttColumnAccess.EngineHidden)]
    [InlineData("LaneId", GanttColumnAccess.EngineHidden)]
    [InlineData("StackIndex", GanttColumnAccess.EngineHidden)]
    [InlineData("StyleKey", GanttColumnAccess.EngineHidden)]
        [InlineData("DateLabelPosition", GanttColumnAccess.EngineHidden)]
    [InlineData("LabelPosition", GanttColumnAccess.EngineHidden)]
    [InlineData("FillColour", GanttColumnAccess.EngineHidden)]
    [InlineData("StrokeColour", GanttColumnAccess.EngineHidden)]
    [InlineData("Visible", GanttColumnAccess.EngineHidden)]
    [InlineData("SortOrder", GanttColumnAccess.EngineHidden)]
    public void D1_classification_membership_is_exact(string columnName, GanttColumnAccess expected)
    {
        // The D1 list is transcribed column-by-column rather than derived from the
        // schema, so a column moved between classes fails here rather than passing
        // because both sides changed together.
        Assert.True(GanttTableSchema.Default.TryGetColumn(columnName, out GanttTableColumn? column));
        Assert.NotNull(column);
        Assert.Equal(expected, column.Access);
    }

    [Fact]
    public void Hidden_and_locked_derive_from_the_classification_without_disagreement()
    {
        // Only engine columns are hidden; only authoring columns are unlocked.
        // Deriving both from the enum is what keeps the initialiser's two writes
        // from contradicting each other.
        foreach (GanttTableColumn column in GanttTableSchema.Default.Columns)
        {
            Assert.Equal(column.Access == GanttColumnAccess.EngineHidden, column.IsHidden);
            Assert.Equal(column.Access != GanttColumnAccess.Authoring, column.IsLocked);
        }
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
            new GanttTableColumn("Id", isRequired: true, GanttColumnAccess.EngineHidden),
            new GanttTableColumn("Id", isRequired: true, GanttColumnAccess.EngineHidden)));
    }

    [Fact]
    public void Schema_constructor_throws_for_an_empty_column_name()
    {
        Assert.Throws<ArgumentException>(() => Schema(new GanttTableColumn("", isRequired: true, GanttColumnAccess.Authoring)));
    }

    [Fact]
    public void Schema_constructor_throws_for_a_whitespace_column_name()
    {
        Assert.Throws<ArgumentException>(() => Schema(new GanttTableColumn("   ", isRequired: true, GanttColumnAccess.Authoring)));
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

    [Fact]
    public void The_date_columns_are_the_two_authoring_date_columns_in_the_schema()
    {
        // The initialiser and the recovery both format this set, so every name in it
        // must be a real column: a typo would silently stop a date column being
        // formatted, and a name that resolved to an engine column would format a column
        // the user never sees.
        Assert.Equal(2, GanttTableSchema.DateColumnNames.Count);
        Assert.Equal("Start", GanttTableSchema.DateColumnNames[0]);
        Assert.Equal("Finish", GanttTableSchema.DateColumnNames[1]);

        foreach (string name in GanttTableSchema.DateColumnNames)
        {
            Assert.True(GanttTableSchema.Default.TryGetColumn(name, out GanttTableColumn? column));
            Assert.Equal(GanttColumnAccess.Authoring, column!.Access);
        }
    }

    [Fact]
    public void Only_the_schema_named_date_columns_are_dates()
    {
        // Ordinal, like every other header lookup in this schema: a differently-cased
        // header is not the contract's column, and a blank or absent name is no column
        // at all. Both refusals are the positive case for the guard.
        Assert.True(GanttTableSchema.IsDateColumn("Start"));
        Assert.True(GanttTableSchema.IsDateColumn("Finish"));
        Assert.False(GanttTableSchema.IsDateColumn("start"));
        Assert.False(GanttTableSchema.IsDateColumn("Duration"));
        Assert.False(GanttTableSchema.IsDateColumn(string.Empty));
        Assert.False(GanttTableSchema.IsDateColumn(null));
    }

    private static GanttTableSchema Schema(params GanttTableColumn[] columns) => new(columns);
}
