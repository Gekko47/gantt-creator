using GanttCreator.Core;

namespace GanttCreator.Core.Tests;

/// <summary>Contract tests for the canonical R2.8 scaffold rows.</summary>
public class GanttRowDefaultsTests
{
    public static TheoryData<GanttEntityType, string, string> ExpectedScaffolds => new()
    {
        { GanttEntityType.AsPlannedActivity, "As-Planned Activity", "AsPlannedActivity" },
        { GanttEntityType.AsPlannedMilestone, "As-Planned Milestone", "AsPlannedMilestone" },
        { GanttEntityType.Delineator, "Delineator", "DefaultDelineator" },
    };

    [Fact]
    public void Build_rejects_a_null_generator()
    {
        Assert.Throws<ArgumentNullException>(
            () => GanttRowDefaults.Build(GanttEntityType.AsPlannedActivity, null!));
    }

    [Fact]
    public void Build_rejects_a_type_outside_the_catalogue()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => GanttRowDefaults.Build((GanttEntityType)999, GanttRowId.New));
    }

    [Theory]
    [MemberData(nameof(ExpectedScaffolds))]
    public void Build_uses_the_exact_catalogue_scaffold(
        GanttEntityType type,
        string expectedType,
        string expectedStyleKey)
    {
        GanttRowId id = GanttRowId.Parse("G-0123456789abcdef0123456789abcdef");

        IReadOnlyList<object?> actual = GanttRowDefaults.Build(type, () => id);

        // One value per schema column, in schema order.
        Assert.Equal(GanttTableSchema.Default.Columns.Count, actual.Count);

        // Read by COLUMN NAME, not by index. The previous version of this test
        // pinned StyleKey at index 8, which is ParentId's position -- and became
        // SiblingOrder's position when R4.7A inserted that column. The test
        // therefore passed for a builder that had been writing the style key into
        // the wrong cell. Looking the value up by name is the contract the caller
        // depends on, and it cannot silently pass on a shifted literal.
        var byName = GanttTableSchema.Default.Columns
            .Select((column, index) => (column.Name, Value: actual[index]))
            .ToDictionary(p => p.Name, p => p.Value, StringComparer.Ordinal);

        Assert.Equal(id.Value, byName["Id"]);
        Assert.Equal(expectedType, byName["Type"]);
        Assert.Equal(expectedStyleKey, byName["StyleKey"]);

        // Everything else is blank so validation reports the required authoring inputs.
        foreach (string column in GanttTableSchema.Default.Columns.Select(c => c.Name))
        {
            if (column is "Id" or "Type" or "StyleKey")
            {
                continue;
            }

            Assert.Null(byName[column]);
        }
    }

    /// <summary>
    /// The guard that would have caught the real defect: every populated value
    /// must land in the column it names. A scaffold that shifts when the schema
    /// grows writes a style key into the sibling-ordering column, which is
    /// invisible until a rendered row carries a nonsense ordering.
    /// </summary>
    [Fact]
    public void Every_populated_scaffold_value_lands_in_its_own_column()
    {
        IReadOnlyList<object?> actual = GanttRowDefaults.Build(GanttEntityType.AsPlannedActivity, GanttRowId.New);
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;

        var populated = columns
            .Select((column, index) => (column.Name, Value: actual[index]))
            .Where(p => p.Value is not null)
            .Select(p => p.Name)
            .ToHashSet(StringComparer.Ordinal);

        Assert.Equal(
            new HashSet<string>(["Id", "Type", "StyleKey"], StringComparer.Ordinal),
            populated);
    }
}
