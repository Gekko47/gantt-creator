namespace GanttCreator.Core.Tests;

/// <summary>
/// R5.17e: one predicate owns the "is this the Gantt data table?" comparison.
/// </summary>
/// <remarks>
/// Every case is derived from <see cref="GanttTableSchema.TableName"/> rather than
/// a copied literal, so a rename of the table fails these tests loudly instead of
/// leaving them green against a name the product no longer uses.
/// </remarks>
public class GanttChartContractTests
{
    [Fact]
    public void The_contract_name_is_the_data_table()
    {
        Assert.True(GanttChartContract.IsDataTableName(GanttTableSchema.TableName));
        Assert.True(GanttChartContract.IsDataTableNameExact(GanttTableSchema.TableName));
    }

    [Fact]
    public void An_upper_case_variant_is_the_data_table_under_the_case_insensitive_rule()
    {
        string variant = GanttTableSchema.TableName.ToUpperInvariant();

        Assert.NotEqual(GanttTableSchema.TableName, variant, StringComparer.Ordinal);
        Assert.True(GanttChartContract.IsDataTableName(variant));
    }

    [Fact]
    public void An_upper_case_variant_is_not_the_data_table_under_the_exact_rule()
    {
        string variant = GanttTableSchema.TableName.ToUpperInvariant();

        Assert.False(GanttChartContract.IsDataTableNameExact(variant));
    }

    [Theory]
    [InlineData("2")]
    [InlineData(" ")]
    [InlineData("Other")]
    public void A_longer_table_name_is_not_the_data_table(string suffix)
    {
        string candidate = GanttTableSchema.TableName + suffix;

        Assert.False(GanttChartContract.IsDataTableName(candidate));
        Assert.False(GanttChartContract.IsDataTableNameExact(candidate));
    }

    [Fact]
    public void An_unrelated_table_name_is_not_the_data_table()
    {
        Assert.False(GanttChartContract.IsDataTableName("SomeOtherTable"));
        Assert.False(GanttChartContract.IsDataTableNameExact("SomeOtherTable"));
    }

    [Fact]
    public void A_prefix_of_the_contract_name_is_not_the_data_table()
    {
        string candidate = GanttTableSchema.TableName[..^1];

        Assert.False(GanttChartContract.IsDataTableName(candidate));
        Assert.False(GanttChartContract.IsDataTableNameExact(candidate));
    }

    [Fact]
    public void A_null_name_is_never_the_data_table()
    {
        Assert.False(GanttChartContract.IsDataTableName(null));
        Assert.False(GanttChartContract.IsDataTableNameExact(null));
    }

    [Fact]
    public void An_empty_name_is_never_the_data_table()
    {
        Assert.False(GanttChartContract.IsDataTableName(string.Empty));
        Assert.False(GanttChartContract.IsDataTableNameExact(string.Empty));
    }
}
