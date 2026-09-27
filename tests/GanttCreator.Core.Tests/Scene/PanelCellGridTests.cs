using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Contract tests for the R3.11 measured data-panel grid (entity guide §3/§4).
/// </summary>
/// <remarks>
/// Every guard has a positive test that constructs the bad input and asserts the
/// exact refusal, so none of these validators can ship silently inert
/// (AGENTS.md).
/// </remarks>
public sealed class PanelCellGridTests
{
    private static readonly string[] Required = ["Id", "Type", "Description", "Start"];

    private static PanelCellGrid Grid(params PanelColumn[] columns) =>
        PanelCellGrid.TryCreate(columns, 12.0, []).Grid
        ?? throw new InvalidOperationException("Fixture grid should be valid.");

    private static PanelColumn Col(string name, double width = 40.0) => new(name, width);

    [Fact]
    public void A_valid_grid_keeps_column_order_widths_and_height()
    {
        // Order is the caller's and is significant: §4's header follows the
        // schema display names, and the cell bounds depend on the order.
        PanelCellGrid grid = Grid(Col("Type", 30.0), Col("Id", 50.0), Col("Description", 120.0));

        Assert.Equal(["Type", "Id", "Description"], grid.Columns.Select(c => c.Name));
        Assert.Equal(200.0, grid.TotalWidthPt);
        Assert.Equal(12.0, grid.RowHeightPt);
    }

    [Fact]
    public void Columns_default_to_left_alignment_and_may_be_overridden()
    {
        PanelCellGrid grid = Grid(Col("Type"), Col("Id") with { Alignment = GanttTextAlignment.Centre });

        Assert.Equal(GanttTextAlignment.Left, grid.Columns[0].Alignment);
        Assert.Equal(GanttTextAlignment.Centre, grid.Columns[1].Alignment);
    }

    [Fact]
    public void A_null_column_list_is_refused()
    {
        Assert.Equal(
            PanelCellGridRefusal.NullRequest,
            PanelCellGrid.TryCreate(null, 12.0, Required).Refusal);
    }

    [Fact]
    public void An_empty_column_list_is_refused()
    {
        // Positive test: a panel with no columns would emit a header with no
        // cells and a body with no content, which §3 does not permit.
        Assert.Equal(
            PanelCellGridRefusal.NoColumns,
            PanelCellGrid.TryCreate([], 12.0, Required).Refusal);
    }

    [Fact]
    public void A_non_positive_or_non_finite_width_is_refused()
    {
        // Positive test for §3's "positive visible widths" rule.
        foreach (double width in new[] { 0.0, -1.0, double.NaN, double.PositiveInfinity })
        {
            Assert.Equal(
                PanelCellGridRefusal.NonPositiveWidth,
                PanelCellGrid.TryCreate([Col("Id", width)], 12.0, Required).Refusal);
        }
    }

    [Fact]
    public void A_non_positive_or_non_finite_row_height_is_refused()
    {
        foreach (double height in new[] { 0.0, -3.0, double.NaN, double.PositiveInfinity })
        {
            Assert.Equal(
                PanelCellGridRefusal.NonPositiveRowHeight,
                PanelCellGrid.TryCreate([Col("Id")], height, Required).Refusal);
        }
    }

    [Fact]
    public void A_blank_column_name_is_refused()
    {
        foreach (string name in new[] { "", "   " })
        {
            Assert.Equal(
                PanelCellGridRefusal.BlankColumnName,
                PanelCellGrid.TryCreate([Col(name)], 12.0, Required).Refusal);
        }
    }

    [Fact]
    public void A_duplicate_column_name_is_refused()
    {
        // Positive test for §4's unique header names and §3's "required columns
        // once": two identical names would emit duplicate header text at
        // different bounds, so it is refused rather than collapsed.
        Assert.Equal(
            PanelCellGridRefusal.DuplicateColumn,
            PanelCellGrid.TryCreate([Col("Id", 30.0), Col("Id", 40.0)], 12.0, Required).Refusal);
    }

    [Fact]
    public void A_missing_required_column_is_refused()
    {
        // Positive test: the grid must carry every required schema column, so a
        // panel missing Description cannot build an export composition.
        Assert.Equal(
            PanelCellGridRefusal.MissingRequiredColumn,
            PanelCellGrid.TryCreate([Col("Id"), Col("Type")], 12.0, Required).Refusal);
    }

    [Fact]
    public void A_blank_required_name_is_refused()
    {
        Assert.Equal(
            PanelCellGridRefusal.MissingRequiredColumn,
            PanelCellGrid.TryCreate([Col("Id")], 12.0, [" "]).Refusal);
    }

    [Fact]
    public void An_undefined_alignment_is_refused()
    {
        Assert.Equal(
            PanelCellGridRefusal.UndefinedAlignment,
            PanelCellGrid.TryCreate([Col("Id") with { Alignment = (GanttTextAlignment)99 }], 12.0, Required)
                .Refusal);
    }

    [Fact]
    public void Column_name_matching_is_ordinal_and_does_not_case_fold()
    {
        // "ID" is not "Id": case-folding would silently satisfy a required
        // column that the live sheet does not actually have.
        Assert.Equal(
            PanelCellGridRefusal.MissingRequiredColumn,
            PanelCellGrid.TryCreate([Col("ID"), Col("Type"), Col("Description"), Col("Start")], 12.0, Required)
                .Refusal);
    }

    [Fact]
    public void A_grid_with_no_required_columns_is_accepted()
    {
        PanelCellGridCreationOutcome outcome = PanelCellGrid.TryCreate([Col("Id")], 12.0, []);

        Assert.NotNull(outcome.Grid);
        Assert.Empty(outcome.Grid.RequiredColumns);
    }

    [Fact]
    public void The_schema_display_names_are_valid_grid_column_names()
    {
        // §4's header text comes from the schema, so every schema column must be
        // usable as a grid column name without tripping the blank-name guard.
        foreach (GanttTableColumn column in GanttTableSchema.Default.Columns)
        {
            PanelCellGrid grid = Grid(Col(column.Name));

            Assert.Equal(column.Name, grid.Columns[0].Name);
        }
    }
}
