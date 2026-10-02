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
        PanelCellGrid.TryCreate(columns, [12.0], 12.0, []).Grid
        ?? throw new InvalidOperationException("Fixture grid should be valid.");

    private static PanelColumn Col(string name, double width = 40.0) => new(name, width);

    /// <summary>
    /// The measured absolute origin is carried on the grid (ADR-0030 D3).
    /// </summary>
    /// <remarks>
    /// This is the input the live chart's vertical origin was missing. Before it, the
    /// plot's top came from the size preset's page coordinates and
    /// <c>LaneLayoutBuilder</c> stacked lanes from there, so a lane could not coincide
    /// with its own row — ADR-0026 D3's <c>Excel Top == Scene lane Top</c> was not
    /// merely unmet but unrepresentable.
    /// </remarks>
    [Fact]
    public void The_measured_origin_is_carried_on_the_grid()
    {
        PanelCellGrid grid =
            PanelCellGrid.TryCreate([Col("Id")], [12.0], 12.0, [], originTopPt: 64.5, originLeftPt: 8.25).Grid
            ?? throw new InvalidOperationException("Fixture grid should be valid.");

        Assert.Equal(64.5, grid.OriginTopPt);
        Assert.Equal(8.25, grid.OriginLeftPt);
    }

    /// <summary>
    /// A zero origin is accepted: it is the sheet's own top-left, a real position.
    /// </summary>
    /// <remarks>
    /// Stated because the opposite assumption is the defect this slice removes — a
    /// zero origin is only wrong when it is a <em>default</em> standing in for an
    /// unmeasured live sheet, which the adapter now prevents by refusing when the
    /// host reports no figure. An export composition legitimately starts at zero.
    /// </remarks>
    [Fact]
    public void A_zero_origin_is_accepted_as_a_real_position()
    {
        PanelCellGrid grid =
            PanelCellGrid.TryCreate([Col("Id")], [12.0], 12.0, [], originTopPt: 0, originLeftPt: 0).Grid
            ?? throw new InvalidOperationException("Fixture grid should be valid.");

        Assert.Equal(0d, grid.OriginTopPt);
        Assert.Equal(0d, grid.OriginLeftPt);
    }

    /// <summary>
    /// A negative or non-finite origin is refused as <c>InvalidOrigin</c>.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator added with ADR-0030 D3.</b> A guessed
    /// origin would place lanes at a plausible but wrong vertical position with
    /// nothing to report it, which is the silence this refusal exists to prevent.
    /// </remarks>
    [Theory]
    [InlineData(-1d, 0d)]
    [InlineData(0d, -0.5d)]
    [InlineData(double.NaN, 0d)]
    [InlineData(0d, double.PositiveInfinity)]
    public void An_unusable_origin_is_refused(double originTopPt, double originLeftPt)
    {
        PanelCellGridCreationOutcome outcome =
            PanelCellGrid.TryCreate([Col("Id")], [12.0], 12.0, [], originTopPt, originLeftPt);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PanelCellGridRefusal.InvalidOrigin, outcome.Refusal);
    }

    /// <summary>
    /// The measured padding-row heights are carried on the grid (ADR-0031 D2).
    /// </summary>
    /// <remarks>
    /// They are the chart frame's top and bottom margin, so they are part of the
    /// measured input rather than a constant the factory could substitute. The two
    /// are asserted with DIFFERENT values on purpose: assuming the bottom row
    /// matches the top one is the bug this guards, and a fixture that used one
    /// number for both could not tell the two apart.
    /// </remarks>
    [Fact]
    public void The_measured_padding_heights_are_carried_on_the_grid()
    {
        PanelCellGrid grid =
            PanelCellGrid.TryCreate(
                [Col("Id")],
                [12.0],
                12.0,
                [],
                topPaddingHeightPt: 18.0,
                bottomPaddingHeightPt: 27.5).Grid
            ?? throw new InvalidOperationException("Fixture grid should be valid.");

        Assert.Equal(18.0, grid.TopPaddingHeightPt);
        Assert.Equal(27.5, grid.BottomPaddingHeightPt);
    }

    /// <summary>
    /// A padding height that is not finite, or is negative, is refused.
    /// </summary>
    /// <remarks>
    /// <b>Positive test for the validator added with ADR-0031 D2.</b> A negative
    /// height would place the chart frame's edge inside its own content, producing
    /// an inverted chart that no later stage would catch. Each row is a separate
    /// case because a check that only read the top one would accept a bad bottom.
    /// </remarks>
    [Theory]
    [InlineData(-1d, 0d)]
    [InlineData(0d, -0.5d)]
    [InlineData(double.NaN, 0d)]
    [InlineData(0d, double.PositiveInfinity)]
    public void An_unusable_padding_height_is_refused(double top, double bottom)
    {
        PanelCellGridCreationOutcome outcome = PanelCellGrid.TryCreate(
            [Col("Id")],
            [12.0],
            12.0,
            [],
            topPaddingHeightPt: top,
            bottomPaddingHeightPt: bottom);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PanelCellGridRefusal.InvalidPaddingHeight, outcome.Refusal);
    }

    /// <summary>
    /// Zero padding on both rows is accepted: the export case, where there are no
    /// padding rows at all.
    /// </summary>
    /// <remarks>
    /// The counterweight to the refusal theory above. A strictly-positive check would
    /// refuse the legitimate zero that <c>TryCreate</c>'s own defaults supply, so the
    /// whole export path would break.
    /// </remarks>
    [Fact]
    public void Zero_padding_is_accepted()
    {
        PanelCellGridCreationOutcome outcome = PanelCellGrid.TryCreate([Col("Id")], [12.0], 12.0, []);

        Assert.True(outcome.Succeeded);
        Assert.Equal(0d, outcome.Grid!.TopPaddingHeightPt);
        Assert.Equal(0d, outcome.Grid.BottomPaddingHeightPt);
    }

    [Fact]
    public void A_valid_grid_keeps_column_order_widths_and_height()
    {
        // Order is the caller's and is significant: §4's header follows the
        // schema display names, and the cell bounds depend on the order.
        PanelCellGrid grid = Grid(Col("Type", 30.0), Col("Id", 50.0), Col("Description", 120.0));

        Assert.Equal(["Type", "Id", "Description"], grid.Columns.Select(c => c.Name));
        Assert.Equal(200.0, grid.TotalWidthPt);
        Assert.Equal([12.0], grid.RowHeightsPt);
        Assert.Equal(12.0, grid.HeaderHeightPt);
    }

    [Fact]
    public void A_mixed_height_body_is_accepted_and_totalled_exactly()
    {
        // The whole point of the change: a worksheet body need not be uniform, so the
        // grid carries a list and the panel bottom is the exact sum. A single sample
        // height could only refuse this table or lay it out wrong.
        PanelCellGrid grid =
            PanelCellGrid.TryCreate([Col("Id")], [12.0, 24.0, 15.0], 18.0, []).Grid
            ?? throw new InvalidOperationException("Fixture grid should be valid.");

        Assert.Equal([12.0, 24.0, 15.0], grid.RowHeightsPt);
        Assert.Equal(51.0, grid.TotalRowHeightPt);
        Assert.Equal(18.0, grid.HeaderHeightPt);
    }

    [Fact]
    public void A_missing_or_empty_row_height_list_is_refused()
    {
        // Positive test: with no measured heights there is no body to reproduce, and
        // defaulting a height would place every cell at a guessed position.
        Assert.Equal(
            PanelCellGridRefusal.NoRows,
            PanelCellGrid.TryCreate([Col("Id")], null, 12.0, Required).Refusal);
        Assert.Equal(
            PanelCellGridRefusal.NoRows,
            PanelCellGrid.TryCreate([Col("Id")], [], 12.0, Required).Refusal);
    }

    [Fact]
    public void A_non_positive_or_non_finite_header_height_is_refused()
    {
        // §4's header follows the live header-cell bounds, so the header height is a
        // measurement in its own right and cannot be defaulted to a body height.
        foreach (double height in new[] { 0.0, -1.0, double.NaN, double.PositiveInfinity })
        {
            Assert.Equal(
                PanelCellGridRefusal.NonPositiveHeaderHeight,
                PanelCellGrid.TryCreate([Col("Id")], [12.0], height, Required).Refusal);
        }
    }

    [Fact]
    public void One_bad_height_among_good_ones_is_refused()
    {
        // The guard is per row, not on an aggregate: an adapter that reads a mixed
        // range could otherwise return one good height and one absent one.
        Assert.Equal(
            PanelCellGridRefusal.NegativeRowHeight,
            PanelCellGrid.TryCreate([Col("Id")], [12.0, -1.0], 12.0, Required).Refusal);
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
            PanelCellGrid.TryCreate(null, [12.0], 12.0, Required).Refusal);
    }

    [Fact]
    public void An_empty_column_list_is_refused()
    {
        // Positive test: a panel with no columns would emit a header with no
        // cells and a body with no content, which §3 does not permit.
        Assert.Equal(
            PanelCellGridRefusal.NoColumns,
            PanelCellGrid.TryCreate([], [12.0], 12.0, Required).Refusal);
    }

    [Fact]
    public void A_non_positive_or_non_finite_width_is_refused()
    {
        // Positive test for §3's "positive visible widths" rule.
        foreach (double width in new[] { 0.0, -1.0, double.NaN, double.PositiveInfinity })
        {
            Assert.Equal(
                PanelCellGridRefusal.NonPositiveWidth,
                PanelCellGrid.TryCreate([Col("Id", width)], [12.0], 12.0, Required).Refusal);
        }
    }

    [Fact]
    public void A_negative_or_non_finite_row_height_is_refused()
    {
        foreach (double height in new[] { -3.0, double.NaN, double.PositiveInfinity })
        {
            Assert.Equal(
                PanelCellGridRefusal.NegativeRowHeight,
                PanelCellGrid.TryCreate([Col("Id")], [height], 12.0, []).Refusal);
        }
    }

    /// <summary>
    /// A hidden row measures zero height (ADR-0034): Excel reports it as zero
    /// and the rows below it share its top, so zero is a real measurement.
    /// </summary>
    [Fact]
    public void A_zero_row_height_is_accepted_as_a_hidden_row()
    {
        PanelCellGridCreationOutcome outcome =
            PanelCellGrid.TryCreate([Col("Id")], [0.0], 12.0, []);

        Assert.True(outcome.Succeeded);
    }

    [Fact]
    public void A_blank_column_name_is_refused()
    {
        foreach (string name in new[] { "", "   " })
        {
            Assert.Equal(
                PanelCellGridRefusal.BlankColumnName,
                PanelCellGrid.TryCreate([Col(name)], [12.0], 12.0, Required).Refusal);
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
            PanelCellGrid.TryCreate([Col("Id", 30.0), Col("Id", 40.0)], [12.0], 12.0, Required).Refusal);
    }

    [Fact]
    public void A_missing_required_column_is_refused()
    {
        // Positive test: the grid must carry every required schema column, so a
        // panel missing Description cannot build an export composition.
        Assert.Equal(
            PanelCellGridRefusal.MissingRequiredColumn,
            PanelCellGrid.TryCreate([Col("Id"), Col("Type")], [12.0], 12.0, Required).Refusal);
    }

    [Fact]
    public void A_blank_required_name_is_refused()
    {
        Assert.Equal(
            PanelCellGridRefusal.MissingRequiredColumn,
            PanelCellGrid.TryCreate([Col("Id")], [12.0], 12.0, [" "]).Refusal);
    }

    [Fact]
    public void An_undefined_alignment_is_refused()
    {
        Assert.Equal(
            PanelCellGridRefusal.UndefinedAlignment,
            PanelCellGrid.TryCreate([Col("Id") with { Alignment = (GanttTextAlignment)99 }], [12.0], 12.0, Required)
                .Refusal);
    }

    [Fact]
    public void Column_name_matching_is_ordinal_and_does_not_case_fold()
    {
        // "ID" is not "Id": case-folding would silently satisfy a required
        // column that the live sheet does not actually have.
        Assert.Equal(
            PanelCellGridRefusal.MissingRequiredColumn,
            PanelCellGrid.TryCreate([Col("ID"), Col("Type"), Col("Description"), Col("Start")], [12.0], 12.0, Required)
                .Refusal);
    }

    [Fact]
    public void A_grid_with_no_required_columns_is_accepted()
    {
        PanelCellGridCreationOutcome outcome = PanelCellGrid.TryCreate([Col("Id")], [12.0], 12.0, []);

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
