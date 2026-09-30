namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="RowHeightNormaliser"/> (R4.7D, ADR-0026 D3/D4): the plan
/// that restores a dragged managed row, and every refusal.
/// </summary>
public sealed class RowHeightNormaliserTests
{
    private const double ManagedPt = 18;
    private const double SplitterPt = 24;
    private const double SpacerPt = 6;

    /// <summary>
    /// A correctly normalised sheet produces an EMPTY write set. This is the case
    /// that keeps normalisation free on the common path: the adapter must not touch
    /// a row that already agrees with the token, or every Refresh would dirty the
    /// workbook.
    /// </summary>
    [Fact]
    public void An_already_normalised_sheet_plans_no_writes()
    {
        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan(
            [new MeasuredRowHeight(2, ManagedPt), new MeasuredRowHeight(3, ManagedPt)],
            ManagedPt,
            SplitterPt,
            SpacerPt);

        Assert.NotNull(plan);
        Assert.Equal(0, plan.WriteCount);
        Assert.All(plan.Rows, row => Assert.False(row.NeedsWrite));
    }

    /// <summary>
    /// A managed row the user has dragged is restored to the token. This is the
    /// behaviour the whole type exists for.
    /// </summary>
    [Fact]
    public void A_dragged_managed_row_is_restored()
    {
        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan(
            [new MeasuredRowHeight(2, 45)],
            ManagedPt,
            SplitterPt,
            SpacerPt);

        Assert.NotNull(plan);
        RowHeightNormalisation row = Assert.Single(plan.Rows);
        Assert.True(row.NeedsWrite);
        Assert.Equal(ManagedPt, row.TargetHeightPt);
    }

    /// <summary>
    /// Structural rows follow their own tokens, not the managed height. A splitter
    /// taken to the managed height would collapse the section header onto an
    /// ordinary bar, and a spacer at 18pt would open a gap as tall as an activity.
    /// </summary>
    [Fact]
    public void Structural_rows_follow_their_own_tokens()
    {
        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan(
            [
                new MeasuredRowHeight(2, 100, MeasuredRowKind.Splitter),
                new MeasuredRowHeight(3, 100, MeasuredRowKind.Spacer),
            ],
            ManagedPt,
            SplitterPt,
            SpacerPt);

        Assert.NotNull(plan);
        Assert.Equal(SplitterPt, plan.Rows[0].TargetHeightPt);
        Assert.Equal(SpacerPt, plan.Rows[1].TargetHeightPt);
        Assert.All(plan.Rows, row => Assert.True(row.IsStructuralRow));
    }

    /// <summary>
    /// A structural row already at its own token is a no-op even though it differs
    /// from the managed height. "Equal to the token" is the test, not "equal to the
    /// managed height".
    /// </summary>
    [Fact]
    public void A_structural_row_already_at_its_token_is_a_no_op()
    {
        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan(
            [new MeasuredRowHeight(2, SplitterPt, MeasuredRowKind.Splitter)],
            ManagedPt,
            SplitterPt,
            SpacerPt);

        Assert.NotNull(plan);
        Assert.Equal(0, plan.WriteCount);
    }

    /// <summary>
    /// Rows come back in worksheet order whatever order they were supplied in, so
    /// the plan and the write set are deterministic.
    /// </summary>
    [Fact]
    public void The_plan_is_in_worksheet_row_order()
    {
        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan(
            [new MeasuredRowHeight(5, 99), new MeasuredRowHeight(2, 99), new MeasuredRowHeight(3, 99)],
            ManagedPt,
            SplitterPt,
            SpacerPt);

        Assert.NotNull(plan);
        Assert.Equal([2, 3, 5], plan.Rows.Select(static row => row.RowNumber));
    }

    /// <summary>
    /// Excel reports row heights with float noise, so a row a hair off the token must
    /// not be rewritten. Rewriting it would dirty the workbook on every Refresh and
    /// defeat the no-op path above.
    /// </summary>
    [Fact]
    public void A_height_within_tolerance_is_left_alone()
    {
        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan(
            [new MeasuredRowHeight(2, ManagedPt + 0.01)],
            ManagedPt,
            SplitterPt,
            SpacerPt);

        Assert.NotNull(plan);
        Assert.Equal(0, plan.WriteCount);
    }

    /// <summary>
    /// A row just outside the tolerance IS rewritten. The boundary matters in both
    /// directions: too loose and a real drag survives, too tight and float noise
    /// causes a permanent rewrite.
    /// </summary>
    [Fact]
    public void A_height_outside_tolerance_is_rewritten()
    {
        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan(
            [new MeasuredRowHeight(2, ManagedPt + 0.5)],
            ManagedPt,
            SplitterPt,
            SpacerPt);

        Assert.NotNull(plan);
        Assert.Equal(1, plan.WriteCount);
    }

    /// <summary>
    /// A non-positive target is refused. Writing it would produce the misalignment
    /// normalisation exists to prevent, so this is a refusal rather than a default.
    /// </summary>
    [Theory]
    [InlineData(0, 24, 6)]
    [InlineData(-18, 24, 6)]
    [InlineData(18, 0, 6)]
    [InlineData(18, 24, -6)]
    [InlineData(double.NaN, 24, 6)]
    public void A_non_positive_target_is_refused(double managed, double splitter, double spacer)
    {
        Assert.Null(RowHeightNormaliser.Plan([new MeasuredRowHeight(2, 45)], managed, splitter, spacer));
    }

    /// <summary>
    /// An unrecognised row kind is refused rather than defaulted to the managed
    /// height. <c>Kind</c> can arrive from an out-of-range cast, and silently
    /// writing the managed height for a row the caller could not classify would be
    /// exactly the wrong-row-height-wrong-lane defect this work removes.
    /// </summary>
    [Fact]
    public void An_unknown_row_kind_is_refused()
    {
        Assert.Null(
            RowHeightNormaliser.Plan(
                [new MeasuredRowHeight(2, 45, (MeasuredRowKind)99)],
                ManagedPt,
                SplitterPt,
                SpacerPt));
    }

    /// <summary>
    /// A null row list is a programming error, and throws rather than returning null
    /// so it cannot be confused with a refusal.
    /// </summary>
    [Fact]
    public void A_null_row_list_throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => RowHeightNormaliser.Plan(null!, ManagedPt, SplitterPt, SpacerPt));
    }

    /// <summary>
    /// An empty table plans an empty write set successfully. A table with no rows is
    /// a legitimate state, not a refusal.
    /// </summary>
    [Fact]
    public void An_empty_table_plans_no_writes()
    {
        RowHeightNormalisationPlan? plan = RowHeightNormaliser.Plan([], ManagedPt, SplitterPt, SpacerPt);

        Assert.NotNull(plan);
        Assert.Empty(plan.Rows);
        Assert.Equal(0, plan.WriteCount);
        Assert.Equal(ManagedPt, plan.ManagedHeightPt);
    }
}
