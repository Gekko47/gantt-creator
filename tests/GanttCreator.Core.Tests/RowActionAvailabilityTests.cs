using System.Globalization;

namespace GanttCreator.Core.Tests;

/// <summary>
/// R5.9b: the structural row-action rules the ribbon disables against. Every
/// flag mirrors a refusal the promote, demote or move command already produces.
/// </summary>
public class RowActionAvailabilityTests
{
    // The table cell holds the catalogue DISPLAY name, not the enum name, so
    // the fixtures read the same text a workbook would.
    private static string ActivityType => TypeText(GanttEntityType.AsPlannedActivity);

    private static string MilestoneType => TypeText(GanttEntityType.AsPlannedMilestone);

    private static string SpacerType => TypeText(GanttEntityType.Spacer);

    private static string CriticalIntervalType => TypeText(GanttEntityType.CriticalInterval);

    private static string TypeText(GanttEntityType type) =>
        EntityTypeCatalog.GetDefinition(type)!.DisplayName;

    [Fact]
    public void First_body_row_offers_down_but_not_up()
    {
        GanttRowDto[] rows = [Row(1, 1, ActivityType), Row(2, 2, ActivityType)];

        RowActionAvailability availability = RowActionAvailability.Evaluate(rows, 1);

        Assert.False(availability.CanMoveUp);
        Assert.True(availability.CanMoveDown);
    }

    [Fact]
    public void Last_body_row_offers_up_but_not_down()
    {
        GanttRowDto[] rows = [Row(1, 1, ActivityType), Row(2, 2, ActivityType)];

        RowActionAvailability availability = RowActionAvailability.Evaluate(rows, 2);

        Assert.True(availability.CanMoveUp);
        Assert.False(availability.CanMoveDown);
    }

    [Fact]
    public void A_top_level_row_cannot_be_promoted()
    {
        GanttRowDto[] rows = [Row(1, 1, ActivityType), Row(2, 2, ActivityType)];

        Assert.False(RowActionAvailability.Evaluate(rows, 1).CanPromote);
    }

    [Fact]
    public void A_child_that_owns_children_cannot_be_promoted()
    {
        // Invalid data (depth 3), but the predicate must still not offer a
        // promotion that would orphan the grandchild.
        GanttRowDto[] rows =
        [
            Row(1, 1, ActivityType),
            Row(2, 2, ActivityType, parent: Id(1)),
            Row(3, 3, ActivityType, parent: Id(2)),
        ];

        Assert.False(RowActionAvailability.Evaluate(rows, 2).CanPromote);
    }

    [Fact]
    public void A_child_with_no_children_can_be_promoted()
    {
        GanttRowDto[] rows =
        [
            Row(1, 1, ActivityType),
            Row(2, 2, ActivityType, parent: Id(1)),
        ];

        Assert.True(RowActionAvailability.Evaluate(rows, 2).CanPromote);
    }

    [Fact]
    public void A_critical_interval_cannot_be_promoted()
    {
        // Critical Interval is child-only by contract, so promoting it is the
        // PromotionForbiddenForType refusal.
        GanttRowDto[] rows =
        [
            Row(1, 1, ActivityType),
            Row(2, 2, CriticalIntervalType, parent: Id(1)),
        ];

        Assert.False(RowActionAvailability.Evaluate(rows, 2).CanPromote);
    }

    [Fact]
    public void The_first_child_of_a_group_offers_no_up()
    {
        // The parent is the row above, but it is a different scope, so the
        // child would move nowhere: the boundary is the run, not the sheet.
        GanttRowDto[] rows =
        [
            Row(1, 1, ActivityType),
            Row(2, 2, ActivityType, parent: Id(1)),
            Row(3, 3, ActivityType, parent: Id(1)),
        ];

        RowActionAvailability availability = RowActionAvailability.Evaluate(rows, 2);

        Assert.False(availability.CanMoveUp);
        Assert.True(availability.CanMoveDown);
    }

    [Fact]
    public void A_top_level_span_row_below_an_eligible_parent_can_be_demoted()
    {
        GanttRowDto[] rows = [Row(1, 1, ActivityType), Row(2, 2, ActivityType)];

        Assert.True(RowActionAvailability.Evaluate(rows, 2).CanDemote);
    }

    [Fact]
    public void A_row_that_cannot_be_a_child_cannot_be_demoted()
    {
        GanttRowDto[] rows = [Row(1, 1, ActivityType), Row(2, 2, SpacerType)];

        Assert.False(RowActionAvailability.Evaluate(rows, 2).CanDemote);
    }

    [Fact]
    public void No_eligible_parent_above_means_no_demote()
    {
        // A milestone cannot own children, so nothing above can accept the row.
        GanttRowDto[] rows = [Row(1, 1, MilestoneType), Row(2, 2, ActivityType)];

        Assert.False(RowActionAvailability.Evaluate(rows, 2).CanDemote);
    }

    [Fact]
    public void An_unknown_type_reports_none()
    {
        GanttRowDto[] rows = [Row(1, 1, "NotAType"), Row(2, 2, ActivityType)];

        Assert.Equal(RowActionAvailability.None, RowActionAvailability.Evaluate(rows, 1));
    }

    [Fact]
    public void An_unparsable_parent_id_reports_none()
    {
        GanttRowDto[] rows = [Row(1, 1, ActivityType), Row(2, 2, ActivityType, parent: "G-not-an-id")];

        Assert.Equal(RowActionAvailability.None, RowActionAvailability.Evaluate(rows, 2));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0)]
    [InlineData(3)]
    public void An_index_outside_the_body_reports_none(int? index)
    {
        GanttRowDto[] rows = [Row(1, 1, ActivityType), Row(2, 2, ActivityType)];

        Assert.Equal(RowActionAvailability.None, RowActionAvailability.Evaluate(rows, index));
    }

    [Fact]
    public void Eligible_parent_agrees_with_the_planner_rule_body()
    {
        // Parity: the ribbon's demote gate asks IsEligibleParent, and the
        // inserter and the planner ask Evaluate. The two must answer the same
        // question the same way for every candidate, or one surface accepts a
        // parent the other refuses.
        GanttRowId parentId = GanttRowId.Parse(Id(50));

        (GanttEntityType Type, GanttRowId? ParentOfParent, int ChildCount)[] candidates =
        [
            (GanttEntityType.AsPlannedActivity, null, 0),
            (GanttEntityType.AsPlannedActivity, null, EntityHierarchyCatalog.MaxChildrenPerParent),
            (GanttEntityType.AsPlannedActivity, parentId, 0),
            (GanttEntityType.AsPlannedMilestone, null, 0),
            (GanttEntityType.Spacer, null, 0),
        ];

        foreach ((GanttEntityType type, GanttRowId? parentOfParent, int childCount) in candidates)
        {
            bool eligible = AddChildPlanner.IsEligibleParent(type, parentOfParent, childCount);
            bool accepted = AddChildPlanner.Evaluate(
                type,
                parentId,
                parentOfParent,
                childCount,
                GanttEntityType.AsPlannedActivity) is null;

            Assert.Equal(eligible, accepted);
        }
    }

    private static GanttRowDto Row(int rowNumber, int id, string type, string? parent = null) =>
        new(
            rowNumber,
            Id(id),
            laneId: null,
            stackIndex: null,
            typeText: type,
            description: null,
            start: null,
            finish: null,
            parentId: parent,
            styleKey: null,
            labelPositionText: null,
            fillColourText: null,
            strokeColourText: null,
            visible: null,
            sortOrder: null);

    private static string Id(int value) =>
        "G-" + value.ToString("x32", CultureInfo.InvariantCulture);
}
