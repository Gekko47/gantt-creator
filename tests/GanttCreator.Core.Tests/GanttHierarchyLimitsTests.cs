namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for the two hierarchy limits R4.7A fixes: the seven-child maximum
/// (D4) and the two-level nesting depth (D5), enforced in
/// <c>GanttRowValidator.CheckChildCapacityAndDepth</c>.
/// </summary>
public sealed class GanttHierarchyLimitsTests
{
    private static string NewId() => GanttRowId.New().Value;

    private static string NewLaneId() => GanttRowId.New().Value;

    /// <summary>
    /// A minimal valid span row. These builders are local to this file rather than
    /// shared with <c>GanttRowValidatorTests</c>: that file's helpers are private
    /// and carry its own defaults, and promoting them to a shared helper would be
    /// a refactor of passing tests that this row has no reason to touch.
    /// </summary>
    private static GanttRowDto ValidSpan(
        int rowNumber = 2,
        string? id = null,
        string? parentId = null,
        DateOnly? start = null,
        DateOnly? finish = null) =>
        new(
            rowNumber,
            id ?? NewId(),
            NewLaneId(),
            0,
            "As-Planned Activity",
            "Build frame",
            start ?? new DateOnly(2026, 1, 1),
            finish ?? new DateOnly(2026, 12, 31),
            parentId,
            null,
            null,
            null,
            null,
            true,
            null);

    private static GanttRowDto CriticalIntervalRow(int rowNumber, string id, string? parentId) =>
        new(
            rowNumber,
            id,
            NewLaneId(),
            0,
            "Critical Interval",
            "Critical part",
            new DateOnly(2026, 3, 1),
            new DateOnly(2026, 3, 15),
            parentId,
            null,
            null,
            null,
            null,
            true,
            null);

    /// <summary>
    /// A parent with a span child, the simplest legal hierarchy.
    /// </summary>
    private static GanttRowDto Parent(int rowNumber, string id) =>
        ValidSpan(rowNumber: rowNumber, id: id, start: new DateOnly(2026, 1, 1), finish: new DateOnly(2026, 12, 31));

    private static GanttRowDto Child(int rowNumber, string id, string parentId) =>
        ValidSpan(rowNumber: rowNumber, id: id, parentId: parentId, start: new DateOnly(2026, 3, 1), finish: new DateOnly(2026, 4, 1));

    private static List<GanttRowDto> FamilyWithChildren(string parentId, int childCount)
    {
        List<GanttRowDto> rows = [Parent(2, parentId)];
        for (var i = 0; i < childCount; i++)
        {
            rows.Add(Child(3 + i, NewId(), parentId));
        }

        return rows;
    }

    /// <summary>
    /// Seven children is the documented maximum, so exactly seven must validate.
    /// This is the boundary case that the eighth-child test below depends on: a
    /// cap that rejected seven would make the positive test meaningless.
    /// </summary>
    [Fact]
    public void Seven_children_are_permitted()
    {
        var parentId = NewId();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(FamilyWithChildren(parentId, 7));

        Assert.True(outcome.IsValid);
        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.TooManyChildren);
        Assert.Equal(8, outcome.Events.Count);
    }

    /// <summary>
    /// The eighth child is refused, and the finding is reported against the
    /// PARENT's row -- the row the user must edit -- carrying the actual count.
    /// </summary>
    [Fact]
    public void An_eighth_child_is_refused_on_the_parent_row()
    {
        var parentId = NewId();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(FamilyWithChildren(parentId, 8));

        Assert.False(outcome.IsValid);
        GanttValidationIssue issue = Assert.Single(
            outcome.Issues,
            i => i.Code == GanttValidationCodes.TooManyChildren);
        Assert.Equal(2, issue.RowNumber);
        Assert.Equal("ParentId", issue.Field);
        Assert.Equal(GanttValidationSeverity.Error, issue.Severity);
        Assert.Contains("8", issue.Message, StringComparison.Ordinal);
        Assert.Contains("7", issue.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// The cap is per parent, not a table-wide total. Two parents with four
    /// children each is 8 children in total and must still validate -- this is
    /// what proves the check groups by parent rather than counting globally.
    /// </summary>
    [Fact]
    public void The_cap_is_per_parent_not_per_table()
    {
        var firstParent = NewId();
        var secondParent = NewId();
        List<GanttRowDto> rows = [Parent(2, firstParent), Parent(3, secondParent)];
        for (var i = 0; i < 4; i++)
        {
            rows.Add(Child(4 + i, NewId(), firstParent));
            rows.Add(Child(8 + i, NewId(), secondParent));
        }

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.True(outcome.IsValid);
        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.TooManyChildren);
    }

    /// <summary>
    /// Determinism: the same table produces the same findings regardless of the
    /// order the rows are supplied in, so a re-sorted worksheet does not change
    /// which row is reported.
    /// </summary>
    [Fact]
    public void The_cap_reports_the_same_parent_in_either_input_order()
    {
        var parentId = NewId();
        List<GanttRowDto> family = FamilyWithChildren(parentId, 8);
        List<GanttRowDto> reversed = [.. Enumerable.Reverse(family)];

        GanttValidationOutcome forwardOutcome = GanttRowValidator.Validate(family);
        GanttValidationOutcome reversedOutcome = GanttRowValidator.Validate(reversed);

        GanttValidationIssue forward = Assert.Single(forwardOutcome.Issues, i => i.Code == GanttValidationCodes.TooManyChildren);
        GanttValidationIssue backward = Assert.Single(reversedOutcome.Issues, i => i.Code == GanttValidationCodes.TooManyChildren);

        Assert.Equal(forward.RowNumber, backward.RowNumber);
        Assert.Equal(forward.Message, backward.Message);
    }

    /// <summary>
    /// A grandchild is refused: a child's own parent is itself a child, which
    /// would nest to depth 3. Reported on the grandchild's row, because that is
    /// the row whose ParentId is wrong.
    /// </summary>
    [Fact]
    public void A_grandchild_is_refused()
    {
        var parentId = NewId();
        var childId = NewId();
        List<GanttRowDto> rows = [Parent(2, parentId), Child(3, childId, parentId), Child(4, NewId(), childId)];

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.False(outcome.IsValid);
        GanttValidationIssue issue = Assert.Single(outcome.Issues, i => i.Code == GanttValidationCodes.HierarchyTooDeep);
        Assert.Equal(4, issue.RowNumber);
    }

    /// <summary>
    /// A Critical Interval parented by an activity, with a Critical Interval child
    /// of its own, is depth 3 and must be refused. The child used to be exempt from
    /// the depth check on the grounds that the shape was pre-existing landed
    /// behaviour; the owner has ruled that a Critical Interval is a level-2 child
    /// and cannot itself own a child, so the exemption is withdrawn. This is the
    /// test that pins the withdrawn exemption, so a later re-introduction is a
    /// deliberate contract change rather than a silent regression.
    /// </summary>
    [Fact]
    public void A_critical_interval_child_of_a_critical_interval_is_too_deep()
    {
        var activityId = NewId();
        var outerIntervalId = NewId();
        GanttRowDto activity = Parent(2, activityId);
        GanttRowDto outerInterval = CriticalIntervalRow(3, outerIntervalId, activityId);
        GanttRowDto innerInterval = CriticalIntervalRow(4, NewId(), outerIntervalId);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([activity, outerInterval, innerInterval]);

        Assert.False(outcome.IsValid);
        GanttValidationIssue issue = Assert.Single(outcome.Issues, i => i.Code == GanttValidationCodes.HierarchyTooDeep);
        Assert.Equal(4, issue.RowNumber);
    }

    /// <summary>
    /// A span child of a Critical Interval parent is the same depth-3 case, and is
    /// refused on the same rule. Together with the test above this shows the depth
    /// check no longer branches on the child's type at all: the two rows differ
    /// only in the type of the row being refused, and both are refused.
    /// </summary>
    [Fact]
    public void A_span_child_of_a_critical_interval_is_too_deep()
    {
        var activityId = NewId();
        var outerIntervalId = NewId();
        List<GanttRowDto> rows = [Parent(2, activityId), CriticalIntervalRow(3, outerIntervalId, activityId), Child(4, NewId(), outerIntervalId)];

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.False(outcome.IsValid);
        GanttValidationIssue issue = Assert.Single(outcome.Issues, i => i.Code == GanttValidationCodes.HierarchyTooDeep);
        Assert.Equal(4, issue.RowNumber);
    }

    /// <summary>
    /// The boundary the owner's rule leaves intact: a Critical Interval that is
    /// itself TOP-LEVEL may own a child, because that is a level-1 parent with a
    /// level-2 child. This is the case the depth rule must not over-reach into,
    /// and it is why <c>EntityHierarchyCatalog</c> still lists CriticalInterval
    /// among the types that may own children.
    /// </summary>
    [Fact]
    public void A_top_level_critical_interval_may_own_a_child()
    {
        var intervalId = NewId();
        GanttRowDto interval = CriticalIntervalRow(2, intervalId, parentId: null);

        GanttValidationOutcome outcome = GanttRowValidator.Validate(
            [interval, Child(3, NewId(), intervalId)]);

        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.HierarchyTooDeep);
    }

    /// <summary>
    /// A child whose parent is already known to be broken must not also raise a
    /// capacity error. The child of an unknown parent is already blocking, and a
    /// second unrelated error on the same problem is noise that hides the real one.
    /// </summary>
    [Fact]
    public void A_child_of_a_broken_parent_does_not_add_a_capacity_error()
    {
        List<GanttRowDto> rows = [];
        for (var i = 0; i < 9; i++)
        {
            rows.Add(Child(2 + i, NewId(), NewId()));
        }

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.ParentUnknown);
        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.TooManyChildren);
    }
}
