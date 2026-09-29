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

    private static GanttRowDto CriticalIntervalRow(int rowNumber, string id, string parentId) =>
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
    /// of its own, is depth-legal: the critical interval's parent is top-level, so
    /// the child sits at depth 2. This is the pre-existing landed shape that the
    /// matrix correction preserved, and the depth check must not break it.
    /// </summary>
    [Fact]
    public void A_critical_interval_child_of_a_critical_interval_is_depth_legal()
    {
        var activityId = NewId();
        var outerIntervalId = NewId();
        GanttRowDto activity = Parent(2, activityId);
        GanttRowDto outerInterval = CriticalIntervalRow(3, outerIntervalId, activityId);
        GanttRowDto innerInterval = CriticalIntervalRow(4, NewId(), outerIntervalId);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([activity, outerInterval, innerInterval]);

        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.HierarchyTooDeep);
    }

    /// <summary>
    /// The counterpart to the test above, and the reason the depth rule is stated
    /// in terms of the CHILD's type: a plain span child of a Critical Interval
    /// parent IS depth 3 and must be refused, even though the critical-interval
    /// parent is itself a child. The first draft of this check tested the
    /// parent's state instead, which wrongly allowed this case and wrongly
    /// refused the legal one above -- the two tests together are what make the
    /// rule unambiguous.
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
