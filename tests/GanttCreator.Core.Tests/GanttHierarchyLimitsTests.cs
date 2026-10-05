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
    /// A row whose parent is blocked gets no second, misleading depth finding.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The fixture is a parent refused for an <b>unrelated</b> fault — a start date
    /// after its finish, nothing to do with hierarchy — with a child beneath it.
    /// </para>
    /// <para>
    /// "One fault, one finding" still holds, and it holds because the depth rule now
    /// reads the parent's own <c>ParentId</c> rather than the parent's blocking flag:
    /// this parent is TOP-LEVEL, so its <c>ParentId</c> is null and the depth rule
    /// never fires. The child is refused once, by <c>PropagateBlockedParents</c>, with
    /// <c>ParentInvalid</c> — which is what is actually true.
    /// </para>
    /// <para>
    /// This test used to drive the same claim from a FOUR-level chain, on the
    /// reasoning that the depth rule blocks row 4 mid-loop and row 5 is therefore
    /// skipped. That only held for the top-down order: authored bottom-up, row 5 was
    /// reached before its parent was blocked and was refused as
    /// <c>HierarchyTooDeep</c> instead, so the same table produced two different
    /// findings for the same row purely from the input order. Depth is now decided
    /// from the parent's original <c>ParentId</c>, which makes it order-independent;
    /// the four-level chain is pinned in both orders by
    /// <c>GanttRowValidatorTests.A_four_level_chain_reports_identically_in_either_input_order</c>.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_row_whose_parent_is_blocked_gets_no_second_depth_finding()
    {
        var parentId = NewId();

        // Refused for a date fault alone; it is a top-level row, so nothing about its
        // own nesting is wrong.
        GanttRowDto parent = ValidSpan(
            rowNumber: 2,
            id: parentId,
            start: new DateOnly(2026, 6, 1),
            finish: new DateOnly(2026, 1, 1));
        GanttRowDto child = Child(3, NewId(), parentId);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.False(outcome.IsValid);

        // The child is refused exactly once, and for the reason that is true.
        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.HierarchyTooDeep);
        GanttValidationIssue childIssue = Assert.Single(
            outcome.Issues,
            i => i.RowNumber == 3);
        Assert.Equal(GanttValidationCodes.ParentInvalid, childIssue.Code);
        Assert.Contains(parentId, childIssue.Message, StringComparison.Ordinal);
    }
    [Fact]
    public void A_top_level_critical_interval_may_own_a_child()
    {
        var intervalId = NewId();
        GanttRowDto interval = CriticalIntervalRow(2, intervalId, parentId: null);

        GanttValidationOutcome outcome = GanttRowValidator.Validate(
            [interval, Child(3, NewId(), intervalId)]);

        // The whole point of the boundary: the pair is VALID, not merely free of a
        // depth finding. A top-level Critical Interval is a level-1 parent with a
        // level-2 child, which is exactly what EntityHierarchyCatalog lists
        // CriticalInterval among the child-owning types for. While the validator
        // still demanded a ParentId for that type, the interval itself was refused,
        // so this hierarchy could never exist and the catalogue entry was
        // unreachable -- the assertion below would have passed for the wrong reason.
        Assert.True(outcome.IsValid);
        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.HierarchyTooDeep);
        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.ParentMissingOrMalformed);
        Assert.Equal(2, outcome.Events.Count);
    }

    /// <summary>
    /// The mirror of the boundary above: a Critical Interval that IS a child cannot
    /// own a child of its own, because that would be depth 3. The depth rule is
    /// uniform in the type of the row being refused, so this is not a
    /// Critical-Interval special case -- it is the same rule
    /// <see cref="A_span_child_of_a_critical_interval_is_too_deep"/> pins for a
    /// non-critical child. Asserted here so the two ends of the owner's rule sit in
    /// one place and cannot drift apart.
    /// </summary>
    [Fact]
    public void A_nested_critical_interval_cannot_own_a_child()
    {
        var activityId = NewId();
        var intervalId = NewId();
        List<GanttRowDto> rows =
        [
            Parent(2, activityId),
            CriticalIntervalRow(3, intervalId, activityId),
            Child(4, NewId(), intervalId),
        ];

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

    /// <summary>
    /// The seven-child cap blocks the PARENT, and every one of that parent's children
    /// must drop with it. Before propagation was generalised and re-run after the
    /// capacity pass, the parent was refused but its eight children stayed in
    /// <c>Events</c> as valid rows pointing at a parent that was not there -- so the
    /// renderer received a hierarchy it could not lay out, and the only signal the
    /// user got was the single finding on the parent's row.
    /// </summary>
    [Fact]
    public void Every_child_of_a_refused_parent_is_absent_from_the_events()
    {
        var parentId = NewId();
        List<GanttRowDto> rows = FamilyWithChildren(parentId, 8);
        List<string> childIds = [.. rows.Skip(1).Select(row => row.IdCell.Value!)];

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.False(outcome.IsValid);
        // The parent is refused for the cap, and each child for the unusable parent.
        Assert.Single(outcome.Issues, i => i.Code == GanttValidationCodes.TooManyChildren);
        Assert.Equal(8, outcome.Issues.Count(i => i.Code == GanttValidationCodes.ParentInvalid));

        // Non-vacuity: the children really were valid rows before this pass ran --
        // they carry ids, types and dates, so nothing but propagation removed them.
        // The parent is dropped too (it carries the TooManyChildren finding), so the
        // whole nine-row family leaves Events together: a hierarchy the chart cannot
        // lay out is refused as a whole rather than half-built.
        Assert.Equal(8, childIds.Distinct(StringComparer.Ordinal).Count());
        Assert.DoesNotContain(outcome.Events, e => e.Id.Value == parentId);
        Assert.DoesNotContain(outcome.Events, e => childIds.Contains(e.Id.Value, StringComparer.Ordinal));
    }

    /// <summary>
    /// Order independence for a chain that crosses both passes. A parent over the cap
    /// is blocked by the capacity pass, which runs AFTER the direct parent pass has
    /// already decided its children -- so a child authored ABOVE its parent is the
    /// case that could only be covered by the second propagation run. The reversed
    /// table must produce the same findings, not merely a failing one.
    /// </summary>
    [Fact]
    public void Shuffling_an_over_capacity_activity_chain_preserves_the_validation_codes()
    {
        var parentId = NewId();
        List<GanttRowDto> forward = FamilyWithChildren(parentId, 8);
        List<GanttRowDto> reversed = [.. Enumerable.Reverse(forward)];

        GanttValidationOutcome forwardOutcome = GanttRowValidator.Validate(forward);
        GanttValidationOutcome reversedOutcome = GanttRowValidator.Validate(reversed);

        Assert.Equal(
            forwardOutcome.Issues.Select(i => (i.RowNumber, i.Code)).OrderBy(x => x.RowNumber).ThenBy(x => x.Code, StringComparer.Ordinal),
            reversedOutcome.Issues.Select(i => (i.RowNumber, i.Code)).OrderBy(x => x.RowNumber).ThenBy(x => x.Code, StringComparer.Ordinal));

        // And the same for which rows survive into the scene.
        Assert.Equal(
            forwardOutcome.Events.Select(e => e.Id.Value).Order(StringComparer.Ordinal),
            reversedOutcome.Events.Select(e => e.Id.Value).Order(StringComparer.Ordinal));
    }

    /// <summary>
    /// Propagation is not a Critical-Interval-only behaviour. A child activity whose
    /// parent is refused for an unrelated reason -- here a start date after its finish,
    /// nothing to do with hierarchy -- must be blocked too, because its parent is not
    /// a usable event. Keying propagation on <c>Type == CriticalInterval</c> left this
    /// child in <c>Events</c> pointing at a row that had been dropped.
    /// </summary>
    [Fact]
    public void A_child_activity_of_a_date_invalid_parent_is_blocked()
    {
        var parentId = NewId();

        // Valid in every respect except the dates: Start is after Finish.
        GanttRowDto parent = ValidSpan(rowNumber: 2, id: parentId, start: new DateOnly(2026, 6, 1), finish: new DateOnly(2026, 1, 1));
        GanttRowDto child = Child(3, NewId(), parentId);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.RowNumber == 2 && i.Code == GanttValidationCodes.StartAfterFinish);

        GanttValidationIssue childIssue = Assert.Single(
            outcome.Issues,
            i => i.RowNumber == 3 && i.Code == GanttValidationCodes.ParentInvalid);
        Assert.Equal("ParentId", childIssue.Field);
        Assert.DoesNotContain(outcome.Events, e => e.RowNumber == 3);
    }
}
