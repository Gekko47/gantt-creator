namespace GanttCreator.Core.Tests;

/// <summary>
/// Exact-text pins for <see cref="GanttRowValidator"/> issue messages.
/// </summary>
/// <remarks>
/// <para>
/// Mutation testing measured 59 of the validator's 117 open mutants as
/// message-string mutants: the interpolated text replaced by <c>""</c> or
/// <c>$""</c>. Every existing test asserted the issue <em>code</em>, so none of
/// them could notice a blanked explanation — and the R5.8 warnings panel
/// renders <see cref="GanttValidationIssue.Message"/> verbatim, which makes the
/// text a user-facing contract rather than a debugging convenience.
/// </para>
/// <para>
/// Each test here builds the smallest row that provokes one code and asserts
/// the whole message, including the interpolated offending value. Asserting a
/// prefix or a substring would leave the interpolation itself untested, which
/// is exactly the mutant class being killed.
/// </para>
/// </remarks>
public class GanttRowValidatorMessageTests
{
    private static int _nextId;

    private static string NewId() => $"G-{Interlocked.Increment(ref _nextId):x32}";

    /// <summary>A complete, otherwise-valid span row.</summary>
    private static GanttRowDto Span(
        int rowNumber = 2,
        string? id = null,
        string? laneId = null,
        int? stackIndex = 0,
        string? typeText = "As-Planned Activity",
        DateOnly? start = null,
        DateOnly? finish = null,
        string? parentId = null
    ) =>
        new(
            rowNumber,
            id ?? NewId(),
            laneId ?? NewId(),
            stackIndex,
            typeText,
            "Build frame",
            start,
            finish,
            parentId,
            null,
            null,
            null,
            null,
            true,
            null
        );

    /// <summary>The message carried by the single issue with <paramref name="code"/>.</summary>
    private static string MessageFor(GanttValidationOutcome outcome, string code) =>
        Assert.Single(outcome.Issues, issue => issue.Code == code).Message;

    [Fact]
    public void An_unknown_type_echoes_the_offending_text_and_names_the_catalogue()
    {
        GanttRowDto row = Span(typeText: "Widget", start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.Equal(
            "Unknown Type 'Widget'; use one of the 16 catalogue display names.",
            MessageFor(outcome, GanttValidationCodes.UnknownType));
    }

    [Fact]
    public void A_duplicate_id_names_the_offending_key_and_the_canonical_row()
    {
        string id = NewId();
        GanttRowDto first = Span(rowNumber: 2, id: id, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));
        GanttRowDto second = Span(rowNumber: 3, id: id, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));

        GanttValidationOutcome outcome = GanttRowValidator.Validate([first, second]);

        Assert.Equal(
            $"Duplicate Id '{id}'; the first row carrying it is canonical.",
            MessageFor(outcome, GanttValidationCodes.DuplicateId));
    }

    [Fact]
    public void A_span_missing_both_dates_reports_the_required_message_for_each()
    {
        GanttRowDto row = Span(start: null, finish: null);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.Equal("Start is required for this Type.", MessageFor(outcome, GanttValidationCodes.StartRequired));
        Assert.Equal("Finish is required for this Type.", MessageFor(outcome, GanttValidationCodes.FinishRequired));
    }

    [Fact]
    public void A_point_type_missing_its_single_date_says_so()
    {
        // A milestone reads Start only, so the message names the single-date
        // rule rather than the span rule.
        GanttRowDto row = new(
            2,
            NewId(),
            null,
            null,
            "As-Planned Milestone",
            "Handover",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.Equal(
            "Start is the single date for this Type and is required.",
            MessageFor(outcome, GanttValidationCodes.StartRequiredForPoint));
    }

    [Fact]
    public void A_malformed_parent_id_explains_the_expected_shape()
    {
        GanttRowDto row = Span(
            start: new DateOnly(2026, 9, 1),
            finish: new DateOnly(2026, 9, 5),
            parentId: "BAD");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.Equal(
            "ParentId must be blank or a well-formed row Id.",
            MessageFor(outcome, GanttValidationCodes.ParentMissingOrMalformed));
    }

    [Fact]
    public void An_unresolvable_parent_id_names_the_key_and_the_failure()
    {
        string parent = NewId();
        GanttRowDto row = Span(
            start: new DateOnly(2026, 9, 1),
            finish: new DateOnly(2026, 9, 5),
            parentId: parent);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.Equal(
            $"ParentId '{parent}' does not match any Id in the table.",
            MessageFor(outcome, GanttValidationCodes.ParentUnknown));
    }

    [Fact]
    public void A_finish_on_a_point_type_is_reported_as_unused_with_the_field_named()
    {
        GanttRowDto row = new(
            2,
            NewId(),
            null,
            null,
            "As-Planned Milestone",
            "Handover",
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 5),
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.Equal(
            "Finish is not used by this Type and is ignored for geometry.",
            MessageFor(outcome, GanttValidationCodes.NotUsedByType));
    }

    [Fact]
    public void Lane_and_stack_on_a_splitter_are_reported_as_unused_with_each_field_named()
    {
        // A Splitter is never lane-bound, so both supplied cells are ignored.
        // Each message must name its own field: a mutant that hard-codes one
        // field name would otherwise pass.
        GanttRowDto row = new(
            2,
            NewId(),
            NewId(),
            0,
            "Splitter",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        string[] messages = [.. outcome.Issues
            .Where(issue => issue.Code == GanttValidationCodes.NotUsedByType)
            .Select(issue => issue.Message)
            .OrderBy(text => text, StringComparer.Ordinal)];

        Assert.Equal(
            [
                "LaneId is not used by this Type and is ignored for geometry.",
                "StackIndex is not used by this Type and is ignored for geometry.",
            ],
            messages);
    }

    [Fact]
    public void A_self_reference_is_reported_as_a_cycle_not_as_a_depth_problem()
    {
        // The row names itself. The accurate finding is a cycle: telling the
        // user the row "is itself a child" describes nesting that does not
        // exist.
        string id = NewId();
        GanttRowDto row = Span(rowNumber: 2, id: id, parentId: id, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.Equal(
            "ParentId forms a parent cycle: a row names itself or an ancestor as its parent.",
            MessageFor(outcome, GanttValidationCodes.ParentCycle));
        Assert.DoesNotContain(outcome.Issues, issue => issue.Code == GanttValidationCodes.HierarchyTooDeep);
    }

    [Fact]
    public void A_grandchild_names_its_parent_and_the_supported_depth()
    {
        string top = NewId();
        string middle = NewId();

        GanttRowDto grandparent = Span(rowNumber: 2, id: top, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));
        GanttRowDto parent = Span(rowNumber: 3, id: middle, parentId: top, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));
        GanttRowDto child = Span(rowNumber: 4, parentId: middle, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));

        GanttValidationOutcome outcome = GanttRowValidator.Validate([grandparent, parent, child]);

        Assert.Equal(
            $"ParentId '{middle}' is itself a child, which would nest beyond the supported depth of {EntityHierarchyCatalog.MaxDepth}.",
            MessageFor(outcome, GanttValidationCodes.HierarchyTooDeep));
    }

    [Fact]
    public void Exceeding_the_child_cap_reports_the_count_and_the_limit()
    {
        int overLimit = EntityHierarchyCatalog.MaxChildrenPerParent + 1;
        string parentId = NewId();

        List<GanttRowDto> rows = [Span(rowNumber: 2, id: parentId, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5))];
        for (int i = 0; i < overLimit; i++)
        {
            rows.Add(Span(
                rowNumber: 3 + i,
                parentId: parentId,
                start: new DateOnly(2026, 9, 1),
                finish: new DateOnly(2026, 9, 5)));
        }

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.Equal(
            $"This row has {overLimit} children; at most {EntityHierarchyCatalog.MaxChildrenPerParent} are permitted.",
            MessageFor(outcome, GanttValidationCodes.TooManyChildren));
    }

    [Fact]
    public void An_ambiguous_parent_names_the_key_and_the_reason()
    {
        string duplicated = NewId();

        GanttRowDto first = Span(rowNumber: 2, id: duplicated, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));
        GanttRowDto second = Span(rowNumber: 3, id: duplicated, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));
        GanttRowDto child = Span(rowNumber: 4, parentId: duplicated, start: new DateOnly(2026, 9, 1), finish: new DateOnly(2026, 9, 5));

        GanttValidationOutcome outcome = GanttRowValidator.Validate([first, second, child]);

        Assert.Equal(
            $"ParentId '{duplicated}' is ambiguous because multiple rows carry that Id.",
            MessageFor(outcome, GanttValidationCodes.ParentAmbiguous));
    }

    [Fact]
    public void A_parent_that_may_not_own_children_is_reported_with_the_rule()
    {
        string milestoneId = NewId();
        GanttRowDto milestone = new(
            2,
            milestoneId,
            null,
            null,
            "As-Planned Milestone",
            "Handover",
            new DateOnly(2026, 9, 1),
            null,
            null,
            null,
            null,
            null,
            null,
            true,
            null);
        GanttRowDto child = Span(
            rowNumber: 3,
            parentId: milestoneId,
            start: new DateOnly(2026, 9, 1),
            finish: new DateOnly(2026, 9, 5));

        GanttValidationOutcome outcome = GanttRowValidator.Validate([milestone, child]);

        Assert.Equal(
            $"ParentId '{milestoneId}' must reference a type that may own children.",
            MessageFor(outcome, GanttValidationCodes.ParentNotSpan));
    }

    [Fact]
    public void A_parent_that_did_not_validate_is_reported_as_unusable()
    {
        string brokenParentId = NewId();
        GanttRowDto brokenParent = Span(
            rowNumber: 2,
            id: brokenParentId,
            start: new DateOnly(2026, 9, 5),
            finish: new DateOnly(2026, 9, 1));
        GanttRowDto child = Span(
            rowNumber: 3,
            parentId: brokenParentId,
            start: new DateOnly(2026, 9, 1),
            finish: new DateOnly(2026, 9, 5));

        GanttValidationOutcome outcome = GanttRowValidator.Validate([brokenParent, child]);

        Assert.Equal(
            $"ParentId '{brokenParentId}' references a row that did not validate as a usable event.",
            MessageFor(outcome, GanttValidationCodes.ParentInvalid));
    }
}
