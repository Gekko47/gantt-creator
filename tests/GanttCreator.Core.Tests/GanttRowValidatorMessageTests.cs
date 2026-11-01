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
}
