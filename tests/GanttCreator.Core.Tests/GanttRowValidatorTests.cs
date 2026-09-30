using System.Globalization;

namespace GanttCreator.Core.Tests;

public class GanttRowValidatorTests
{
    private static int s_nextId;

    private static string NewId() => $"G-{Interlocked.Increment(ref s_nextId):x32}";

    private static string NewLaneId() => NewId();

    private const string AutoId = "\0auto";

    private static GanttRowDto ValidSpan(
        int rowNumber = 2,
        string? typeText = "As-Planned Activity",
        string? id = AutoId,
        string? laneId = AutoId,
        int? stackIndex = 0,
        DateOnly? start = null,
        DateOnly? finish = null,
        string? description = "Build frame",
        string? parentId = null,
        string? styleKey = null,
        string? labelPositionText = null,
        string? fillColourText = null,
        string? strokeColourText = null,
        bool? visible = true,
        string? sortOrder = null
    ) =>
        new(
            rowNumber,
            id == AutoId ? NewId() : id,
            laneId == AutoId ? NewLaneId() : laneId,
            stackIndex,
            typeText,
            description,
            start ?? new DateOnly(2026, 9, 1),
            finish ?? new DateOnly(2026, 9, 5),
            parentId,
            styleKey,
            labelPositionText,
            fillColourText,
            strokeColourText,
            visible,
            sortOrder
        );

    [Fact]
    public void Validate_rejects_null_input_without_swallowing_programmer_error() =>
        _ = Assert.Throws<ArgumentNullException>(() => GanttRowValidator.Validate(null!));

    [Fact]
    public void Valid_span_row_maps_to_event_with_defaults()
    {
        var id = NewId();
        var lane = NewLaneId();
        GanttRowDto row = ValidSpan(id: id, laneId: lane);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.Empty(outcome.Issues);
        GanttEvent @event = Assert.Single(outcome.Events);
        Assert.Equal(2, @event.RowNumber);
        Assert.Equal(id, @event.Id.Value);
        Assert.Equal(lane, @event.LaneId!.Value);
        Assert.Equal(0, @event.StackIndex);
        Assert.Equal(GanttEntityType.AsPlannedActivity, @event.Type);
        Assert.Equal("Build frame", @event.Description);
        Assert.Equal(new DateOnly(2026, 9, 1), @event.Start);
        Assert.Equal(new DateOnly(2026, 9, 5), @event.Finish);
        Assert.True(@event.Visible);
    }

    [Theory]
    [InlineData("Splitter")]
    [InlineData("Spacer")]
    [InlineData("As-Built Activity")]
    [InlineData("As-Planned Activity")]
    [InlineData("Baseline Activity")]
    [InlineData("Critical Interval")]
    [InlineData("Delay Event")]
    [InlineData("As-Built Procurement")]
    [InlineData("As-Planned Procurement")]
    [InlineData("Baseline Procurement")]
    [InlineData("Custom Activity")]
    [InlineData("As-Built Milestone")]
    [InlineData("As-Planned Milestone")]
    [InlineData("Baseline Milestone")]
    [InlineData("Critical Milestone")]
    [InlineData("Delineator")]
    public void All_16_catalogue_types_validate_when_required_fields_present(string typeText)
    {
        GanttEntityType type = EntityTypeCatalog.GetDefinition(typeText)!.Type;
        EntityTypeDefinition definition = EntityTypeCatalog.GetDefinition(type)!;
        var id = NewId();
        var lane = definition.Kind == EntityKind.Span && type != GanttEntityType.CriticalInterval ? NewLaneId() : null;
        var stack = definition.Kind == EntityKind.Span && type != GanttEntityType.CriticalInterval ? (int?)0 : null;

        DateOnly? start = definition.DateMode == EntityDateMode.None ? null : new DateOnly(2026, 9, 1);
        DateOnly? finish = definition.DateMode == EntityDateMode.StartFinish ? new DateOnly(2026, 9, 5) : null;
        string? parentId = null;
        var styleKey = type == GanttEntityType.CustomActivity ? "MyStyle" : null;

        if (type == GanttEntityType.CriticalInterval)
        {
            var parentIdText = NewId();
            GanttRowDto parent = ValidSpan(rowNumber: 1, typeText: "As-Planned Activity", id: parentIdText);
            var child = new GanttRowDto(
                2,
                NewId(),
                null,
                0,
                typeText,
                "Critical part",
                new DateOnly(2026, 9, 2),
                new DateOnly(2026, 9, 3),
                parentIdText,
                null,
                null,
                null,
                null,
                true,
                null
            );
            GanttValidationOutcome batch = GanttRowValidator.Validate([parent, child]);
            Assert.True(batch.IsValid);
            Assert.Equal(2, batch.Events.Count);
            return;
        }

        var row = new GanttRowDto(2, id, lane, stack, typeText, null, start, finish, parentId, styleKey, null, null, null, null, null);
        // Milestones carry no lane/stack here (optional); spans carry both.
        if (definition.Kind == EntityKind.Milestone)
        {
            row = row with { LaneIdCell = GanttCells.Empty<string>(), StackIndexCell = GanttCells.Empty<int?>() };
        }

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.Empty(outcome.Issues);
        GanttEvent @event = Assert.Single(outcome.Events);
        Assert.Equal(type, @event.Type);
        Assert.True(@event.Visible);
        Assert.Null(@event.Description);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("G-SHORT")]
    [InlineData("X-0123456789abcdef0123456789abcdef")]
    [InlineData("G-0123456789ABCDEF0123456789ABCDEF")]
    public void Malformed_id_is_a_blocking_error(string? id)
    {
        GanttRowDto row = ValidSpan(id: id);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Empty(outcome.Events);
        Assert.Contains(
            outcome.Issues,
            i => i.Field == "Id" && i.Code == GanttValidationCodes.IdMissingOrMalformed && i.Severity == GanttValidationSeverity.Error
        );
    }

    [Fact]
    public void Duplicate_id_keeps_first_canonical_and_errors_later_rows()
    {
        var shared = NewId();
        GanttRowDto first = ValidSpan(rowNumber: 2, id: shared);
        GanttRowDto second = ValidSpan(rowNumber: 3, id: shared);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([first, second]);

        Assert.False(outcome.IsValid);
        GanttEvent single = Assert.Single(outcome.Events);
        Assert.Equal(2, single.RowNumber);
        GanttValidationIssue issue = Assert.Single(outcome.Issues, i => i.Code == GanttValidationCodes.DuplicateId);
        Assert.Equal(3, issue.RowNumber);
        Assert.Equal(GanttValidationSeverity.Error, issue.Severity);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Widget")]
    [InlineData("as-planned activity")]
    public void Unknown_type_is_a_blocking_error(string? typeText)
    {
        GanttRowDto row = ValidSpan(typeText: typeText);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "Type" && i.Code == GanttValidationCodes.UnknownType);
    }

    [Fact]
    public void Span_missing_start_is_a_blocking_error()
    {
        GanttRowDto row = ValidSpan() with { StartCell = GanttCells.Empty<DateOnly?>() };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.StartRequired);
    }

    [Fact]
    public void Span_missing_finish_is_a_blocking_error()
    {
        GanttRowDto row = ValidSpan() with { FinishCell = GanttCells.Empty<DateOnly?>() };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.FinishRequired);
    }

    [Fact]
    public void Span_start_after_finish_is_a_blocking_error()
    {
        GanttRowDto row = ValidSpan(start: new DateOnly(2026, 9, 5), finish: new DateOnly(2026, 9, 1));

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.StartAfterFinish);
    }

    [Fact]
    public void One_day_span_with_equal_dates_is_valid()
    {
        var day = new DateOnly(2024, 2, 29);
        GanttRowDto row = ValidSpan(start: day, finish: day);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
    }

    [Fact]
    public void Milestone_missing_start_is_a_blocking_error()
    {
        var row = new GanttRowDto(
            2,
            NewId(),
            null,
            null,
            "As-Built Milestone",
            "Pour",
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

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.StartRequiredForPoint);
    }

    [Fact]
    public void Milestone_finish_is_warned_and_ignored_for_geometry()
    {
        var row = new GanttRowDto(
            2,
            NewId(),
            null,
            null,
            "As-Built Milestone",
            "Pour",
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 9),
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.Contains(
            outcome.Issues,
            i => i.Field == "Finish" && i.Code == GanttValidationCodes.NotUsedByType && i.Severity == GanttValidationSeverity.Warning
        );
        Assert.Null(Assert.Single(outcome.Events).Finish);
    }

    [Fact]
    public void Delineator_lane_and_stack_are_warned_and_cleared()
    {
        var row = new GanttRowDto(
            2,
            NewId(),
            NewLaneId(),
            1,
            "Delineator",
            null,
            new DateOnly(2026, 9, 1),
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

        Assert.True(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "LaneId" && i.Code == GanttValidationCodes.NotUsedByType);
        Assert.Contains(outcome.Issues, i => i.Field == "StackIndex" && i.Code == GanttValidationCodes.NotUsedByType);
        GanttEvent @event = Assert.Single(outcome.Events);
        Assert.Null(@event.LaneId);
        Assert.Null(@event.StackIndex);
    }

    [Theory]
    [InlineData("Splitter")]
    [InlineData("Spacer")]
    public void Splitter_and_spacer_dates_are_warned_and_cleared(string typeText)
    {
        var row = new GanttRowDto(
            2,
            NewId(),
            NewLaneId(),
            0,
            typeText,
            "Section",
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 2),
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "Start" && i.Code == GanttValidationCodes.NotUsedByType);
        Assert.Contains(outcome.Issues, i => i.Field == "Finish" && i.Code == GanttValidationCodes.NotUsedByType);
    }

    [Fact]
    public void Span_missing_lane_is_a_blocking_error()
    {
        GanttRowDto row = ValidSpan() with { LaneIdCell = GanttCells.Empty<string>() };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.LaneIdMissingOrMalformed);
    }

    [Fact]
    public void Malformed_lane_is_a_blocking_error()
    {
        GanttRowDto row = ValidSpan(laneId: "NOT-AN-ID");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "LaneId" && i.Code == GanttValidationCodes.LaneIdMissingOrMalformed);
    }

    [Fact]
    public void Span_missing_stack_is_a_blocking_error()
    {
        GanttRowDto row = ValidSpan(stackIndex: null);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.StackIndexRequired);
    }

    [Fact]
    public void Critical_interval_without_parent_is_a_blocking_error()
    {
        var row = new GanttRowDto(
            2,
            NewId(),
            NewLaneId(),
            0,
            "Critical Interval",
            "Critical part",
            new DateOnly(2026, 9, 2),
            new DateOnly(2026, 9, 3),
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.ParentMissingOrMalformed);
    }

    [Fact]
    public void Critical_interval_with_duplicate_parent_is_reported_as_ambiguous()
    {
        var parentId = NewId();
        GanttRowDto firstParent = ValidSpan(rowNumber: 2, id: parentId);
        GanttRowDto secondParent = ValidSpan(rowNumber: 3, id: parentId);
        GanttRowDto child = ValidSpan(
            rowNumber: 4,
            typeText: "Critical Interval",
            id: NewId(),
            parentId: parentId);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([firstParent, secondParent, child]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, issue =>
            issue.RowNumber == 4
            && issue.Code == GanttValidationCodes.ParentAmbiguous);
    }

    [Fact]
    public void Critical_interval_with_duplicate_parent_is_reported_as_ambiguous_not_as_a_cycle()
    {
        // The parent Id is duplicated, so the reference cannot be resolved to one
        // row. Claiming a cycle through that edge would replace the accurate
        // ambiguity finding with a diagnosis the data does not support.
        string duplicatedId = NewId();
        string childId = NewId();
        GanttValidationOutcome outcome = GanttRowValidator.Validate([
            CriticalIntervalRow(2, duplicatedId, childId),
            ValidSpan(rowNumber: 3, id: duplicatedId),
            CriticalIntervalRow(4, childId, duplicatedId),
        ]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, issue =>
            issue.RowNumber == 4
            && issue.Code == GanttValidationCodes.ParentAmbiguous);
        Assert.DoesNotContain(outcome.Issues, issue => issue.Code == GanttValidationCodes.ParentCycle);
    }

    [Fact]
    public void Critical_interval_with_unknown_parent_is_a_blocking_error()
    {
        var row = new GanttRowDto(
            2,
            NewId(),
            NewLaneId(),
            0,
            "Critical Interval",
            "Critical part",
            new DateOnly(2026, 9, 2),
            new DateOnly(2026, 9, 3),
            NewId(),
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.ParentUnknown);
    }

    [Fact]
    public void Critical_interval_with_non_span_parent_is_a_blocking_error()
    {
        var milestoneId = NewId();
        var milestone = new GanttRowDto(
            2,
            milestoneId,
            null,
            null,
            "As-Built Milestone",
            "Pour",
            new DateOnly(2026, 9, 1),
            null,
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );
        var child = new GanttRowDto(
            3,
            NewId(),
            NewLaneId(),
            0,
            "Critical Interval",
            "Part",
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 1),
            milestoneId,
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([milestone, child]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.RowNumber == 3 && i.Code == GanttValidationCodes.ParentNotSpan);
    }

    [Fact]
    public void Parent_on_a_child_capable_span_is_resolved_not_ignored()
    {
        // R4.7A D9 replaced this test's contract. A `ParentId` on an
        // As-Planned Activity used to be a NotUsedByType warning, because only a
        // Critical Interval could carry one. ParentId is now authoritative for
        // general hierarchy, so this row is a real child whose parent does not
        // exist -- which is a blocking error, not a warning. The superseded
        // NotUsedByType case is still covered for structural types by
        // `Parent_on_a_structural_type_is_a_not_used_warning`.
        GanttRowDto row = ValidSpan(parentId: NewId());

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "ParentId" && i.Severity == GanttValidationSeverity.Error);
        Assert.DoesNotContain(outcome.Issues, i => i.Code == GanttValidationCodes.NotUsedByType);
    }

    [Fact]
    public void Child_capable_span_with_a_resolvable_parent_validates_as_a_child()
    {
        // The positive case for D9: an As-Planned Activity carrying a ParentId
        // that resolves to a real parent must be VALID and produce an event whose
        // ParentId is preserved. Without this, the widening would only be proven
        // by the absence of a warning, which a later change could reintroduce.
        string parentId = NewId();
        GanttRowDto parent = ValidSpan(rowNumber: 2, id: parentId, start: new DateOnly(2026, 1, 1), finish: new DateOnly(2026, 12, 31));
        GanttRowDto child = ValidSpan(rowNumber: 3, id: NewId(), parentId: parentId);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.True(outcome.IsValid, "A resolved child must validate.");
        Assert.DoesNotContain(outcome.Issues, i => i.Field == "ParentId");

        GanttRowId childId = GanttRowId.Parse(child.Id);
        GanttEvent childEvent = Assert.Single(outcome.Events, e => e.Id == childId);
        Assert.Equal(parentId, childEvent.ParentId!.Value);
    }

    [Fact]
    public void Child_capable_span_with_a_blank_parent_is_top_level_and_valid()
    {
        // Blank ParentId on a child-capable type is the promoted state: after a
        // parent is deleted its children become top-level, so this must be valid
        // rather than reported as a missing required field.
        GanttRowDto row = ValidSpan(parentId: null);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.DoesNotContain(outcome.Issues, i => i.Field == "ParentId");
    }

    [Fact]
    public void Child_capable_span_with_a_malformed_parent_is_blocked()
    {
        GanttRowDto row = ValidSpan(parentId: "not-an-id");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(
            outcome.Issues,
            i => i.Field == "ParentId" && i.Code == GanttValidationCodes.ParentMissingOrMalformed && i.Severity == GanttValidationSeverity.Error);
    }

    [Fact]
    public void Child_of_a_type_that_may_not_own_children_is_blocked()
    {
        // A milestone is child-capable but may not own children, so a child
        // naming a milestone parent is refused with ParentNotSpan. This is the
        // widened replacement for the old Span-only test, and it proves the
        // matrix -- not a hard-coded Span check -- is the authority.
        string parentId = NewId();
        GanttRowDto parent = ValidSpan(
            rowNumber: 2,
            id: parentId,
            typeText: "As-Planned Milestone",
            stackIndex: null,
            finish: new DateOnly(2026, 9, 5));
        GanttRowDto child = ValidSpan(rowNumber: 3, id: NewId(), parentId: parentId);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.RowNumber == 3 && i.Code == GanttValidationCodes.ParentNotSpan);
    }

    [Fact]
    public void Parent_on_a_structural_type_is_a_not_used_warning()
    {
        // The D9 guarantee that nothing becomes silently permitted: a type the
        // matrix does not classify as child-capable still reports NotUsedByType
        // rather than being accepted as a child or blocked as an unresolvable one.
        foreach (string typeText in new[] { "Splitter", "Spacer", "Delineator" })
        {
            GanttRowDto row = new(
                2,
                NewId(),
                null,
                null,
                typeText,
                "Section",
                new DateOnly(2026, 9, 1),
                new DateOnly(2026, 9, 5),
                NewId(),
                null,
                null,
                null,
                null,
                true,
                null
            );

            GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

            Assert.Contains(
                outcome.Issues,
                i => i.Field == "ParentId" && i.Code == GanttValidationCodes.NotUsedByType && i.Severity == GanttValidationSeverity.Warning);
        }
    }

    [Fact]
    public void Custom_activity_without_style_key_is_a_blocking_error()
    {
        GanttRowDto row = ValidSpan(typeText: "Custom Activity", styleKey: null);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.StyleKeyRequired);
    }

    [Fact]
    public void Custom_activity_with_any_style_key_passes_without_capability_checks()
    {
        GanttRowDto row = ValidSpan(
            typeText: "Custom Activity",
            styleKey: "MyStyle",
            labelPositionText: "Inside",
            fillColourText: "#ff0000",
            strokeColourText: "#00ff00"
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.Empty(outcome.Issues);
    }

    [Fact]
    public void Custom_activity_with_registry_rejects_unknown_style_key()
    {
        GanttRowDto row = ValidSpan(typeText: "Custom Activity", styleKey: "MissingStyle");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row], GanttStyleRegistry.Empty);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "StyleKey" && i.Code == GanttValidationCodes.StyleKeyUnknown);
    }

    [Fact]
    public void Custom_activity_uses_style_label_and_colour_capabilities()
    {
        var style = new GanttStyleDefinition(
            "MyStyle",
            new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
            EntityColourCapability.Fill);
        var registry = new GanttStyleRegistry([style]);
        GanttRowDto badLabel = ValidSpan(
            typeText: "Custom Activity",
            styleKey: "MyStyle",
            // A position that parses but is outside the style's set. It used to be
            // "Above", which the owner ruling retired: an unparseable name can
            // only ever produce UnknownLabelPosition, so it no longer exercises
            // the capability check this test is about.
            labelPositionText: nameof(GanttLabelPosition.TopLeft));
        GanttRowDto badColour = ValidSpan(
            typeText: "Custom Activity",
            styleKey: "MyStyle",
            strokeColourText: "#FF0000");

        GanttValidationOutcome labelOutcome = GanttRowValidator.Validate([badLabel], registry);
        GanttValidationOutcome colourOutcome = GanttRowValidator.Validate([badColour], registry);

        Assert.Contains(labelOutcome.Issues, i => i.Code == GanttValidationCodes.LabelNotAllowedForType);
        Assert.Contains(colourOutcome.Issues, i => i.Code == GanttValidationCodes.ColourNotAllowedForType);
    }

    [Fact]
    public void A_retired_label_position_is_reported_and_never_coerced()
    {
        // ADR-0029 D6 (owner ruling 2026-09-30): a stored "Above"/"Below" is
        // REPORTED, never coerced to a neighbouring position. This is the validator
        // half of the removal, and it is the reason the enum members were deleted
        // rather than merely denied: there is nothing left to silently substitute.
        foreach (string stored in new[] { "Above", "Below" })
        {
            GanttRowDto row = ValidSpan(
                typeText: "As-Planned Activity",
                labelPositionText: stored);

            GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

            Assert.False(outcome.IsValid);
            GanttValidationIssue issue = Assert.Single(
                outcome.Issues,
                i => i.Code == GanttValidationCodes.UnknownLabelPosition);
            Assert.Equal("LabelPosition", issue.Field);
            // Reported verbatim: the message names the stored value so the analyst
            // can see which cell needs changing.
            Assert.Contains(stored, issue.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Unknown_label_position_is_a_blocking_error()
    {
        GanttRowDto row = ValidSpan(labelPositionText: "Sideways");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.UnknownLabelPosition);
    }


    [Fact]
    public void Label_inside_is_not_allowed_for_milestones()
    {
        var row = new GanttRowDto(
            2,
            NewId(),
            null,
            null,
            "As-Built Milestone",
            "Pour",
            new DateOnly(2026, 9, 1),
            null,
            null,
            null,
            "Inside",
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.LabelNotAllowedForType);
    }

    [Fact]
    public void Critical_interval_with_any_label_is_a_blocking_error()
    {
        var parentId = NewId();
        GanttRowDto parent = ValidSpan(rowNumber: 2, id: parentId);
        var child = new GanttRowDto(
            3,
            NewId(),
            null,
            null,
            "Critical Interval",
            "Part",
            new DateOnly(2026, 9, 2),
            new DateOnly(2026, 9, 3),
            parentId,
            null,
            "Auto",
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.RowNumber == 3 && i.Code == GanttValidationCodes.LabelNotAllowedForType);
    }

    [Theory]
    [InlineData("#GG0000")]
    [InlineData("red")]
    [InlineData("#FFF")]
    [InlineData("#1234567")]
    public void Bad_colour_format_is_a_blocking_error(string colour)
    {
        GanttRowDto row = ValidSpan(fillColourText: colour);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "FillColour" && i.Code == GanttValidationCodes.BadColourFormat);
    }

    [Fact]
    public void Lowercase_colour_is_accepted_and_normalised_to_uppercase()
    {
        GanttRowDto row = ValidSpan(fillColourText: "#ff0000", strokeColourText: "#00ff00");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        GanttEvent @event = Assert.Single(outcome.Events);
        Assert.Equal("#FF0000", @event.FillColour);
        Assert.Equal("#00FF00", @event.StrokeColour);
    }

    [Fact]
    public void Fill_on_critical_interval_is_accepted()
    {
        // ADR-0027 D5 inverted this rule. It is kept as its own test rather than
        // deleted so the inversion stays visible: the identical row that was a
        // blocking error now validates, and the fill reaches the event.
        var parentId = NewId();
        GanttRowDto parent = ValidSpan(rowNumber: 2, id: parentId);
        var child = new GanttRowDto(
            3,
            NewId(),
            null,
            null,
            "Critical Interval",
            "Part",
            new DateOnly(2026, 9, 2),
            new DateOnly(2026, 9, 3),
            parentId,
            null,
            null,
            "#FF0000",
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.True(outcome.IsValid);
        GanttEvent resolved = Assert.Single(outcome.Events, e => e.Type == GanttEntityType.CriticalInterval);
        Assert.Equal("#FF0000", resolved.FillColour);
    }

    [Fact]
    public void Fill_on_procurement_is_a_capability_error()
    {
        GanttRowDto row = ValidSpan(typeText: "As-Built Procurement", fillColourText: "#FF0000");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "FillColour" && i.Code == GanttValidationCodes.ColourNotAllowedForType);
    }

    [Fact]
    public void Stroke_on_procurement_is_allowed()
    {
        GanttRowDto row = ValidSpan(typeText: "As-Built Procurement", strokeColourText: "#ff0000");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("-1")]
    [InlineData("1.5")]
    [InlineData("TRUE")]
    public void Bad_sort_order_is_a_blocking_error(string sortOrder)
    {
        GanttRowDto row = ValidSpan(sortOrder: sortOrder);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.BadSortOrder);
    }

    [Fact]
    public void Blank_description_is_allowed_for_every_type()
    {
        GanttRowDto[] rows = new[]
        {
            ValidSpan(rowNumber: 2, description: null),
            new GanttRowDto(
                3,
                NewId(),
                null,
                null,
                "As-Built Milestone",
                null,
                new DateOnly(2026, 9, 1),
                null,
                null,
                null,
                null,
                null,
                null,
                true,
                null
            ),
            new GanttRowDto(
                4,
                NewId(),
                null,
                null,
                "Delineator",
                null,
                new DateOnly(2026, 9, 1),
                null,
                null,
                null,
                "None",
                null,
                null,
                true,
                null
            ),
        };

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.True(outcome.IsValid);
        Assert.Equal(3, outcome.Events.Count);
    }

    [Fact]
    public void Blank_visible_resolves_to_true()
    {
        GanttRowDto row = ValidSpan() with { VisibleCell = GanttCells.Empty<bool?>() };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.True(Assert.Single(outcome.Events).Visible);
    }

    [Fact]
    public void Relevant_excel_error_is_one_blocking_issue_without_duplicate_required_issue()
    {
        GanttRowDto row = ValidSpan() with { StartCell = GanttCells.ExcelError<DateOnly?>(GanttExcelErrorCode.NotAvailable) };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Empty(outcome.Events);
        GanttValidationIssue issue = Assert.Single(
            outcome.Issues,
            i => i.Field == "Start" && i.Code == GanttValidationCodes.CellContainsExcelError
        );
        Assert.Equal(GanttValidationSeverity.Error, issue.Severity);
        Assert.DoesNotContain(outcome.Issues, i => i.Field == "Start" && i.Code == GanttValidationCodes.StartRequired);
    }

    [Fact]
    public void Relevant_unsupported_value_is_one_blocking_issue_without_duplicate_format_issue()
    {
        GanttRowDto row = ValidSpan() with { FillColourCell = GanttCells.Unsupported<string>() };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Empty(outcome.Events);
        Assert.Contains(outcome.Issues, i => i.Field == "FillColour" && i.Code == GanttValidationCodes.CellValueUnsupported);
        Assert.DoesNotContain(outcome.Issues, i => i.Field == "FillColour" && i.Code == GanttValidationCodes.BadColourFormat);
    }

    [Fact]
    public void Irrelevant_populated_cell_warns_once_and_is_not_a_format_error()
    {
        GanttRowDto row = ValidSpan(typeText: "As-Built Milestone") with
        {
            FinishCell = GanttCells.Value<DateOnly?>(new DateOnly(2026, 9, 9)),
        };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.Single(outcome.Events);
        GanttValidationIssue issue = Assert.Single(outcome.Issues, i => i.Field == "Finish");
        Assert.Equal(GanttValidationCodes.NotUsedByType, issue.Code);
        Assert.Equal(GanttValidationSeverity.Warning, issue.Severity);
        Assert.Null(Assert.Single(outcome.Events).Finish);
    }

    [Fact]
    public void Irrelevant_excel_error_warns_once_and_is_not_blocking()
    {
        GanttRowDto row = ValidSpan(typeText: "As-Built Milestone") with
        {
            FinishCell = GanttCells.ExcelError<DateOnly?>(GanttExcelErrorCode.Value),
        };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.True(outcome.IsValid);
        Assert.Single(outcome.Events);
        GanttValidationIssue issue = Assert.Single(outcome.Issues, i => i.Field == "Finish");
        Assert.Equal(GanttValidationCodes.NotUsedByType, issue.Code);
        Assert.Equal(GanttValidationSeverity.Warning, issue.Severity);
        Assert.Null(Assert.Single(outcome.Events).Finish);
    }

    [Fact]
    public void Critical_interval_with_invalid_parent_event_is_a_blocking_parent_invalid_issue()
    {
        var parentId = NewId();
        GanttRowDto parent = ValidSpan(rowNumber: 2, id: parentId, laneId: null);
        GanttRowDto child = ValidSpan(
            rowNumber: 3,
            typeText: "Critical Interval",
            id: NewId(),
            laneId: null,
            stackIndex: 0,
            start: new DateOnly(2026, 9, 2),
            finish: new DateOnly(2026, 9, 3),
            parentId: parentId
        ) with
        {
            LaneIdCell = GanttCells.Empty<string>(),
        };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.RowNumber == 3 && i.Code == GanttValidationCodes.ParentInvalid);
        Assert.DoesNotContain(outcome.Events, e => e.RowNumber == 3);
    }

    [Fact]
    public void Custom_activity_style_key_error_is_blocking_and_not_missing_style_key()
    {
        GanttRowDto row = ValidSpan(typeText: "Custom Activity", styleKey: "MyStyle") with
        {
            StyleKeyCell = GanttCells.ExcelError<string>(GanttExcelErrorCode.Value),
        };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Field == "StyleKey" && i.Code == GanttValidationCodes.CellContainsExcelError);
        Assert.DoesNotContain(outcome.Issues, i => i.Field == "StyleKey" && i.Code == GanttValidationCodes.StyleKeyRequired);
    }

    [Fact]
    public void All_errors_returns_every_independent_fault()
    {
        var row = new GanttRowDto(
            7,
            "BAD",
            null,
            null,
            "As-Planned Activity",
            null,
            new DateOnly(2026, 9, 5),
            new DateOnly(2026, 9, 1),
            null,
            null,
            "Sideways",
            "red",
            null,
            true,
            "abc"
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Empty(outcome.Events);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.IdMissingOrMalformed);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.LaneIdMissingOrMalformed);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.StackIndexRequired);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.StartAfterFinish);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.UnknownLabelPosition);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.BadColourFormat);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.BadSortOrder);
        Assert.True(outcome.Issues.Count >= 5);
    }

    [Fact]
    public void Unknown_type_still_reports_other_row_faults()
    {
        var row = new GanttRowDto(7, "BAD", null, null, "Widget", null, null, null, null, null, null, null, null, true, "abc");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Empty(outcome.Events);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.IdMissingOrMalformed);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.UnknownType);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.BadSortOrder);
    }

    [Fact]
    public void Issues_sort_by_severity_then_row_then_field_then_code()
    {
        var badRow = new GanttRowDto(5, "BAD", null, null, "Widget", null, null, null, null, null, null, null, null, true, "abc");
        var warnRow = new GanttRowDto(
            2,
            NewId(),
            null,
            null,
            "As-Built Milestone",
            null,
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 9),
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );

        GanttValidationOutcome outcome = GanttRowValidator.Validate([badRow, warnRow]);

        Assert.False(outcome.IsValid);
        GanttValidationIssue[] expected = outcome
            .Issues.OrderBy(i => i.Severity)
            .ThenBy(i => i.RowNumber)
            .ThenBy(i => i.Field, StringComparer.Ordinal)
            .ThenBy(i => i.Code, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(expected, outcome.Issues);
        // Errors precede the row-2 warning even though the warning has the smaller row number.
        Assert.Equal(GanttValidationSeverity.Error, outcome.Issues[0].Severity);
        Assert.Equal(GanttValidationSeverity.Warning, outcome.Issues[^1].Severity);
    }

    [Fact]
    public void Validation_is_deterministic_under_shuffled_input()
    {
        var parentId = NewId();
        GanttRowDto[] rows = new[]
        {
            ValidSpan(rowNumber: 4, id: parentId),
            ValidSpan(rowNumber: 2, id: "BAD"),
            new GanttRowDto(
                3,
                NewId(),
                null,
                null,
                "As-Built Milestone",
                null,
                new DateOnly(2026, 9, 1),
                new DateOnly(2026, 9, 9),
                null,
                null,
                null,
                null,
                null,
                true,
                null
            ),
        };

        GanttValidationOutcome first = GanttRowValidator.Validate(rows);
        GanttValidationOutcome shuffled = GanttRowValidator.Validate([rows[2], rows[0], rows[1]]);
        GanttValidationOutcome reversed = GanttRowValidator.Validate([rows[1], rows[2], rows[0]]);

        Assert.Equal(first.Issues, shuffled.Issues);
        Assert.Equal(first.Issues, reversed.Issues);
    }

    [Fact]
    public void Critical_interval_with_self_parent_is_a_blocking_cycle_error()
    {
        string id = NewId();
        GanttRowDto row = CriticalIntervalRow(2, id, id);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.Code == GanttValidationCodes.ParentCycle);
        Assert.Empty(outcome.Events);
    }

    [Fact]
    public void Critical_interval_with_two_node_parent_cycle_is_a_blocking_error()
    {
        string firstId = NewId();
        string secondId = NewId();
        GanttValidationOutcome outcome = GanttRowValidator.Validate([
            CriticalIntervalRow(2, firstId, secondId),
            CriticalIntervalRow(3, secondId, firstId),
        ]);

        Assert.False(outcome.IsValid);
        Assert.Equal(2, outcome.Issues.Count(i => i.Code == GanttValidationCodes.ParentCycle));
        Assert.Empty(outcome.Events);
    }

    [Fact]
    public void Critical_interval_with_longer_parent_cycle_is_a_blocking_error()
    {
        string firstId = NewId();
        string secondId = NewId();
        string thirdId = NewId();
        GanttValidationOutcome outcome = GanttRowValidator.Validate([
            CriticalIntervalRow(2, firstId, secondId),
            CriticalIntervalRow(3, secondId, thirdId),
            CriticalIntervalRow(4, thirdId, firstId),
        ]);

        Assert.False(outcome.IsValid);
        Assert.Equal(3, outcome.Issues.Count(i => i.Code == GanttValidationCodes.ParentCycle));
        Assert.Empty(outcome.Events);
    }

    [Fact]
    public void Critical_interval_child_of_an_ambiguous_interval_is_blocked_in_either_input_order()
    {
        // A Critical Interval can itself be a parent. The child is authored
        // above its parent here, so the direct parent pass decides the child
        // before the parent is known to be ambiguous; propagation must block it
        // anyway, and the same rows reversed must produce the same outcome.
        string ambiguousId = NewId();
        string childId = NewId();
        string duplicatedId = NewId();
        GanttRowDto child = CriticalIntervalRow(2, childId, ambiguousId);
        GanttRowDto ambiguousInterval = CriticalIntervalRow(3, ambiguousId, duplicatedId);
        GanttRowDto firstDuplicate = ValidSpan(rowNumber: 4, id: duplicatedId);
        GanttRowDto secondDuplicate = ValidSpan(rowNumber: 5, id: duplicatedId);

        GanttValidationOutcome childFirst = GanttRowValidator.Validate(
            [child, ambiguousInterval, firstDuplicate, secondDuplicate]);
        GanttValidationOutcome parentFirst = GanttRowValidator.Validate(
            [ambiguousInterval, child, firstDuplicate, secondDuplicate]);

        foreach (GanttValidationOutcome outcome in new[] { childFirst, parentFirst })
        {
            Assert.False(outcome.IsValid);
            Assert.Contains(outcome.Issues, i => i.RowNumber == 3 && i.Code == GanttValidationCodes.ParentAmbiguous);
            Assert.Contains(outcome.Issues, i => i.RowNumber == 2 && i.Code == GanttValidationCodes.ParentInvalid);
            Assert.DoesNotContain(outcome.Events, e => e.RowNumber is 2 or 3);
        }

        Assert.Equal(childFirst.Issues, parentFirst.Issues);

        // The canonical duplicate row stays a valid event; only the two
        // Critical Intervals must drop out.
        Assert.DoesNotContain(childFirst.Events, e => e.Type == GanttEntityType.CriticalInterval);
        Assert.DoesNotContain(parentFirst.Events, e => e.Type == GanttEntityType.CriticalInterval);
    }

    [Fact]
    public void Critical_interval_child_of_a_cycle_is_blocked_as_parent_invalid()
    {
        string firstId = NewId();
        string secondId = NewId();
        string childId = NewId();
        GanttRowDto child = CriticalIntervalRow(4, childId, firstId);

        GanttValidationOutcome outcome = GanttRowValidator.Validate([
            CriticalIntervalRow(2, firstId, secondId),
            CriticalIntervalRow(3, secondId, firstId),
            child,
        ]);

        Assert.False(outcome.IsValid);
        Assert.Contains(outcome.Issues, i => i.RowNumber == 4 && i.Code == GanttValidationCodes.ParentInvalid);
        Assert.DoesNotContain(outcome.Events, e => e.Id.Value == childId);
    }

    [Fact]
    public void Fill_on_a_critical_interval_is_accepted_and_reaches_the_event()
    {
        // ADR-0027 D5/D6. The entity renders as a filled rectangle, so a user
        // FillColour override is now permitted. This is the POSITIVE test for the
        // changed path: before the capability flip this row was a blocking error,
        // and a preset that fills while the validator refuses the user's own fill
        // is a defect that looks like a rendering choice.
        string parentId = NewId();
        string childId = NewId();
        GanttRowDto parent = ValidSpan(typeText: "As-Planned Activity") with { IdCell = GanttCells.Value(parentId) };
        GanttRowDto child = CriticalIntervalRow(2, childId, parentId, "#0000FF");

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.True(outcome.IsValid, "A fill override on a critical interval is now legal.");

        // And it must actually REACH the event, not merely stop being an error: a
        // silent drop here would be the same class of defect as a silent refusal.
        GanttEvent resolved = Assert.Single(outcome.Events, e => e.Id.Value == childId);
        Assert.Equal("#0000FF", resolved.FillColour);
    }

    [Fact]
    public void Stroke_on_a_critical_interval_stays_allowed()
    {
        // The capability gained Fill; Stroke must not have been lost in the flip.
        string parentId = NewId();
        GanttRowDto parent = ValidSpan(typeText: "As-Planned Activity") with { IdCell = GanttCells.Value(parentId) };
        GanttRowDto child = CriticalIntervalRow(2, NewId(), parentId) with
        {
            StrokeColourCell = GanttCells.Value("#FF00FF"),
        };

        GanttValidationOutcome outcome = GanttRowValidator.Validate([parent, child]);

        Assert.True(outcome.IsValid);
    }

    [Fact]
    public void Critical_interval_cycle_member_with_a_field_error_is_not_reported_as_a_cycle()
    {
        string firstId = NewId();
        string secondId = NewId();
        GanttValidationOutcome outcome = GanttRowValidator.Validate([
            // The field error this test needs used to be a FillColour override on a
            // critical interval, which ADR-0027 D5 made LEGAL. LabelPosition is
            // used instead: a critical interval still permits only `None`, so
            // "Left" is a real field error that must suppress the cycle report
            // without being misreported as one.
            CriticalIntervalRow(2, firstId, secondId) with
            {
                LabelPositionCell = GanttCells.Value("Left"),
            },
            CriticalIntervalRow(3, secondId, firstId),
        ]);

        Assert.False(outcome.IsValid);
        Assert.Contains(
            outcome.Issues,
            i => i.RowNumber == 2 && i.Code == GanttValidationCodes.LabelNotAllowedForType);
        Assert.DoesNotContain(
            outcome.Issues,
            i => i.RowNumber == 2 && i.Code == GanttValidationCodes.ParentCycle);
        Assert.Contains(outcome.Issues, i => i.RowNumber == 3 && i.Code == GanttValidationCodes.ParentInvalid);
    }

    [Fact]
    public void Critical_interval_cycle_reports_are_identical_under_shuffled_input()
    {
        string firstId = NewId();
        string secondId = NewId();
        string thirdId = NewId();
        GanttRowDto[] rows =
        [
            CriticalIntervalRow(2, firstId, secondId),
            CriticalIntervalRow(3, secondId, thirdId),
            CriticalIntervalRow(4, thirdId, firstId),
        ];

        GanttValidationOutcome ordered = GanttRowValidator.Validate(rows);
        GanttValidationOutcome shuffled = GanttRowValidator.Validate([rows[2], rows[0], rows[1]]);

        Assert.Equal(ordered.Issues, shuffled.Issues);
    }

    private static GanttRowDto CriticalIntervalRow(int rowNumber, string id, string parentId, string? fillColourText = null) =>
        new(
            rowNumber,
            id,
            NewLaneId(),
            0,
            "Critical Interval",
            "Critical part",
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 1),
            parentId,
            null,
            null,
            fillColourText,
            null,
            true,
            null
        );

    [Fact]
    public void Validation_preserves_exact_type_case_and_accepts_valid_label_under_tr_tr()
    {
        CultureInfo originalCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("tr-TR");

            // The #RRGGBB contract has no i/I, so this pins exact label and
            // type behaviour plus colour normalization; it does not claim to
            // distinguish Turkish casing for hex colours.
            GanttRowDto row = ValidSpan(fillColourText: "#ff0000", sortOrder: "3", labelPositionText: "Inside");
            GanttValidationOutcome outcome = GanttRowValidator.Validate([row]);

            Assert.True(outcome.IsValid);
            GanttEvent @event = Assert.Single(outcome.Events);
            Assert.Equal("#FF0000", @event.FillColour);
            Assert.Equal(GanttLabelPosition.Inside, @event.LabelPosition);
            Assert.False(GanttRowValidator.Validate([ValidSpan(typeText: "as-planned activity")]).IsValid);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }
}
