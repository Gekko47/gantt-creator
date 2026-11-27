using System.Globalization;

namespace GanttCreator.Core.Tests;

/// <summary>
/// R5.23: one Type edit must leave a row that is valid for the NEW type, on every
/// ordered pair of the 17 catalogue types, and the ribbon, validator and scene must read
/// the same healed row.
/// </summary>
/// <remarks>
/// The style registry is built from the real code-owned presets, so a style rule that
/// only works against a test double fails here. Only the label/vocabulary codes are
/// asserted, because hierarchy consequences (R7) are reported rather than repaired: a
/// converted parent keeps its children and the validator says so.
/// </remarks>
public sealed class EntityTypeTransitionTests
{
    private static readonly string[] _labelCodes =
    [
        GanttValidationCodes.LabelNotAllowedForType,
        GanttValidationCodes.UnknownLabelPosition,
        GanttValidationCodes.DateLabelNotAllowedForType,
        GanttValidationCodes.UnknownDateLabelPosition,
        GanttValidationCodes.StyleKeyRequired,
    ];

    [Fact]
    public void Every_ordered_type_pair_leaves_a_row_valid_for_the_new_type()
    {
        GanttStyleRegistry registry = PresetRegistry();
        List<string> failures = [];

        foreach (GanttEntityType oldType in Enum.GetValues<GanttEntityType>())
        {
            TypeTransitionRow seed = LegalRow(oldType, registry);
            foreach (GanttEntityType newType in Enum.GetValues<GanttEntityType>())
            {
                IReadOnlyList<TypeTransitionChange> changes =
                    EntityTypeTransition.Apply(seed, newType, oldType, registry);
                TypeTransitionRow healed = Healed(seed, changes);

                AssertValidPair(oldType, newType, healed, registry, failures);

                IReadOnlyList<TypeTransitionChange> second =
                    EntityTypeTransition.Apply(healed, newType, oldType, registry);
                if (second.Count != 0)
                {
                    failures.Add(
                        $"{oldType} -> {newType}: a second pass rewrote {string.Join(", ", second.Select(c => c.Column))}.");
                }
            }
        }

        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
    }

    [Fact]
    public void A_span_style_on_a_delineator_is_replaced_by_the_default()
    {
        GanttStyleRegistry registry = RegistryWithSpanStyle();
        TypeTransitionRow row = new(
            StyleKey: "SpanStyle",
            LabelPosition: null,
            DateLabelPosition: null,
            FillColour: null,
            StrokeColour: null,
            ParentId: null,
            LaneId: null,
            StackIndex: null,
            SiblingOrder: null);

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.Delineator, GanttEntityType.CustomActivity, registry);

        TypeTransitionChange change = Assert.Single(changes, c => c.Column == EntityTypeTransition.StyleKeyColumn);
        Assert.Equal("DefaultDelineator", change.NewValue);
    }

    [Fact]
    public void A_blank_style_key_on_a_requiring_type_gets_the_default()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { StyleKey = " " };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.Delineator, null, registry);

        Assert.Equal(
            "DefaultDelineator",
            Assert.Single(changes, c => c.Column == EntityTypeTransition.StyleKeyColumn).NewValue);
    }

    [Fact]
    public void The_shared_tokens_are_cleared_on_a_vocabulary_boundary_in_both_directions()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow fromDelineator = EmptyRow() with { DateLabelPosition = "None" };
        TypeTransitionRow fromSpan = EmptyRow() with { DateLabelPosition = "None" };

        IReadOnlyList<TypeTransitionChange> leaving =
            EntityTypeTransition.Apply(fromDelineator, GanttEntityType.AsPlannedActivity, GanttEntityType.Delineator, registry);
        IReadOnlyList<TypeTransitionChange> entering =
            EntityTypeTransition.Apply(fromSpan, GanttEntityType.Delineator, GanttEntityType.AsPlannedActivity, registry);

        Assert.Contains(leaving, c => c.Column == EntityTypeTransition.DateLabelPositionColumn);
        Assert.Contains(entering, c => c.Column == EntityTypeTransition.DateLabelPositionColumn);
    }

    [Fact]
    public void A_corner_left_on_a_span_is_cleared_without_an_old_type()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { DateLabelPosition = "TopLeft" };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.AsPlannedActivity, null, registry);

        Assert.Contains(changes, c => c.Column == EntityTypeTransition.DateLabelPositionColumn);
    }

    [Fact]
    public void A_valid_corner_on_a_delineator_survives_an_ordinary_sweep()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { DateLabelPosition = "TopLeft" };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.Delineator, null, registry);

        Assert.DoesNotContain(changes, c => c.Column == EntityTypeTransition.DateLabelPositionColumn);
    }

    [Fact]
    public void An_unparseable_date_label_is_left_alone()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { DateLabelPosition = "NotAPosition" };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.AsPlannedActivity, null, registry);

        Assert.DoesNotContain(changes, c => c.Column == EntityTypeTransition.DateLabelPositionColumn);
    }

    [Fact]
    public void A_dateless_type_clears_the_date_label()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { DateLabelPosition = "Start" };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.Splitter, null, registry);

        Assert.Contains(changes, c => c.Column == EntityTypeTransition.DateLabelPositionColumn);
    }

    [Fact]
    public void A_delineator_keeps_a_populated_lane_and_stack()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { LaneId = NewId('b'), StackIndex = "1" };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.Delineator, GanttEntityType.AsPlannedActivity, registry);

        Assert.DoesNotContain(changes, c => c.Column is EntityTypeTransition.LaneIdColumn or EntityTypeTransition.StackIndexColumn);
    }

    [Fact]
    public void A_spacer_still_clears_a_populated_lane_and_stack()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { LaneId = NewId('b'), StackIndex = "1" };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.Spacer, null, registry);

        Assert.Contains(changes, c => c.Column == EntityTypeTransition.LaneIdColumn);
        Assert.Contains(changes, c => c.Column == EntityTypeTransition.StackIndexColumn);
    }

    [Fact]
    public void A_fill_on_a_stroke_only_type_is_cleared_and_the_stroke_is_kept()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { FillColour = "#FFFFFF", StrokeColour = "#000000" };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.Delineator, null, registry);

        Assert.Contains(changes, c => c.Column == EntityTypeTransition.FillColourColumn);
        Assert.DoesNotContain(changes, c => c.Column == EntityTypeTransition.StrokeColourColumn);
    }

    [Fact]
    public void A_style_key_is_untouched_when_the_new_type_does_not_require_one()
    {
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow row = EmptyRow() with { StyleKey = "DefaultDelineator" };

        IReadOnlyList<TypeTransitionChange> changes =
            EntityTypeTransition.Apply(row, GanttEntityType.AsPlannedActivity, null, registry);

        Assert.DoesNotContain(changes, c => c.Column == EntityTypeTransition.StyleKeyColumn);
    }

    [Fact]
    public void A_span_style_is_usable_for_a_span_and_not_for_a_delineator()
    {
        GanttStyleRegistry registry = RegistryWithSpanStyle();

        Assert.True(EntityTypeCatalog.IsStyleKeyUsable(GanttEntityType.AsPlannedActivity, "SpanStyle", registry));
        Assert.False(EntityTypeCatalog.IsStyleKeyUsable(GanttEntityType.Delineator, "SpanStyle", registry));
    }

    [Fact]
    public void An_unknown_key_is_not_usable_and_a_missing_registry_keeps_the_choice()
    {
        GanttStyleRegistry registry = PresetRegistry();

        Assert.False(EntityTypeCatalog.IsStyleKeyUsable(GanttEntityType.Delineator, "Nope", registry));
        Assert.True(EntityTypeCatalog.IsStyleKeyUsable(GanttEntityType.Delineator, "Nope", null));
    }

    [Fact]
    public void A_style_with_only_shared_positions_is_usable_for_any_type()
    {
        // Auto and None are placeable by every builder, so a style that names only
        // those is neutral rather than a foreign family.
        GanttStyleRegistry registry = RegistryWithNeutralStyle();

        Assert.True(EntityTypeCatalog.IsStyleKeyUsable(GanttEntityType.Delineator, "NeutralStyle", registry));
        Assert.True(EntityTypeCatalog.IsStyleKeyUsable(GanttEntityType.AsPlannedActivity, "NeutralStyle", registry));
    }

    [Fact]
    public void A_custom_activity_healed_to_a_delineator_offers_only_corner_label_positions()
    {
        // The reported symptom, end to end within Core: the healed row's style resolves
        // to the delineator's corner set, so the ribbon cannot offer a span position and
        // the scene's corner builder has a position it can place.
        GanttStyleRegistry registry = PresetRegistry();
        TypeTransitionRow seed = new(
            StyleKey: "CustomActivity",
            LabelPosition: GanttLabelPosition.Left.ToString(),
            DateLabelPosition: GanttDateLabelPosition.Auto.ToString(),
            FillColour: "#FFFFFF",
            StrokeColour: "#000000",
            ParentId: null,
            LaneId: null,
            StackIndex: null,
            SiblingOrder: null);

        TypeTransitionRow healed = Healed(
            seed,
            EntityTypeTransition.Apply(
                seed, GanttEntityType.Delineator, GanttEntityType.CustomActivity, registry));

        IReadOnlyList<GanttLabelPosition> options = EntityTypeCatalog.GetAllowedLabelPositions(
            GanttEntityType.Delineator, healed.StyleKey, registry);

        Assert.NotEmpty(options);
        Assert.DoesNotContain(GanttLabelPosition.Left, options);
        foreach (GanttLabelPosition option in options)
        {
            Assert.True(
                option is GanttLabelPosition.Auto
                    or GanttLabelPosition.None
                    or GanttLabelPosition.TopLeft
                    or GanttLabelPosition.TopRight
                    or GanttLabelPosition.BottomLeft
                    or GanttLabelPosition.BottomRight,
                $"A delineator offered '{option}', which is not a corner.");
        }

        Assert.Null(healed.LabelPosition);
        Assert.Null(healed.DateLabelPosition);
    }

    private static GanttStyleRegistry RegistryWithNeutralStyle()
    {
        List<GanttStyleDefinition> definitions = PresetDefinitions();
        definitions.Add(new GanttStyleDefinition(
            "NeutralStyle",
            new HashSet<GanttLabelPosition>([GanttLabelPosition.Auto, GanttLabelPosition.None]),
            EntityColourCapability.Fill));
        return new GanttStyleRegistry(definitions);
    }

    private static void AssertValidPair(
        GanttEntityType oldType,
        GanttEntityType newType,
        TypeTransitionRow healed,
        GanttStyleRegistry registry,
        List<string> failures)
    {
        string pair = $"{oldType} -> {newType}";
        EntityTypeDefinition definition = EntityTypeCatalog.GetDefinition(newType)!;

        if (definition.RequiresStyleKey)
        {
            if (string.IsNullOrEmpty(healed.StyleKey))
            {
                failures.Add($"{pair}: the required style key is blank.");
            }
            else if (!EntityTypeCatalog.IsStyleKeyUsable(newType, healed.StyleKey, registry))
            {
                failures.Add($"{pair}: style key '{healed.StyleKey}' is not usable.");
            }
        }

        if (!string.IsNullOrEmpty(healed.LabelPosition))
        {
            bool parsed = Enum.TryParse(healed.LabelPosition, ignoreCase: false, out GanttLabelPosition position)
                && Enum.IsDefined(position);
            if (!parsed || !EntityTypeCatalog.IsLabelPositionAllowed(newType, position, healed.StyleKey, registry))
            {
                failures.Add($"{pair}: label position '{healed.LabelPosition}' is not allowed.");
            }
        }

        if (!string.IsNullOrEmpty(healed.DateLabelPosition))
        {
            bool valid = newType == GanttEntityType.Delineator
                ? Enum.TryParse(healed.DateLabelPosition, ignoreCase: false, out GanttLabelPosition placement)
                    && Enum.IsDefined(placement)
                    && EntityTypeCatalog.IsDateLabelPlacementAllowed(newType, placement)
                : Enum.TryParse(healed.DateLabelPosition, ignoreCase: false, out GanttDateLabelPosition which)
                    && Enum.IsDefined(which)
                    && EntityTypeCatalog.IsDateLabelPositionAllowed(newType, which);

            if (!valid)
            {
                failures.Add($"{pair}: date label '{healed.DateLabelPosition}' is not valid for the new vocabulary.");
            }
        }

        GanttValidationOutcome outcome = GanttRowValidator.Validate(
            [Dto(newType, healed)], registry);
        foreach (GanttValidationIssue issue in outcome.Issues)
        {
            if (Array.IndexOf(_labelCodes, issue.Code) >= 0)
            {
                failures.Add($"{pair}: {issue.Code} on {issue.Field}: {issue.Message}");
            }
        }
    }

    private static GanttRowDto Dto(GanttEntityType type, TypeTransitionRow row)
    {
        EntityTypeDefinition definition = EntityTypeCatalog.GetDefinition(type)!;
        DateOnly? start = definition.DateMode == EntityDateMode.None ? null : new DateOnly(2030, 1, 7);
        DateOnly? finish = definition.DateMode == EntityDateMode.StartFinish ? new DateOnly(2030, 1, 11) : null;
        int? stackIndex = int.TryParse(
            row.StackIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : null;

        return new GanttRowDto(
            1,
            NewId('a'),
            row.LaneId,
            stackIndex,
            definition.DisplayName,
            "Row",
            start,
            finish,
            row.ParentId,
            row.StyleKey,
            row.LabelPosition,
            row.FillColour,
            row.StrokeColour,
            true,
            null,
            row.DateLabelPosition);
    }

    private static TypeTransitionRow LegalRow(GanttEntityType type, GanttStyleRegistry registry)
    {
        EntityTypeDefinition definition = EntityTypeCatalog.GetDefinition(type)!;
        string? styleKey = definition.RequiresStyleKey ? definition.DefaultStyleKey : null;
        IReadOnlyList<GanttLabelPosition> labels =
            EntityTypeCatalog.GetAllowedLabelPositions(type, styleKey, registry);
        string? label = labels.Count > 0 ? labels[0].ToString() : null;

        string? dateLabel = null;
        if (type == GanttEntityType.Delineator)
        {
            dateLabel = EntityTypeCatalog.GetAllowedDateLabelPlacements(type)[0].ToString();
        }
        else
        {
            IReadOnlyList<GanttDateLabelPosition> which =
                EntityTypeCatalog.GetAllowedDateLabelPositions(type);
            if (which.Count > 0)
            {
                dateLabel = which[0].ToString();
            }
        }

        bool usesLane = definition.Kind is EntityKind.Span or EntityKind.Milestone;
        return new TypeTransitionRow(
            StyleKey: styleKey,
            LabelPosition: label,
            DateLabelPosition: dateLabel,
            FillColour: Has(definition, EntityColourCapability.Fill) ? "#FFFFFF" : null,
            StrokeColour: Has(definition, EntityColourCapability.Stroke) ? "#000000" : null,
            ParentId: EntityHierarchyCatalog.MayBeChild(type) ? NewId('b') : null,
            LaneId: usesLane ? NewId('c') : null,
            StackIndex: usesLane ? "1" : null,
            SiblingOrder: EntityHierarchyCatalog.MayBeChild(type) ? "1" : null);
    }

    private static TypeTransitionRow Healed(
        TypeTransitionRow row,
        IReadOnlyList<TypeTransitionChange> changes)
    {
        TypeTransitionRow next = row;
        foreach (TypeTransitionChange change in changes)
        {
            // A cleared cell reads back as null, exactly as the worksheet reader reports a
            // blank cell; keeping "" would make the validator treat it as present-and-bad.
            string? value = change.NewValue.Length == 0 ? null : change.NewValue;
            next = change.Column switch
            {
                EntityTypeTransition.StyleKeyColumn => next with { StyleKey = value },
                EntityTypeTransition.LabelPositionColumn => next with { LabelPosition = value },
                EntityTypeTransition.DateLabelPositionColumn => next with { DateLabelPosition = value },
                EntityTypeTransition.FillColourColumn => next with { FillColour = value },
                EntityTypeTransition.StrokeColourColumn => next with { StrokeColour = value },
                EntityTypeTransition.ParentIdColumn => next with { ParentId = value },
                EntityTypeTransition.LaneIdColumn => next with { LaneId = value },
                EntityTypeTransition.StackIndexColumn => next with { StackIndex = value },
                EntityTypeTransition.SiblingOrderColumn => next with { SiblingOrder = value },
                _ => throw new InvalidOperationException($"Unknown transition column '{change.Column}'."),
            };
        }

        return next;
    }

    private static TypeTransitionRow EmptyRow() =>
        new(null, null, null, null, null, null, null, null, null);

    private static string NewId(char last) => "G-" + new string('0', 31) + last;

    private static bool Has(EntityTypeDefinition definition, EntityColourCapability capability) =>
        (definition.ColourCapability & capability) != 0;

    private static GanttStyleRegistry PresetRegistry() =>
        new(PresetDefinitions());

    private static GanttStyleRegistry RegistryWithSpanStyle()
    {
        List<GanttStyleDefinition> definitions = PresetDefinitions();
        definitions.Add(new GanttStyleDefinition(
            "SpanStyle",
            new HashSet<GanttLabelPosition>(
            [
                GanttLabelPosition.Auto,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.Inside,
                GanttLabelPosition.None,
            ]),
            EntityColourCapability.Fill | EntityColourCapability.Stroke));
        return new GanttStyleRegistry(definitions);
    }

    private static List<GanttStyleDefinition> PresetDefinitions()
    {
        List<GanttStyleDefinition> definitions = [];
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            // Capability-only definitions: the transition reads only AllowedLabelPositions,
            // and passing a DefaultLabelPosition would make the registry demand the full
            // formatting block (hatch, metrics) that this test does not exercise.
            definitions.Add(new GanttStyleDefinition(
                preset.StyleKey,
                preset.AllowedLabelPositions,
                preset.ColourCapability));
        }

        return definitions;
    }
}
