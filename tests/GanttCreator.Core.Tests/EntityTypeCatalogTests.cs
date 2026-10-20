using System.Globalization;

namespace GanttCreator.Core.Tests;

public class EntityTypeCatalogTests
{
    private static readonly string[] s_guideDisplayNames =
    [
        "Splitter",
        "Spacer",
        "As-Built Activity",
        "As-Planned Activity",
        "Baseline Activity",
        "Critical Interval",
        "Delay Event",
        "As-Built Procurement",
        "As-Planned Procurement",
        "Baseline Procurement",
        "Custom Activity",
        "As-Built Milestone",
        "As-Planned Milestone",
        "Baseline Milestone",
        "Critical Milestone",
        "Custom Milestone",
        "Delineator",
    ];

    private static readonly GanttLabelPosition[] s_spanLabels =
        [GanttLabelPosition.Auto, GanttLabelPosition.Left, GanttLabelPosition.Right, GanttLabelPosition.Inside, GanttLabelPosition.None];

    private static readonly GanttLabelPosition[] s_milestoneLabels =
        [GanttLabelPosition.Auto, GanttLabelPosition.Left, GanttLabelPosition.Right, GanttLabelPosition.None];

    private static readonly GanttLabelPosition[] s_delineatorLabels =
        [GanttLabelPosition.Auto, GanttLabelPosition.TopLeft, GanttLabelPosition.TopRight, GanttLabelPosition.BottomLeft, GanttLabelPosition.BottomRight, GanttLabelPosition.None];

    [Fact]
    public void Default_catalogue_contains_exactly_the_17_guide_entries()
    {
        Assert.Equal(17, EntityTypeCatalog.Entries.Count);
        Assert.Equal(EntityTypeCatalog.TypeCount, EntityTypeCatalog.Entries.Count);
    }

    [Fact]
    public void Every_enum_value_has_a_catalogue_definition_and_counts_match()
    {
        var enumValues = Enum.GetValues<GanttEntityType>();
        Assert.Equal(enumValues.Length, EntityTypeCatalog.Entries.Count);

        foreach (var value in enumValues)
        {
            var definition = EntityTypeCatalog.GetDefinition(value);
            Assert.NotNull(definition);
            Assert.Equal(value, definition!.Type);
        }
    }

    [Fact]
    public void Default_catalogue_matches_the_entity_guide_display_names_in_order()
    {
        Assert.Equal(17, s_guideDisplayNames.Length);
        for (var i = 0; i < s_guideDisplayNames.Length; i++)
        {
            Assert.Equal(s_guideDisplayNames[i], EntityTypeCatalog.Entries[i].DisplayName);
        }
    }

    [Fact]
    public void Display_names_and_types_are_unique_in_the_catalogue()
    {
        Assert.Equal(
            EntityTypeCatalog.Entries.Count,
            EntityTypeCatalog.Entries.Select(e => e.DisplayName).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(
            EntityTypeCatalog.Entries.Count,
            EntityTypeCatalog.Entries.Select(e => e.Type).Distinct().Count());
    }

    [Fact]
    public void Default_catalogue_metadata_matches_the_entity_guide()
    {
        var expected = new (GanttEntityType Type, string DisplayName, EntityKind Kind, EntityDateMode DateMode, string DefaultStyleKey, EntityColourCapability ColourCapability, bool RequiresStyleKey, GanttLabelPosition[] Labels)[]
        {
            (GanttEntityType.Splitter, "Splitter", EntityKind.SectionHeader, EntityDateMode.None, "Splitter", EntityColourCapability.Fill, false, [GanttLabelPosition.DataPanelLeft, GanttLabelPosition.PlotCentre, GanttLabelPosition.Both, GanttLabelPosition.None]),
            (GanttEntityType.Spacer, "Spacer", EntityKind.Spacer, EntityDateMode.None, "Spacer", EntityColourCapability.None, false, [GanttLabelPosition.None]),
            (GanttEntityType.AsBuiltActivity, "As-Built Activity", EntityKind.Span, EntityDateMode.StartFinish, "AsBuiltActivity", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.AsPlannedActivity, "As-Planned Activity", EntityKind.Span, EntityDateMode.StartFinish, "AsPlannedActivity", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.BaselineActivity, "Baseline Activity", EntityKind.Span, EntityDateMode.StartFinish, "BaselineActivity", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.CriticalInterval, "Critical Interval", EntityKind.Span, EntityDateMode.StartFinish, "CriticalInterval", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, [GanttLabelPosition.None]),
            (GanttEntityType.DelayEvent, "Delay Event", EntityKind.Span, EntityDateMode.StartFinish, "DelayEvent", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.AsBuiltProcurement, "As-Built Procurement", EntityKind.Span, EntityDateMode.StartFinish, "AsBuiltProcurement", EntityColourCapability.Hatch | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.AsPlannedProcurement, "As-Planned Procurement", EntityKind.Span, EntityDateMode.StartFinish, "AsPlannedProcurement", EntityColourCapability.Hatch | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.BaselineProcurement, "Baseline Procurement", EntityKind.Span, EntityDateMode.StartFinish, "BaselineProcurement", EntityColourCapability.Hatch | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.CustomActivity, "Custom Activity", EntityKind.Span, EntityDateMode.StartFinish, "CustomActivity", EntityColourCapability.StyleDefined, true, s_spanLabels),
            (GanttEntityType.AsBuiltMilestone, "As-Built Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "AsBuiltMilestone", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_milestoneLabels),
            (GanttEntityType.AsPlannedMilestone, "As-Planned Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "AsPlannedMilestone", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_milestoneLabels),
            (GanttEntityType.BaselineMilestone, "Baseline Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "BaselineMilestone", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_milestoneLabels),
            (GanttEntityType.CriticalMilestone, "Critical Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "CriticalMilestone", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_milestoneLabels),
            (GanttEntityType.CustomMilestone, "Custom Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "CustomMilestone", EntityColourCapability.StyleDefined, true, s_milestoneLabels),
            (GanttEntityType.Delineator, "Delineator", EntityKind.Delineator, EntityDateMode.StartOnly, "DefaultDelineator", EntityColourCapability.Stroke, false, s_delineatorLabels),
        };

        Assert.Equal(17, expected.Length);

        // Indexed comparison also pins the catalogue to the exact guide order.
        for (var i = 0; i < expected.Length; i++)
        {
            var e = expected[i];
            var entry = EntityTypeCatalog.Entries[i];

            Assert.Equal(e.Type, entry.Type);
            Assert.Equal(e.DisplayName, entry.DisplayName);
            Assert.Equal(e.Kind, entry.Kind);
            Assert.Equal(e.DateMode, entry.DateMode);
            Assert.Equal(e.DefaultStyleKey, entry.DefaultStyleKey);
            Assert.Equal(e.ColourCapability, entry.ColourCapability);
            Assert.Equal(e.RequiresStyleKey, entry.RequiresStyleKey);
            Assert.Equal(e.Labels.OrderBy(p => p), entry.AllowedLabelPositions.OrderBy(p => p));
        }
    }

    [Fact]
    public void Format_then_parse_round_trips_every_catalogue_entry()
    {
        foreach (var entry in EntityTypeCatalog.Entries)
        {
            var formatted = EntityTypeCatalog.Format(entry.Type);

            Assert.Equal(entry.DisplayName, formatted);
            Assert.True(EntityTypeCatalog.TryParse(formatted, out var parsed));
            Assert.Equal(entry.Type, parsed);
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Widget")]
    [InlineData("as-built activity")]
    [InlineData("AS-BUILT ACTIVITY")]
    [InlineData("As-Built  Activity")]
    [InlineData("As-Built-Activity")]
    public void TryParse_rejects_text_that_does_not_match_a_display_name(string? text)
    {
        Assert.False(EntityTypeCatalog.TryParse(text, out var parsed));
        Assert.Equal(default(GanttEntityType), parsed);
    }

    [Theory]
    [InlineData(" As-Built Activity")]
    [InlineData("Delineator\t")]
    [InlineData("  Custom Activity  ")]
    public void TryParse_matches_after_trimming_surrounding_whitespace(string text)
    {
        Assert.True(EntityTypeCatalog.TryParse(text, out var parsed));
        Assert.Equal(EntityTypeCatalog.GetDefinition(text)!.Type, parsed);
    }

    [Fact]
    public void GetDefinition_returns_null_for_unknown_input_and_the_entry_for_exact_input()
    {
        Assert.Null(EntityTypeCatalog.GetDefinition((GanttEntityType)999));
        Assert.Null(EntityTypeCatalog.GetDefinition("Widget"));
        Assert.Null(EntityTypeCatalog.GetDefinition(null));
        Assert.Null(EntityTypeCatalog.GetDefinition(""));

        var delineator = EntityTypeCatalog.GetDefinition("  Delineator ");
        Assert.NotNull(delineator);
        Assert.Equal(GanttEntityType.Delineator, delineator.Type);
    }

    [Fact]
    public void Format_throws_for_a_type_outside_the_catalogue()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => EntityTypeCatalog.Format((GanttEntityType)999));
    }

    [Fact]
    public void Only_custom_types_require_a_style_key()
    {
        Assert.All(EntityTypeCatalog.Entries, entry =>
            Assert.Equal(
                entry.Type == GanttEntityType.CustomActivity
                    || entry.Type == GanttEntityType.CustomMilestone,
                entry.RequiresStyleKey));
    }

    [Fact]
    public void Custom_types_have_a_catalogue_default_style_and_their_family_label_set()
    {
        // R5.6a: a style-key type's catalogue label set is the family set, not
        // a {None} placeholder. The row's own named style is the authority for
        // the label positions it allows, but the catalogue set is what the
        // type's code-owned default preset carries, so a healed row (one whose
        // user-created style was deleted) can still render a label.
        foreach (var type in new[] { GanttEntityType.CustomActivity, GanttEntityType.CustomMilestone })
        {
            var custom = Assert.Single(EntityTypeCatalog.Entries, e => e.Type == type);

            string expectedKey = type == GanttEntityType.CustomActivity ? "CustomActivity" : "CustomMilestone";
            Assert.Equal(expectedKey, custom.DefaultStyleKey);
            Assert.True(custom.RequiresStyleKey);
            Assert.Equal(
                type == GanttEntityType.CustomActivity
                    ? new[] { GanttLabelPosition.None, GanttLabelPosition.Auto, GanttLabelPosition.Left, GanttLabelPosition.Right, GanttLabelPosition.Inside }
                    : new[] { GanttLabelPosition.None, GanttLabelPosition.Auto, GanttLabelPosition.Left, GanttLabelPosition.Right },
                custom.AllowedLabelPositions.OrderBy(p => p));
        }
    }

    [Fact]
    public void Milestones_and_delineators_read_start_only()
    {
        Assert.All(EntityTypeCatalog.Entries, entry =>
        {
            var expected = entry.Type switch
            {
                GanttEntityType.AsBuiltMilestone
                    or GanttEntityType.AsPlannedMilestone
                    or GanttEntityType.BaselineMilestone
                    or GanttEntityType.CriticalMilestone
                    or GanttEntityType.CustomMilestone
                    or GanttEntityType.Delineator => EntityDateMode.StartOnly,
                GanttEntityType.Splitter or GanttEntityType.Spacer => EntityDateMode.None,
                _ => EntityDateMode.StartFinish,
            };
            Assert.Equal(expected, entry.DateMode);
        });
    }

    [Fact]
    public void CriticalInterval_permits_only_None_label_position()
    {
        var critical = Assert.Single(EntityTypeCatalog.Entries, e => e.Type == GanttEntityType.CriticalInterval);

        Assert.Equal([GanttLabelPosition.None], critical.AllowedLabelPositions);
    }

    [Fact]
    public void Catalogue_surfaces_are_culture_invariant_under_tr_tr()
    {
        var originalCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("tr-TR");

            foreach (var entry in EntityTypeCatalog.Entries)
            {
                Assert.True(EntityTypeCatalog.TryParse(EntityTypeCatalog.Format(entry.Type), out var parsed));
                Assert.Equal(entry.Type, parsed);
            }

            Assert.False(EntityTypeCatalog.TryParse("as-built activity", out _));
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }

    [Fact]
    public void GetAllowedLabelPositions_projects_span_set_in_display_order()
    {
        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(GanttEntityType.AsBuiltActivity);

        Assert.Equal(
            new[]
            {
                GanttLabelPosition.Auto,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.Inside,
                GanttLabelPosition.None,
            },
            positions);
    }

    [Fact]
    public void GetAllowedLabelPositions_projects_milestone_set_without_inside()
    {
        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(GanttEntityType.AsBuiltMilestone);

        Assert.Equal(
            new[]
            {
                GanttLabelPosition.Auto,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.None,
            },
            positions);
    }

    [Fact]
    public void GetAllowedLabelPositions_projects_delineator_corners_in_display_order()
    {
        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(GanttEntityType.Delineator);

        Assert.Equal(
            new[]
            {
                GanttLabelPosition.Auto,
                GanttLabelPosition.TopLeft,
                GanttLabelPosition.TopRight,
                GanttLabelPosition.BottomLeft,
                GanttLabelPosition.BottomRight,
                GanttLabelPosition.None,
            },
            positions);
    }

    [Fact]
    public void GetAllowedLabelPositions_projects_splitter_set_without_auto()
    {
        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(GanttEntityType.Splitter);

        Assert.Equal(
            new[]
            {
                GanttLabelPosition.DataPanelLeft,
                GanttLabelPosition.PlotCentre,
                GanttLabelPosition.Both,
                GanttLabelPosition.None,
            },
            positions);
    }

    [Theory]
    [InlineData(GanttEntityType.Spacer)]
    [InlineData(GanttEntityType.CriticalInterval)]
    public void GetAllowedLabelPositions_single_value_types_permit_only_none(GanttEntityType type)
    {
        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(type);

        Assert.Equal(new[] { GanttLabelPosition.None }, positions);
    }

    [Fact]
    public void GetAllowedLabelPositions_custom_activity_resolves_from_its_style()
    {
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                "MyStyle",
                new HashSet<GanttLabelPosition>(
                    [GanttLabelPosition.Left, GanttLabelPosition.Right, GanttLabelPosition.None]),
                EntityColourCapability.StyleDefined),
        ]);

        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(
                GanttEntityType.CustomActivity, "MyStyle", registry);

        Assert.Equal(
            new[]
            {
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.None,
            },
            positions);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("Missing")]
    public void GetAllowedLabelPositions_custom_activity_without_style_falls_back_to_the_type_default(
        string? styleKey)
    {
        // R5.6d: a row whose style the registry does not know still renders -
        // the scene falls back to the type's code-owned default style - so the
        // ribbon must offer the positions that default permits. Returning an
        // empty set left the description-label dropdown disabled on a
        // perfectly renderable row.
        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(
                GanttEntityType.CustomActivity, styleKey, GanttStyleRegistry.Empty);

        Assert.Equal(
            new[]
            {
                GanttLabelPosition.Auto,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.Inside,
                GanttLabelPosition.None,
            },
            positions);
    }

    [Fact]
    public void GetAllowedLabelPositions_custom_milestone_resolves_from_its_style()
    {
        // R5.6a regression: Custom Milestone is a style-key type, so its
        // positions come from the row style exactly as a Custom Activity's
        // do — the catalogue's placeholder set is never the authority.
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                "MyMilestoneStyle",
                new HashSet<GanttLabelPosition>(
                    [GanttLabelPosition.Left, GanttLabelPosition.TopRight]),
                EntityColourCapability.StyleDefined),
        ]);

        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(
                GanttEntityType.CustomMilestone, "MyMilestoneStyle", registry);

        Assert.Equal(
            new[] { GanttLabelPosition.Left, GanttLabelPosition.TopRight },
            positions);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("Missing")]
    public void GetAllowedLabelPositions_custom_milestone_without_style_falls_back_to_the_type_default(
        string? styleKey)
    {
        // R5.6d: the milestone fallback is the milestone vocabulary, so a
        // Custom Milestone row that has not chosen a style still offers the
        // milestone positions (and no Inside, which a milestone has not).
        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(
                GanttEntityType.CustomMilestone, styleKey, GanttStyleRegistry.Empty);

        Assert.Equal(
            new[]
            {
                GanttLabelPosition.Auto,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.None,
            },
            positions);
    }

    [Fact]
    public void GetAllowedLabelPositions_a_known_style_wins_over_the_type_default()
    {
        // R5.6d: the fallback covers a style the registry does not KNOW. A
        // style it does know is the authority, exactly as before.
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                "NarrowStyle",
                new HashSet<GanttLabelPosition>([GanttLabelPosition.TopRight]),
                EntityColourCapability.StyleDefined),
        ]);

        IReadOnlyList<GanttLabelPosition> positions =
            EntityTypeCatalog.GetAllowedLabelPositions(
                GanttEntityType.CustomActivity, "NarrowStyle", registry);

        Assert.Equal(new[] { GanttLabelPosition.TopRight }, positions);
    }

    [Fact]
    public void IsLabelPositionAllowed_custom_milestone_delegates_to_its_style()
    {
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                "MyMilestoneStyle",
                new HashSet<GanttLabelPosition>([GanttLabelPosition.TopRight]),
                EntityColourCapability.StyleDefined),
        ]);

        Assert.True(EntityTypeCatalog.IsLabelPositionAllowed(
            GanttEntityType.CustomMilestone, GanttLabelPosition.TopRight, "MyMilestoneStyle", registry));
        Assert.False(EntityTypeCatalog.IsLabelPositionAllowed(
            GanttEntityType.CustomMilestone, GanttLabelPosition.Left, "MyMilestoneStyle", registry));
    }

    [Fact]
    public void GetAllowedLabelPositions_unknown_type_is_empty()
    {
        Assert.Empty(EntityTypeCatalog.GetAllowedLabelPositions((GanttEntityType)999));
    }

    [Theory]
    [InlineData(GanttEntityType.AsBuiltActivity, GanttLabelPosition.TopLeft, false)]
    [InlineData(GanttEntityType.AsBuiltMilestone, GanttLabelPosition.Inside, false)]
    [InlineData(GanttEntityType.Splitter, GanttLabelPosition.Left, false)]
    [InlineData(GanttEntityType.Spacer, GanttLabelPosition.Left, false)]
    [InlineData(GanttEntityType.CriticalInterval, GanttLabelPosition.Right, false)]
    [InlineData(GanttEntityType.AsBuiltActivity, GanttLabelPosition.Right, true)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.BottomRight, true)]
    [InlineData(GanttEntityType.Splitter, GanttLabelPosition.Both, true)]
    [InlineData(GanttEntityType.Spacer, GanttLabelPosition.None, true)]
    public void IsLabelPositionAllowed_refuses_cross_kind_positions(
        GanttEntityType type, GanttLabelPosition position, bool expected)
    {
        Assert.Equal(expected, EntityTypeCatalog.IsLabelPositionAllowed(type, position));
    }

    [Fact]
    public void TryParseLabelSelection_default_clears_to_null()
    {
        Assert.True(EntityTypeCatalog.TryParseLabelSelection("Default", out GanttLabelPosition? position));
        Assert.Null(position);
    }

    [Theory]
    [InlineData("Left", GanttLabelPosition.Left)]
    [InlineData("Inside", GanttLabelPosition.Inside)]
    [InlineData("TopRight", GanttLabelPosition.TopRight)]
    [InlineData("PlotCentre", GanttLabelPosition.PlotCentre)]
    [InlineData("None", GanttLabelPosition.None)]
    public void TryParseLabelSelection_parses_defined_positions(
        string selectionId, GanttLabelPosition expected)
    {
        Assert.True(EntityTypeCatalog.TryParseLabelSelection(selectionId, out GanttLabelPosition? position));
        Assert.Equal(expected, position);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Widget")]
    [InlineData("left")]
    [InlineData("Centre")]
    [InlineData("Above")]
    public void TryParseLabelSelection_refuses_unknown_text(string? selectionId)
    {
        Assert.False(EntityTypeCatalog.TryParseLabelSelection(selectionId, out GanttLabelPosition? position));
        Assert.Null(position);
    }

    [Fact]
    public void GetAllowedDateLabelPositions_projects_span_set_in_date_order()
    {
        IReadOnlyList<GanttDateLabelPosition> choices =
            EntityTypeCatalog.GetAllowedDateLabelPositions(GanttEntityType.AsPlannedActivity);

        Assert.Equal(
            new[]
            {
                GanttDateLabelPosition.Auto,
                GanttDateLabelPosition.Start,
                GanttDateLabelPosition.Finish,
                GanttDateLabelPosition.None,
            },
            choices);
    }

    [Theory]
    [InlineData(GanttEntityType.AsBuiltMilestone)]
    public void GetAllowedDateLabelPositions_projects_start_only_types_without_finish(
        GanttEntityType type)
    {
        IReadOnlyList<GanttDateLabelPosition> choices =
            EntityTypeCatalog.GetAllowedDateLabelPositions(type);

        Assert.Equal(
            new[]
            {
                GanttDateLabelPosition.Auto,
                GanttDateLabelPosition.Start,
                GanttDateLabelPosition.None,
            },
            choices);
    }

    [Theory]
    [InlineData(GanttEntityType.Splitter)]
    [InlineData(GanttEntityType.Spacer)]
    public void GetAllowedDateLabelPositions_projects_dateless_types_as_empty(
        GanttEntityType type)
    {
        Assert.Empty(EntityTypeCatalog.GetAllowedDateLabelPositions(type));
    }

    [Fact]
    public void GetAllowedDateLabelPositions_unknown_type_is_empty()
    {
        Assert.Empty(EntityTypeCatalog.GetAllowedDateLabelPositions((GanttEntityType)999));
    }

    [Fact]
    public void GetAllowedDateLabelPositions_projects_the_delineator_as_empty()
    {
        // Owner ruling 2026-10-14: the delineator's date label is
        // positionable, not selectable per date — Start is its only
        // date — so its which-date set is empty and the positionable
        // set below carries it instead.
        Assert.Empty(EntityTypeCatalog.GetAllowedDateLabelPositions(GanttEntityType.Delineator));
    }

    [Fact]
    public void GetAllowedDateLabelPlacements_projects_the_corner_set_for_the_delineator_only()
    {
        // The description label's corner set, in display order, for
        // the one type whose date label is positionable (owner
        // ruling 2026-10-14).
        Assert.Equal(
            new[]
            {
                GanttLabelPosition.Auto,
                GanttLabelPosition.TopLeft,
                GanttLabelPosition.TopRight,
                GanttLabelPosition.BottomLeft,
                GanttLabelPosition.BottomRight,
                GanttLabelPosition.None,
            },
            EntityTypeCatalog.GetAllowedDateLabelPlacements(GanttEntityType.Delineator));

        // Every other type's date label selects which date shows,
        // not where it sits, so the positionable set is empty for
        // them — a span's Left/Right anchoring belongs to the
        // description label, not to a date label.
        List<GanttEntityType> nonEmpty =
        [
            .. Enum.GetValues<GanttEntityType>()
                .Where(type => type != GanttEntityType.Delineator)
                .Where(type => EntityTypeCatalog.GetAllowedDateLabelPlacements(type).Count > 0),
        ];
        Assert.Empty(nonEmpty);
    }

    [Theory]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.Auto, true)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.TopLeft, true)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.TopRight, true)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.BottomLeft, true)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.BottomRight, true)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.None, true)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.Left, false)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.Right, false)]
    [InlineData(GanttEntityType.Delineator, GanttLabelPosition.Inside, false)]
    [InlineData(GanttEntityType.AsPlannedActivity, GanttLabelPosition.TopLeft, false)]
    public void IsDateLabelPlacementAllowed_accepts_only_the_delineators_corners(
        GanttEntityType type,
        GanttLabelPosition position,
        bool expected)
    {
        Assert.Equal(expected, EntityTypeCatalog.IsDateLabelPlacementAllowed(type, position));
    }

    [Theory]
    [InlineData("Default")]
    public void TryParseDateLabelPlacementSelection_default_clears_to_null(string selectionId)
    {
        Assert.True(
            EntityTypeCatalog.TryParseDateLabelPlacementSelection(selectionId, out GanttLabelPosition? placement));
        Assert.Null(placement);
    }

    [Theory]
    [InlineData("Auto", GanttLabelPosition.Auto)]
    [InlineData("TopLeft", GanttLabelPosition.TopLeft)]
    [InlineData("TopRight", GanttLabelPosition.TopRight)]
    [InlineData("BottomLeft", GanttLabelPosition.BottomLeft)]
    [InlineData("BottomRight", GanttLabelPosition.BottomRight)]
    [InlineData("None", GanttLabelPosition.None)]
    public void TryParseDateLabelPlacementSelection_parses_defined_positions(
        string selectionId,
        GanttLabelPosition expected)
    {
        Assert.True(
            EntityTypeCatalog.TryParseDateLabelPlacementSelection(selectionId, out GanttLabelPosition? placement));
        Assert.Equal(expected, placement);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Widget")]
    // "Start" is a valid DateLabelPosition but NOT a valid
    // GanttLabelPosition, so this pins that the two enumerations
    // are distinct authorities.
    [InlineData("Start")]
    public void TryParseDateLabelPlacementSelection_refuses_unknown_text(string? selectionId)
    {
        Assert.False(
            EntityTypeCatalog.TryParseDateLabelPlacementSelection(selectionId, out GanttLabelPosition? placement));
        Assert.Null(placement);
    }

    [Theory]
    [InlineData(GanttEntityType.AsPlannedActivity, GanttDateLabelPosition.Finish, true)]
    [InlineData(GanttEntityType.AsPlannedActivity, GanttDateLabelPosition.Start, true)]
    [InlineData(GanttEntityType.AsPlannedActivity, GanttDateLabelPosition.None, true)]
    [InlineData(GanttEntityType.AsPlannedMilestone, GanttDateLabelPosition.Finish, false)]
    [InlineData(GanttEntityType.AsPlannedMilestone, GanttDateLabelPosition.None, true)]
    [InlineData(GanttEntityType.Splitter, GanttDateLabelPosition.Auto, false)]
    [InlineData(GanttEntityType.Splitter, GanttDateLabelPosition.None, false)]
    [InlineData(GanttEntityType.Delineator, GanttDateLabelPosition.Auto, false)]
    [InlineData(GanttEntityType.Delineator, GanttDateLabelPosition.Start, false)]
    [InlineData(GanttEntityType.Delineator, GanttDateLabelPosition.None, false)]
    public void IsDateLabelPositionAllowed_refuses_cross_mode_positions(
        GanttEntityType type, GanttDateLabelPosition position, bool expected)
    {
        Assert.Equal(expected, EntityTypeCatalog.IsDateLabelPositionAllowed(type, position));
    }

    [Fact]
    public void TryParseDateLabelSelection_default_clears_to_null()
    {
        Assert.True(EntityTypeCatalog.TryParseDateLabelSelection("Default", out GanttDateLabelPosition? position));
        Assert.Null(position);
    }

    [Theory]
    [InlineData("Auto", GanttDateLabelPosition.Auto)]
    [InlineData("Start", GanttDateLabelPosition.Start)]
    [InlineData("Finish", GanttDateLabelPosition.Finish)]
    [InlineData("None", GanttDateLabelPosition.None)]
    public void TryParseDateLabelSelection_parses_defined_choices(
        string selectionId, GanttDateLabelPosition expected)
    {
        Assert.True(EntityTypeCatalog.TryParseDateLabelSelection(selectionId, out GanttDateLabelPosition? position));
        Assert.Equal(expected, position);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Widget")]
    // "Both" is a valid LabelPosition but NOT a valid DateLabelPosition,
    // so this pins that the two enumerations are distinct authorities.
    [InlineData("Both")]
    public void TryParseDateLabelSelection_refuses_unknown_text(string? selectionId)
    {
        Assert.False(EntityTypeCatalog.TryParseDateLabelSelection(selectionId, out GanttDateLabelPosition? position));
        Assert.Null(position);
    }
}
