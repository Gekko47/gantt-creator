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
        "Delineator",
    ];

    private static readonly GanttLabelPosition[] s_spanLabels =
        [GanttLabelPosition.Auto, GanttLabelPosition.Left, GanttLabelPosition.Right, GanttLabelPosition.Inside, GanttLabelPosition.Above, GanttLabelPosition.Below, GanttLabelPosition.None];

    private static readonly GanttLabelPosition[] s_milestoneLabels =
        [GanttLabelPosition.Auto, GanttLabelPosition.Left, GanttLabelPosition.Right, GanttLabelPosition.Above, GanttLabelPosition.Below, GanttLabelPosition.None];

    private static readonly GanttLabelPosition[] s_delineatorLabels =
        [GanttLabelPosition.Auto, GanttLabelPosition.TopLeft, GanttLabelPosition.TopRight, GanttLabelPosition.BottomLeft, GanttLabelPosition.BottomRight, GanttLabelPosition.None];

    [Fact]
    public void Default_catalogue_contains_exactly_the_16_guide_entries()
    {
        Assert.Equal(16, EntityTypeCatalog.Entries.Count);
        Assert.Equal(EntityTypeCatalog.TypeCount, EntityTypeCatalog.Entries.Count);
    }

    [Fact]
    public void Default_catalogue_matches_the_entity_guide_display_names_in_order()
    {
        Assert.Equal(16, s_guideDisplayNames.Length);
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
            (GanttEntityType.CriticalInterval, "Critical Interval", EntityKind.Span, EntityDateMode.StartFinish, "CriticalInterval", EntityColourCapability.Stroke, false, [GanttLabelPosition.None]),
            (GanttEntityType.DelayEvent, "Delay Event", EntityKind.Span, EntityDateMode.StartFinish, "DelayEvent", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.AsBuiltProcurement, "As-Built Procurement", EntityKind.Span, EntityDateMode.StartFinish, "AsBuiltProcurement", EntityColourCapability.Hatch | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.AsPlannedProcurement, "As-Planned Procurement", EntityKind.Span, EntityDateMode.StartFinish, "AsPlannedProcurement", EntityColourCapability.Hatch | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.BaselineProcurement, "Baseline Procurement", EntityKind.Span, EntityDateMode.StartFinish, "BaselineProcurement", EntityColourCapability.Hatch | EntityColourCapability.Stroke, false, s_spanLabels),
            (GanttEntityType.CustomActivity, "Custom Activity", EntityKind.Span, EntityDateMode.StartFinish, string.Empty, EntityColourCapability.StyleDefined, true, Array.Empty<GanttLabelPosition>()),
            (GanttEntityType.AsBuiltMilestone, "As-Built Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "AsBuiltMilestone", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_milestoneLabels),
            (GanttEntityType.AsPlannedMilestone, "As-Planned Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "AsPlannedMilestone", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_milestoneLabels),
            (GanttEntityType.BaselineMilestone, "Baseline Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "BaselineMilestone", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_milestoneLabels),
            (GanttEntityType.CriticalMilestone, "Critical Milestone", EntityKind.Milestone, EntityDateMode.StartOnly, "CriticalMilestone", EntityColourCapability.Fill | EntityColourCapability.Stroke, false, s_milestoneLabels),
            (GanttEntityType.Delineator, "Delineator", EntityKind.Delineator, EntityDateMode.StartOnly, "DefaultDelineator", EntityColourCapability.Stroke, false, s_delineatorLabels),
        };

        Assert.Equal(16, expected.Length);

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
    public void Only_CustomActivity_requires_a_style_key()
    {
        Assert.All(EntityTypeCatalog.Entries, entry =>
            Assert.Equal(entry.Type == GanttEntityType.CustomActivity, entry.RequiresStyleKey));
    }

    [Fact]
    public void CustomActivity_has_no_catalogue_default_style_and_no_label_set()
    {
        var custom = Assert.Single(EntityTypeCatalog.Entries, e => e.Type == GanttEntityType.CustomActivity);

        Assert.Equal(string.Empty, custom.DefaultStyleKey);
        Assert.True(custom.RequiresStyleKey);
        Assert.Empty(custom.AllowedLabelPositions);
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
                    or GanttEntityType.Delineator => EntityDateMode.StartOnly,
                GanttEntityType.Splitter or GanttEntityType.Spacer => EntityDateMode.None,
                _ => EntityDateMode.StartFinish,
            };
            Assert.Equal(expected, entry.DateMode);
        });
    }

    [Fact]
    public void Critical_interval_has_no_label_positions()
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
}
