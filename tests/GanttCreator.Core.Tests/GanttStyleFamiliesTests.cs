using System;
using System.Collections.Generic;
using System.Linq;
using GanttCreator.Core;
using Xunit;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for the <see cref="GanttStyleFamilies"/> authority: the map that
/// decides which family heading a built-in style appears under.
/// </summary>
/// <remarks>
/// The "no duplicated heading" property of the Entity styles list used to be a
/// property of the ORDER the dialog happened to render the catalogue in, so a
/// style landing next to a different family produced a second copy of that
/// family's heading. It is now a property of the data: every preset maps to
/// exactly one family, and the family list is fixed.
/// </remarks>
public class GanttStyleFamiliesTests
{
    [Fact]
    public void EveryBuiltInPresetMapsToExactlyOneFamily()
    {
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            GanttEntityFamily family = GanttStyleFamilies.For(preset.StyleKey);

            Assert.True(
                Enum.IsDefined(family),
                $"Style '{preset.StyleKey}' mapped to an undefined family.");
        }
    }

    [Fact]
    public void EveryStyleInAFamilyIsListedUnderThatFamilyOnly()
    {
        // The same style key must not be reachable through two families, and a
        // family must not be produced by two different keys with different
        // headings.
        Dictionary<GanttEntityFamily, List<string>> byFamily =
            GanttCatalogues.StylePresets
                .GroupBy(preset => GanttStyleFamilies.For(preset.StyleKey))
                .ToDictionary(
                    group => group.Key,
                    group => group.Select(preset => preset.StyleKey).ToList());

        foreach ((GanttEntityFamily family, List<string> keys) in byFamily)
        {
            Assert.Equal(keys.Count, keys.Distinct(StringComparer.Ordinal).Count());
            Assert.All(
                keys,
                key => Assert.Equal(family, GanttStyleFamilies.For(key)));
        }
    }

    [Fact]
    public void AllFourFamiliesArePresent()
    {
        HashSet<GanttEntityFamily> present = GanttCatalogues.StylePresets
            .Select(preset => GanttStyleFamilies.For(preset.StyleKey))
            .ToHashSet();

        Assert.Equal(
            GanttStyleFamilies.OrderedFamilies.Count,
            present.Count);

        foreach (GanttEntityFamily family in GanttStyleFamilies.OrderedFamilies)
        {
            Assert.Contains(family, present);
        }
    }

    [Fact]
    public void TheCustomTypesShareTheirFamiliesWithTheirRelatives()
    {
        // The two custom types used to be filed under a bucket of their own, so
        // the family list showed a heading that named no other style.
        Assert.Equal(
            GanttEntityFamily.Activities,
            GanttStyleFamilies.For("CustomActivity"));

        Assert.Equal(
            GanttEntityFamily.Milestones,
            GanttStyleFamilies.For("CustomMilestone"));
    }

    [Fact]
    public void EveryFamilyHasADistinctHeading()
    {
        string[] headings = GanttStyleFamilies.OrderedFamilies
            .Select(GanttStyleFamilies.DisplayNameFor)
            .ToArray();

        Assert.Equal(headings.Length, headings.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void AnUnknownStyleKeyIsRefused()
    {
        // The positive test for the guard: an unmapped key throws rather than
        // defaulting to a family, because a silent fallback is a second
        // authority that disagrees with the map at some margin.
        Assert.Throws<ArgumentOutOfRangeException>(
            () => GanttStyleFamilies.For("NotAPreset"));

        Assert.Throws<ArgumentOutOfRangeException>(
            () => GanttStyleFamilies.For("CustomUserStyle"));
    }

    [Fact]
    public void TheUserAuthoredStylesAreExactlyTheStyleKeyTypesFallbacks()
    {
        // A style-key type takes the style the USER authored, so its built-in
        // fallback is a placeholder rather than a style the user chose. The
        // Entity styles list does not offer it; the Custom styles tab is where
        // those styles live.
        HashSet<string> expected = EntityTypeCatalog.Entries
            .Where(entry => entry.RequiresStyleKey)
            .Select(entry => entry.DefaultStyleKey)
            .Where(key => key.Length > 0)
            .ToHashSet(StringComparer.Ordinal);

        Assert.NotEmpty(expected);
        Assert.True(GanttUserAuthoredStyles.IsUserAuthored("CustomActivity"));
        Assert.True(GanttUserAuthoredStyles.IsUserAuthored("CustomMilestone"));

        // Every code-owned style the user did not author is not in the set.
        Assert.False(GanttUserAuthoredStyles.IsUserAuthored("AsPlannedActivity"));
        Assert.False(GanttUserAuthoredStyles.IsUserAuthored("AsPlannedMilestone"));
        Assert.True(GanttUserAuthoredStyles.IsUserAuthored("DefaultDelineator"));

        // The set is derived from the type catalogue, so it cannot drift from
        // the types that require a style key.
        foreach (string key in expected)
        {
            Assert.True(GanttUserAuthoredStyles.IsUserAuthored(key), key);
        }
    }

    [Fact]
    public void TheEditableMetricSetIsExactlyTheLiveChartWideKnobs()
    {
        // The bar height and the outline width are resolved from the STYLE
        // row, not from these metrics, and the hatch pitch and line width
        // cannot be expressed by the host's pattern at all. Offering them would
        // promise a change the chart never renders.
        string[] expected =
        {
            "MilestoneSizePt",
            "GanttRowHeightPt", "SplitterHeightPt", "SpacerHeightPt",
            "GridLinePt", "MajorBoundaryPt", "DelineatorLinePt",
        };

        Assert.Equal(
            expected.OrderBy(name => name, StringComparer.Ordinal),
            ThemeEditRequest.EditableMetrics.OrderBy(name => name, StringComparer.Ordinal));

        foreach (string inert in new[] { "ActivityHeightPt", "StandardOutlinePt", "HatchPitchPt", "HatchLinePt" })
        {
            Assert.DoesNotContain(inert, ThemeEditRequest.EditableMetrics);
        }
    }
}
