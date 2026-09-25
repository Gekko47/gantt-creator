namespace GanttCreator.Core.Tests;

/// <summary>Positive tests for the immutable named-style capability registry.</summary>
public class GanttStyleRegistryTests
{
    [Fact]
    public void Registry_projects_capabilities_and_finds_exact_keys()
    {
        var style = new GanttStyleDefinition(
            "CustomStyle",
            new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
            EntityColourCapability.Fill | EntityColourCapability.Stroke);
        var registry = new GanttStyleRegistry([style]);

        Assert.Equal(1, registry.Count);
        Assert.True(registry.TryGet("CustomStyle", out GanttStyleDefinition? found));
        Assert.Same(style, found);
        Assert.False(registry.TryGet("customstyle", out _));
    }

    [Fact]
    public void Resolver_uses_selected_style_defaults_then_row_overrides()
    {
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.AsPlannedActivity,
            "AsBuiltActivity",
            "#112233",
            null,
            GanttLabelPosition.Inside,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("AsBuiltActivity", resolved!.StyleKey);
        Assert.Equal("#112233", resolved.FillColour!.ToString());
        Assert.Equal("#0070C0", resolved.StrokeColour!.ToString());
        Assert.Equal(GanttLabelPosition.Inside, resolved.LabelPosition);
        Assert.False(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_falls_back_to_the_type_default_for_an_unknown_style()
    {
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.AsPlannedActivity,
            "MissingStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("AsPlannedActivity", resolved!.StyleKey);
        Assert.True(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_refuses_a_type_without_a_default_style()
    {
        Assert.False(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            "MissingCustomStyle",
            null,
            null,
            null,
            out _,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Equal(GanttStyleResolutionRefusal.NoDefaultStyle, refusal);
    }

    [Fact]
    public void ChangeStyleKey_clears_all_per_row_formatting_overrides()
    {
        GanttRowDto row = new(
            2,
            "G-0123456789abcdef0123456789abcdef",
            null,
            0,
            "As-Planned Activity",
            null,
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 2),
            null,
            "AsPlannedActivity",
            "Inside",
            "#112233",
            "#445566",
            true,
            null);

        GanttRowDto changed = GanttStyleChange.ChangeStyleKey(row, "AsBuiltActivity");

        Assert.Equal("AsBuiltActivity", changed.StyleKey);
        Assert.Null(changed.LabelPositionText);
        Assert.Null(changed.FillColourText);
        Assert.Null(changed.StrokeColourText);
    }

    [Fact]
    public void Registry_rejects_null_input() =>
        Assert.Throws<ArgumentNullException>(() => new GanttStyleRegistry(null!));

    [Fact]
    public void Registry_rejects_blank_keys() =>
        Assert.Throws<ArgumentException>(() => new GanttStyleRegistry([
            new GanttStyleDefinition(" ", new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto }, EntityColourCapability.Fill)]));

    [Fact]
    public void Registry_rejects_duplicate_keys()
    {
        var first = new GanttStyleDefinition("Same", new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto }, EntityColourCapability.Fill);
        var second = new GanttStyleDefinition("Same", new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside }, EntityColourCapability.Stroke);

        Assert.Throws<ArgumentException>(() => new GanttStyleRegistry([first, second]));
    }

    [Fact]
    public void Registry_rejects_empty_or_undefined_label_capabilities()
    {
        Assert.Throws<ArgumentException>(() => new GanttStyleRegistry([
            new GanttStyleDefinition("Empty", new HashSet<GanttLabelPosition>(), EntityColourCapability.Fill)]));
        Assert.Throws<ArgumentException>(() => new GanttStyleRegistry([
            new GanttStyleDefinition("Undefined", new HashSet<GanttLabelPosition> { (GanttLabelPosition)999 }, EntityColourCapability.Fill)]));
    }

    [Fact]
    public void Registry_rejects_unknown_colour_capability_bits() =>
        Assert.Throws<ArgumentOutOfRangeException>(() => new GanttStyleRegistry([
            new GanttStyleDefinition("Invalid", new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto }, (EntityColourCapability)16)]));
}
