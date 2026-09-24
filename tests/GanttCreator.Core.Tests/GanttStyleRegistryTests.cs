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
