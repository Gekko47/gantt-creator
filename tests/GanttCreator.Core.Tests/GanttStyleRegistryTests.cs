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
    public void Resolver_refuses_a_supplied_style_key_that_cannot_resolve()
    {
        // A supplied-but-unresolvable key is "unavailable", not "no default": the
        // row asked for a style and none could be produced.
        Assert.False(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            "MissingCustomStyle",
            null,
            null,
            null,
            out _,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Equal(GanttStyleResolutionRefusal.StyleUnavailable, refusal);
    }

    [Fact]
    public void Resolver_refuses_a_blank_style_key_on_a_type_without_a_default()
    {
        // Custom Activity has no Type default, so a row that supplies no key has
        // nothing to fall back to.
        Assert.False(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            null,
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
    public void ChangeStyleKey_keeps_the_per_row_overrides_when_the_style_is_unchanged()
    {
        // The overrides are cleared only when the style actually changes. Picking
        // the row's current style in the dropdown again must not discard the
        // user's explicit per-row formatting.
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

        GanttRowDto unchanged = GanttStyleChange.ChangeStyleKey(row, "AsPlannedActivity");

        Assert.Equal("AsPlannedActivity", unchanged.StyleKey);
        Assert.Equal("Inside", unchanged.LabelPositionText);
        Assert.Equal("#112233", unchanged.FillColourText);
        Assert.Equal("#445566", unchanged.StrokeColourText);
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

/// <summary>
/// Positive tests for the R2.7c resolved-formatting payload on
/// <see cref="GanttStyleDefinition"/> and the registry-based resolver path.
/// </summary>
public sealed class GanttStyleFormattingTests
{
    private static GanttStyleDefinition Formatted(
        string styleKey = "UserStyle",
        string? fill = "#112233",
        string? stroke = "#445566",
        string? text = "#FFFFFF",
        GanttLabelPosition defaultLabel = GanttLabelPosition.Inside,
        GanttHatchPattern? hatch = GanttHatchPattern.None,
        double hatchPitchPt = 4,
        double hatchLinePt = 0.5,
        double outlinePt = 0.75,
        double heightPt = 8,
        double milestoneSizePt = 10) =>
        new(
            styleKey,
            new HashSet<GanttLabelPosition> { defaultLabel },
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            defaultLabel,
            fill,
            stroke,
            text,
            hatch,
            hatchPitchPt,
            hatchLinePt,
            outlinePt,
            heightPt,
            milestoneSizePt);

    [Fact]
    public void A_capability_only_definition_is_legal_and_reports_no_formatting()
    {
        // D1: the compatibility path must keep compiling and stay legal.
        var capabilityOnly = new GanttStyleDefinition(
            "CustomStyle",
            new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
            EntityColourCapability.Fill);

        var registry = new GanttStyleRegistry([capabilityOnly]);

        Assert.False(capabilityOnly.HasFormatting);
        Assert.True(registry.TryGet("CustomStyle", out GanttStyleDefinition? found));
        Assert.Same(capabilityOnly, found);
    }

    [Fact]
    public void A_formatted_definition_reports_has_formatting()
    {
        var registry = new GanttStyleRegistry([Formatted()]);

        Assert.True(registry.TryGet("UserStyle", out GanttStyleDefinition? found));
        Assert.True(found!.HasFormatting);
        Assert.Equal("#112233", found.FillColour);
        Assert.Equal("#445566", found.StrokeColour);
        Assert.Equal("#FFFFFF", found.TextColour);
        Assert.Equal(GanttLabelPosition.Inside, found.DefaultLabelPosition);
        Assert.Equal(8, found.ActivityHeightPt);
        Assert.Equal(10, found.MilestoneSizePt);
    }

    [Theory]
    [InlineData("#ff0000")]
    [InlineData("112233")]
    [InlineData("red")]
    public void Registry_rejects_a_lowercase_or_malformed_fill_colour(string fill) =>
        Assert.Throws<ArgumentException>(() => new GanttStyleRegistry([Formatted(fill: fill)]));

    [Fact]
    public void Registry_rejects_a_malformed_stroke_colour() =>
        Assert.Throws<ArgumentException>(() => new GanttStyleRegistry([Formatted(stroke: "nope")]));

    [Fact]
    public void Registry_rejects_a_malformed_text_colour() =>
        Assert.Throws<ArgumentException>(() => new GanttStyleRegistry([Formatted(text: "#GGGGGG")]));

    [Fact]
    public void Registry_rejects_a_negative_metric() =>
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new GanttStyleRegistry([Formatted(hatchPitchPt: -1)]));

    [Fact]
    public void Registry_rejects_a_non_finite_metric() =>
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new GanttStyleRegistry([Formatted(hatchPitchPt: double.NaN)]));

    [Fact]
    public void Registry_rejects_a_missing_hatch_pattern_when_formatting_is_present() =>
        Assert.Throws<ArgumentOutOfRangeException>(() => new GanttStyleRegistry([Formatted(hatch: null)]));

    [Fact]
    public void Registry_rejects_an_undefined_hatch_pattern() =>
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new GanttStyleRegistry([Formatted(hatch: (GanttHatchPattern)99)]));

    [Fact]
    public void Registry_rejects_an_undefined_default_label_position() =>
        Assert.Throws<ArgumentOutOfRangeException>(() => new GanttStyleRegistry([
            Formatted() with { DefaultLabelPosition = (GanttLabelPosition)99 }]));

    [Fact]
    public void Resolver_reads_a_user_style_from_the_registry()
    {
        // The gap this row closes: before R2.7c a user-authored style resolved to
        // nothing, because the resolver only read the built-in presets.
        var registry = new GanttStyleRegistry([Formatted()]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            registry,
            "UserStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("UserStyle", resolved!.StyleKey);
        Assert.Equal("#112233", resolved.FillColour!.ToString());
        Assert.Equal("#445566", resolved.StrokeColour!.ToString());
        Assert.Equal(GanttLabelPosition.Inside, resolved.LabelPosition);
        Assert.False(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_applies_row_overrides_over_a_user_style()
    {
        var registry = new GanttStyleRegistry([Formatted()]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            registry,
            "UserStyle",
            "#ABCDEF",
            null,
            GanttLabelPosition.None,
            out GanttResolvedStyle? resolved,
            out _));

        Assert.Equal("#ABCDEF", resolved!.FillColour!.ToString());
        Assert.Equal("#445566", resolved.StrokeColour!.ToString());
        Assert.Equal(GanttLabelPosition.None, resolved.LabelPosition);
    }

    [Fact]
    public void Resolver_falls_back_to_the_built_in_preset_for_an_unknown_key()
    {
        // D2: an unknown key still falls back to the Type default preset, so the
        // compatibility overload and the Type-default rule are unchanged.
        var registry = new GanttStyleRegistry([Formatted()]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.AsPlannedActivity,
            registry,
            "MissingStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out _));

        Assert.Equal("AsPlannedActivity", resolved!.StyleKey);
        Assert.True(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_refuses_a_capability_only_registry_entry_for_custom_activity()
    {
        // A capability-only definition has no formatting, so a row that selects it
        // must refuse as unavailable rather than silently resolve to a blank style.
        var registry = new GanttStyleRegistry([
            new GanttStyleDefinition(
                "CustomStyle",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
                EntityColourCapability.Fill)]);

        Assert.False(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            registry,
            "CustomStyle",
            null,
            null,
            null,
            out _,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Equal(GanttStyleResolutionRefusal.StyleUnavailable, refusal);
    }

    [Fact]
    public void A_user_style_resolves_identically_after_a_registry_rebuild()
    {
        // R2.7b's save/reopen gate, expressed for formatting: rebuilding the registry
        // from the same definitions must not change the resolved values.
        var registry = new GanttStyleRegistry([Formatted()]);
        var rebuilt = new GanttStyleRegistry([Formatted()]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            registry,
            "UserStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? first,
            out _));
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            rebuilt,
            "UserStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? second,
            out _));

        Assert.Equal(first!.StyleKey, second!.StyleKey);
        Assert.Equal(first.FillColour, second.FillColour);
        Assert.Equal(first.StrokeColour, second.StrokeColour);
        Assert.Equal(first.LabelPosition, second.LabelPosition);
    }

    [Fact]
    public void Resolver_uses_the_registry_type_default_when_the_row_leaves_style_blank()
    {
        var registry = new GanttStyleRegistry([
            Formatted(styleKey: "AsPlannedActivity", defaultLabel: GanttLabelPosition.BottomRight)]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.AsPlannedActivity,
            registry,
            null,
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out _));

        Assert.Equal("AsPlannedActivity", resolved!.StyleKey);
        Assert.Equal(GanttLabelPosition.BottomRight, resolved.LabelPosition);
        Assert.False(resolved.UsedFallback);
    }
}

