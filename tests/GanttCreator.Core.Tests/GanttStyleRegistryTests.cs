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
        var registry = new GanttStyleRegistry([
            new GanttStyleDefinition(
                "AsBuiltActivity",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                GanttLabelPosition.Inside,
                FillColour: "#00B0F0",
                StrokeColour: "#0070C0",
                HatchPattern: GanttHatchPattern.None)]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.AsPlannedActivity,
            registry,
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
        var registry = new GanttStyleRegistry([
            new GanttStyleDefinition(
                "AsPlannedActivity",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                GanttLabelPosition.Inside,
                FillColour: "#92D050",
                StrokeColour: "#548235",
                HatchPattern: GanttHatchPattern.None)]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.AsPlannedActivity,
            registry,
            "MissingStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("AsPlannedActivity", resolved!.StyleKey);
        Assert.Equal("#92D050", resolved.FillColour!.ToString());
        Assert.Equal("#548235", resolved.StrokeColour!.ToString());
        Assert.True(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_resolves_custom_activity_via_builtin_preset_when_style_missing()
    {
        // A Custom Activity row whose user-chosen custom style was deleted from
        // the registry resolves through the code-owned CustomActivity preset
        // instead of refusing with StyleUnavailable.
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            GanttStyleRegistry.Empty,
            "DeletedCustomStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("CustomActivity", resolved!.StyleKey);
        Assert.True(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_resolves_custom_activity_via_builtin_preset_when_style_blank()
    {
        // A Custom Activity row that never chose a style key resolves through
        // the code-owned CustomActivity preset instead of refusing with
        // NoDefaultStyle.
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            GanttStyleRegistry.Empty,
            null,
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("CustomActivity", resolved!.StyleKey);
        Assert.False(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_uses_registry_entry_for_type_default()
    {
        // A blank style key resolves to the type's DefaultStyleKey entry in the
        // registry, using the registry's themed colours rather than code-owned
        // preset defaults.
        var registry = new GanttStyleRegistry([
            new GanttStyleDefinition(
                "AsBuiltActivity",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                GanttLabelPosition.Inside,
                FillColour: "#00B0F0",
                StrokeColour: "#0070C0",
                HatchPattern: GanttHatchPattern.None)]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.AsBuiltActivity,
            registry,
            null,
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("AsBuiltActivity", resolved!.StyleKey);
        Assert.Equal("#00B0F0", resolved.FillColour!.ToString());
        Assert.Equal("#0070C0", resolved.StrokeColour!.ToString());
        Assert.False(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_carries_the_named_style_text_colour()
    {
        // Entity guide section 17 is a statement about this value surviving
        // resolution: without it a delay label has no colour to be white.
        var registry = new GanttStyleRegistry([
            new GanttStyleDefinition(
                "DelayEvent",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                GanttLabelPosition.Inside,
                FillColour: "#FF0000",
                StrokeColour: "#C00000",
                TextColour: "#FFFFFF",
                HatchPattern: GanttHatchPattern.None)]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.DelayEvent,
            registry,
            "DelayEvent",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));


        Assert.Equal("#FFFFFF", resolved!.TextColour!.ToString());
    }

    [Fact]
    public void Resolver_leaves_the_text_colour_null_when_the_style_declares_none()
    {
        // Null means "the style resolved no text colour", which the adapter reads as
        // "leave the font alone". Substituting DefaultText here would be exactly the
        // renderer defaulting the architecture forbids, and would make an
        // intentionally unstyled label indistinguishable from a styled one.
        var registry = new GanttStyleRegistry([TextlessStyle()]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            registry,
            "TextlessStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));


        Assert.Null(resolved!.TextColour);
    }

    [Fact]
    public void A_malformed_text_colour_is_refused_at_construction_not_merely_ignored()
    {
        // Where the text colour is actually rejected. Both GanttStyleDefinition and
        // GanttStylePreset validate their colours when they are built, so a bad
        // value never reaches GanttStyleResolver to be silently dropped - which
        // would have reached the host as "no text colour" and read as "leave the
        // font alone". The resolver's own colour parse is defence-in-depth behind
        // this check, symmetric with the fill and stroke guards.
        var definition = new GanttStyleDefinition(
            "BadTextStyle",
            new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
            EntityColourCapability.Fill,
            GanttLabelPosition.Inside,
            FillColour: "#112233",
            StrokeColour: "#445566",
            TextColour: "not-a-colour",
            HatchPattern: GanttHatchPattern.None);

        Assert.Throws<ArgumentException>(() => new GanttStyleRegistry([definition]));
    }

    [Fact]
    public void A_well_formed_text_colour_is_accepted_so_the_refusal_is_not_vacuous()
    {
        // The control for the row above: the identical definition with a valid
        // colour must construct, so the refusal above cannot be passing for
        // "custom styles never build".
        var registry = new GanttStyleRegistry([StyleWithTextColour("#123456")]);
        Assert.True(registry.TryGet("TextlessStyle", out GanttStyleDefinition? found));
        Assert.Equal("#123456", found!.TextColour);
    }

    [Fact]
    public void A_registry_text_colour_reaches_the_resolved_style_verbatim()
    {
        // The reachable end-to-end path: a user-authored style's text colour is
        // carried into the resolved style rather than dropped. Section 17 depends on
        // this surviving, and it is the only assertion here that exercises the
        // resolver's new TextColour member with a value it must preserve.
        var registry = new GanttStyleRegistry([StyleWithTextColour("#123456")]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            registry,
            "TextlessStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));


        Assert.Equal("#123456", resolved!.TextColour!.ToString());
    }

    /// <summary>
    /// A custom style declaring no text colour at all, so the "unresolved stays
    /// null" row is testing an absent value rather than a malformed one.
    /// </summary>
    private static GanttStyleDefinition TextlessStyle() => StyleWithTextColour(string.Empty);

    private static GanttStyleDefinition StyleWithTextColour(string textColour) =>
        new(
            "TextlessStyle",
            new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
            EntityColourCapability.Fill,
            GanttLabelPosition.Inside,
            FillColour: "#112233",
            StrokeColour: "#445566",
            TextColour: textColour,
            // A named style with a default label position is a formatting style, and
            // ValidateFormatting requires every formatting member to be present.
            HatchPattern: GanttHatchPattern.None);

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
            out GanttStyleResolutionRefusal? refusal));


        Assert.Equal("#ABCDEF", resolved!.FillColour!.ToString());
        Assert.Equal("#445566", resolved.StrokeColour!.ToString());
        Assert.Equal(GanttLabelPosition.None, resolved.LabelPosition);
    }

    [Fact]
    public void Resolver_refuses_unknown_style_key_even_when_type_has_preset()
    {
        // The registry is the single source of truth for a BUILT-IN type: an
        // unknown key is refused even though the type has a code-owned built-in
        // preset. The code-owned preset fallback is scoped to the two
        // style-key types (Custom Activity, Custom Milestone), whose style is
        // user-supplied by design; for every other type a missing style means
        // the configuration sheet is damaged, and silently rendering as a
        // preset would hide that (R2.7c).
        var registry = new GanttStyleRegistry([Formatted()]);

        Assert.False(GanttStyleResolver.TryResolve(
            GanttEntityType.AsPlannedActivity,
            registry,
            "MissingStyle",
            null,
            null,
            null,
            out _,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Equal(GanttStyleResolutionRefusal.StyleUnavailable, refusal);
    }

    [Fact]
    public void Resolver_resolves_a_capability_only_registry_entry_for_custom_activity()
    {
        // A capability-only definition has no formatting values, but it is present
        // in the registry (read from the workbook), so it must resolve rather than
        // being refused. Colours fall back to empty string; DefaultLabelPosition
        // falls back to Auto.
        var registry = new GanttStyleRegistry([
            new GanttStyleDefinition(
                "CustomStyle",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
                EntityColourCapability.Fill)]);

        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            registry,
            "CustomStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));


        Assert.Equal("CustomStyle", resolved!.StyleKey);
        Assert.Null(resolved.FillColour);
        Assert.Null(resolved.StrokeColour);
        Assert.Null(resolved.TextColour);
        Assert.Equal(GanttLabelPosition.Auto, resolved.LabelPosition);
        Assert.False(resolved.UsedFallback);
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
            out GanttStyleResolutionRefusal? refusal));


        Assert.Equal("AsPlannedActivity", resolved!.StyleKey);
        Assert.Equal(GanttLabelPosition.BottomRight, resolved.LabelPosition);
        Assert.False(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_falls_back_to_builtin_custom_activity_preset_when_custom_style_is_deleted()
    {
        // Simulates a workbook whose tblGanttStyles no longer contains the
        // custom style that a Custom Activity row referenced. The row's explicit
        // key is gone, so resolution falls back to the code-owned CustomActivity
        // preset rather than refusing with StyleUnavailable.
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            GanttStyleRegistry.Empty,
            "DeletedCustomStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("CustomActivity", resolved!.StyleKey);
        Assert.True(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_falls_back_to_builtin_custom_milestone_preset_when_custom_style_is_deleted()
    {
        // Mirrors the Custom Activity case for Custom Milestone: a row whose
        // user-chosen custom style was deleted from the registry still renders
        // through the code-owned CustomMilestone preset.
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomMilestone,
            GanttStyleRegistry.Empty,
            "DeletedCustomStyle",
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("CustomMilestone", resolved!.StyleKey);
        Assert.True(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_falls_back_to_builtin_custom_activity_preset_when_style_blank()
    {
        // A Custom Activity row that never chose a style key still resolves
        // through the built-in CustomActivity preset instead of refusing
        // with NoDefaultStyle.
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomActivity,
            GanttStyleRegistry.Empty,
            null,
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("CustomActivity", resolved!.StyleKey);
        Assert.False(resolved.UsedFallback);
    }

    [Fact]
    public void Resolver_falls_back_to_builtin_custom_milestone_preset_when_style_blank()
    {
        // Mirrors the Custom Activity blank-key case for Custom Milestone.
        Assert.True(GanttStyleResolver.TryResolve(
            GanttEntityType.CustomMilestone,
            GanttStyleRegistry.Empty,
            null,
            null,
            null,
            null,
            out GanttResolvedStyle? resolved,
            out GanttStyleResolutionRefusal? refusal));

        Assert.Null(refusal);
        Assert.Equal("CustomMilestone", resolved!.StyleKey);
        Assert.False(resolved.UsedFallback);
    }
}

