using System.Globalization;
using GanttCreator.Core;
using Xunit;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Contract tests for the code-owned configuration catalogue (work item
/// R2.7, ADR-0007, and ADR-0014). The count pins (23/21/6/16/15/12) and the pinned
/// catalogue hash detect any entity-guide drift: the guide remains
/// authoritative, so a changed token is a guide change first and then a
/// deliberate test update in the same commit. Every constructor guard has a
/// positive test that constructs the bad input and asserts the exact
/// exception (AGENTS.md validator rule).
/// </summary>
public class GanttConfigCatalogueTests
{
    // ------------------------------------------------------------------
    // Count pins (entity-guide revision 3, 2026-09-21)
    // ------------------------------------------------------------------

    [Fact]
    public void Metrics_contains_exactly_the_23_entity_guide_tokens() =>
        Assert.Equal(23, GanttCatalogues.Metrics.Count);

    [Fact]
    public void Colours_contains_exactly_the_21_entity_guide_tokens() =>
        Assert.Equal(21, GanttCatalogues.Colours.Count);

    [Fact]
    public void Typography_contains_exactly_the_6_entity_guide_tokens() =>
        Assert.Equal(6, GanttCatalogues.Typography.Count);

    [Fact]
    public void TypeRows_contains_exactly_the_16_catalogue_entries() =>
        Assert.Equal(16, GanttCatalogues.TypeRows.Count);

    [Fact]
    public void StylePresets_contains_one_row_per_non_empty_DefaultStyleKey()
    {
        var expected = EntityTypeCatalog.Entries
            .Where(entry => entry.DefaultStyleKey.Length > 0)
            .Select(entry => entry.DefaultStyleKey)
            .ToArray();
        Assert.Equal(15, expected.Length);
        Assert.Equal(
            expected,
            GanttCatalogues.StylePresets.Select(preset => preset.StyleKey).ToArray());
    }

    [Fact]
    public void Settings_contains_exactly_the_12_approved_keys()
    {
        Assert.Equal(12, GanttCatalogues.Settings.Count);
        Assert.Equal(
            SettingsKeys,
            GanttCatalogues.Settings.Select(setting => setting.Key).ToArray());
    }

    /// <summary>The approved first-release setting keys, in contract order.</summary>
    private static readonly string[] SettingsKeys =
    [
        "ChartTitle",
        "ShowTitle",
        "TimeScale",
        "PeriodLabelFormat",
        "LegendPosition",
        "ExportIncludeDataPanel",
        "ExportIncludeLegend",
        "PlotStartMode",
        "PlotFinishMode",
        "AlternateBanding",
        "ShowMinorGrid",
        "ShowMajorGrid",
    ];

    [Fact]
    public void Chart_settings_have_the_approved_defaults_and_compatibility_matrix()
    {
        Assert.Equal("Gantt Chart", GanttCatalogues.Settings.Single(setting => setting.Key == "ChartTitle").DefaultValue);
        Assert.Equal("MMM", GanttCatalogues.Settings.Single(setting => setting.Key == "PeriodLabelFormat").DefaultValue);

        Assert.True(GanttChartSettings.TryParseTimeScale("Month", out GanttTimeScale month));
        Assert.True(GanttChartSettings.TryParsePeriodLabelFormat("MMM", out GanttPeriodLabelFormat mmm));
        Assert.True(GanttChartSettings.IsCompatible(month, mmm));
        Assert.True(GanttChartSettings.TryParseTimeScale("Quarter", out GanttTimeScale quarter));
        Assert.True(GanttChartSettings.TryParsePeriodLabelFormat("Quarter", out GanttPeriodLabelFormat quarterFormat));
        Assert.True(GanttChartSettings.IsCompatible(quarter, quarterFormat));
        Assert.True(GanttChartSettings.TryParseTimeScale("Year", out GanttTimeScale year));
        Assert.True(GanttChartSettings.TryParsePeriodLabelFormat("Year", out GanttPeriodLabelFormat yearFormat));
        Assert.True(GanttChartSettings.IsCompatible(year, yearFormat));
    }

    [Theory]
    [InlineData("Day")]
    [InlineData("month")]
    [InlineData("")]
    public void Chart_settings_reject_unknown_time_scales(string text)
    {
        Assert.False(GanttChartSettings.TryParseTimeScale(text, out _));
    }

    [Theory]
    [InlineData("MMM ")]
    [InlineData("mMM")]
    [InlineData("")]
    public void Chart_settings_reject_unknown_period_formats(string text)
    {
        Assert.False(GanttChartSettings.TryParsePeriodLabelFormat(text, out _));
    }

    [Fact]
    public void Chart_settings_reject_incompatible_scale_and_format_pairs()
    {
        Assert.False(GanttChartSettings.IsCompatible(GanttTimeScale.Month, GanttPeriodLabelFormat.Quarter));
        Assert.False(GanttChartSettings.IsCompatible(GanttTimeScale.Quarter, GanttPeriodLabelFormat.MM));
        Assert.False(GanttChartSettings.IsCompatible(GanttTimeScale.Year, GanttPeriodLabelFormat.MMM));
    }

    // ------------------------------------------------------------------
    // Range and format pins
    // ------------------------------------------------------------------

    [Fact]
    public void Every_metric_default_lies_within_its_range_and_the_range_is_ordered()
    {
        foreach (GanttMetricToken metric in GanttCatalogues.Metrics)
        {
            Assert.True(metric.Minimum <= metric.Maximum, metric.Name);
            Assert.InRange(metric.DefaultValue, metric.Minimum, metric.Maximum);
        }
    }

    [Fact]
    public void Every_colour_value_is_uppercase_rrggbb()
    {
        foreach (GanttColourToken colour in GanttCatalogues.Colours)
        {
            Assert.True(GanttColourToken.IsValidHex(colour.HexValue), colour.Name);
        }
    }

    [Fact]
    public void Token_names_are_unique_within_each_catalogue()
    {
        Assert.Equal(
            GanttCatalogues.Metrics.Count,
            GanttCatalogues.Metrics.Select(metric => metric.Name).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(
            GanttCatalogues.Colours.Count,
            GanttCatalogues.Colours.Select(colour => colour.Name).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(
            GanttCatalogues.Typography.Count,
            GanttCatalogues.Typography.Select(token => token.Name).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void TypeRows_pin_each_DisplayName_and_DefaultStyleKey_pair_against_the_catalogue()
    {
        Assert.Equal(
            EntityTypeCatalog.Entries.Select(entry => (entry.DisplayName, entry.DefaultStyleKey)).ToArray(),
            GanttCatalogues.TypeRows.Select(row => (row.DisplayName, row.DefaultStyleKey)).ToArray());
    }

    [Fact]
    public void TypeRows_project_the_catalogue_contract_metadata_exactly()
    {
        for (var index = 0; index < EntityTypeCatalog.Entries.Count; index++)
        {
            EntityTypeDefinition entry = EntityTypeCatalog.Entries[index];
            GanttTypeCatalogueRow row = GanttCatalogues.TypeRows[index];
            Assert.Equal(entry.Type.ToString(), row.TypeName);
            Assert.Equal(entry.Kind.ToString(), row.Kind);
            Assert.Equal(entry.DateMode.ToString(), row.DateMode);
            Assert.Equal(entry.RequiresStyleKey ? "TRUE" : "FALSE", row.RequiresStyleKey);
        }
    }

    [Fact]
    public void Representative_style_presets_resolve_the_entity_guide_token_defaults()
    {
        // As-built activity: ActualFill / ActualOutline (guide §16 default style).
        GanttStylePreset asBuilt = GanttCatalogues.GetPreset("AsBuiltActivity");
        Assert.Equal("#00B0F0", asBuilt.FillColour);
        Assert.Equal("#0070C0", asBuilt.StrokeColour);
        Assert.Equal(GanttHatchPattern.None, asBuilt.HatchPattern);
        Assert.Equal(8, asBuilt.ActivityHeightPt);
        Assert.Equal(0.75, asBuilt.StandardOutlinePt);
        Assert.Equal(0, asBuilt.MilestoneSizePt);

        // Delay event: DelayFill / CriticalOutline with DelayText.
        GanttStylePreset delay = GanttCatalogues.GetPreset("DelayEvent");
        Assert.Equal("#FF0000", delay.FillColour);
        Assert.Equal("#C00000", delay.StrokeColour);
        Assert.Equal("#FFFFFF", delay.TextColour);

        // Procurement: forward diagonal hatch on the planned colours.
        GanttStylePreset procurement = GanttCatalogues.GetPreset("AsPlannedProcurement");
        Assert.Equal("#92D050", procurement.FillColour);
        Assert.Equal(GanttHatchPattern.ForwardDiagonal, procurement.HatchPattern);
        Assert.Equal(4, procurement.HatchPitchPt);
        Assert.Equal(0.5, procurement.HatchLinePt);

        // Critical milestone: CriticalStroke / CriticalOutline (guide §21).
        GanttStylePreset criticalMilestone = GanttCatalogues.GetPreset("CriticalMilestone");
        Assert.Equal("#FF0000", criticalMilestone.FillColour);
        Assert.Equal("#C00000", criticalMilestone.StrokeColour);
        Assert.Equal(8, criticalMilestone.MilestoneSizePt);

        // Delineator: stroke only, full height (no height token).
        GanttStylePreset delineator = GanttCatalogues.GetPreset("DefaultDelineator");
        Assert.Equal(string.Empty, delineator.FillColour);
        Assert.Equal("#404040", delineator.StrokeColour);
        Assert.Equal(0.75, delineator.StandardOutlinePt);
        Assert.Equal(0, delineator.ActivityHeightPt);

        // Splitter: fill band only.
        GanttStylePreset splitter = GanttCatalogues.GetPreset("Splitter");
        Assert.Equal("#FFE699", splitter.FillColour);
        Assert.Equal(string.Empty, splitter.StrokeColour);
        Assert.Equal(18, splitter.ActivityHeightPt);
    }
    [Fact]
    public void Style_presets_project_label_and_colour_capabilities_from_the_type_catalogue()
    {
        GanttStylePreset splitter = GanttCatalogues.GetPreset("Splitter");
        Assert.Equal(GanttLabelPosition.DataPanelLeft, splitter.DefaultLabelPosition);
        Assert.Equal(EntityColourCapability.Fill, splitter.ColourCapability);

        GanttStylePreset spacer = GanttCatalogues.GetPreset("Spacer");
        Assert.Equal(GanttLabelPosition.None, spacer.DefaultLabelPosition);
        Assert.Equal(EntityColourCapability.None, spacer.ColourCapability);

        GanttStylePreset critical = GanttCatalogues.GetPreset("CriticalInterval");
        Assert.Equal(GanttLabelPosition.None, critical.DefaultLabelPosition);
        Assert.Equal(EntityColourCapability.Stroke, critical.ColourCapability);

        GanttStylePreset delay = GanttCatalogues.GetPreset("DelayEvent");
        Assert.Equal(GanttLabelPosition.Inside, delay.DefaultLabelPosition);

        GanttStylePreset procurement = GanttCatalogues.GetPreset("AsPlannedProcurement");
        Assert.Equal(
            EntityColourCapability.Hatch | EntityColourCapability.Stroke,
            procurement.ColourCapability);
        Assert.Contains(GanttLabelPosition.Inside, procurement.AllowedLabelPositions);

        GanttStylePreset milestone = GanttCatalogues.GetPreset("AsBuiltMilestone");
        Assert.Equal(GanttLabelPosition.Auto, milestone.DefaultLabelPosition);
        Assert.DoesNotContain(GanttLabelPosition.Inside, milestone.AllowedLabelPositions);
    }



    // ------------------------------------------------------------------
    // Catalogue hash (ADR-0007 D6)
    // ------------------------------------------------------------------

    [Fact]
    public void The_catalogue_hash_is_64_lowercase_hex_characters()
    {
        var hash = GanttCatalogues.ComputeCatalogueHash();
        Assert.Equal(64, hash.Length);
        Assert.All(hash.ToCharArray(), c => Assert.True(
            c is (>= '0' and <= '9') or (>= 'a' and <= 'f'),
            $"Hash character '{c}' is not lowercase hex."));
    }

    [Fact]
    public void The_catalogue_hash_is_deterministic_across_calls() =>
        Assert.Equal(GanttCatalogues.ComputeCatalogueHash(), GanttCatalogues.ComputeCatalogueHash());

    [Fact]
    public void The_catalogue_hash_is_stable_across_host_cultures()
    {
        var reference = GanttCatalogues.ComputeCatalogueHash();
        foreach (var cultureName in new[] { "tr-TR", "de-DE" })
        {
            var culture = new CultureInfo(cultureName);
            CultureInfo original = CultureInfo.CurrentCulture;
            CultureInfo originalUi = CultureInfo.CurrentUICulture;
            try
            {
                CultureInfo.CurrentCulture = culture;
                CultureInfo.CurrentUICulture = culture;
                Assert.Equal(reference, GanttCatalogues.ComputeCatalogueHash());
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
                CultureInfo.CurrentUICulture = originalUi;
            }
        }
    }

    // ------------------------------------------------------------------
    // Guard positive tests (AGENTS.md validator rule)
    // ------------------------------------------------------------------

    [Fact]
    public void Metric_token_rejects_an_empty_name() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttMetricToken(string.Empty, 1, 0, 2));

    [Fact]
    public void Metric_token_rejects_a_null_name() =>
        Assert.Throws<ArgumentNullException>(
            () => _ = new GanttMetricToken(null!, 1, 0, 2));

    [Fact]
    public void Metric_token_rejects_a_non_finite_default() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttMetricToken("Bad", double.NaN, 0, 2));

    [Fact]
    public void Metric_token_rejects_an_infinite_minimum() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttMetricToken("Bad", 1, double.PositiveInfinity, 2));

    [Fact]
    public void Metric_token_rejects_an_inverted_range() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttMetricToken("Bad", 1, 3, 2));

    [Fact]
    public void Metric_token_rejects_a_default_outside_the_range() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttMetricToken("Bad", 9, 0, 2));

    [Fact]
    public void Colour_token_rejects_a_lower_case_hex_value() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttColourToken("Bad", "#00b0f0"));

    [Fact]
    public void Colour_token_rejects_a_hex_value_without_the_hash_prefix() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttColourToken("Bad", "00B0F0"));

    [Fact]
    public void Colour_token_rejects_a_non_hex_value() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttColourToken("Bad", "#GGGGGG"));

    [Fact]
    public void Colour_token_rejects_a_null_value() =>
        Assert.Throws<ArgumentNullException>(
            () => _ = new GanttColourToken("Bad", null!));

    [Fact]
    public void IsValidHex_accepts_only_exact_uppercase_rrggbb()
    {
        Assert.True(GanttColourToken.IsValidHex("#00B0F0"));
        Assert.False(GanttColourToken.IsValidHex(null!));
        Assert.False(GanttColourToken.IsValidHex(string.Empty));
        Assert.False(GanttColourToken.IsValidHex("#00B0F"));
        Assert.False(GanttColourToken.IsValidHex("#00B0F00"));
        Assert.False(GanttColourToken.IsValidHex("00B0F0"));
        Assert.False(GanttColourToken.IsValidHex("#00b0f0"));
    }

    [Fact]
    public void Typography_token_rejects_an_empty_value() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttTypographyToken("FontFamily", string.Empty, isBold: false));

    [Fact]
    public void Typography_token_rejects_a_null_name() =>
        Assert.Throws<ArgumentNullException>(
            () => _ = new GanttTypographyToken(null!, "Aptos", isBold: false));

    [Fact]
    public void Setting_definition_rejects_an_empty_key() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttSettingDefinition(string.Empty, "TRUE"));

    [Fact]
    public void Setting_definition_rejects_a_null_value() =>
        Assert.Throws<ArgumentNullException>(
            () => _ = new GanttSettingDefinition("Key", null!));

    [Fact]
    public void Style_preset_rejects_a_bad_fill_colour() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttStylePreset("Bad", "Bad", "red", string.Empty,
                GanttHatchPattern.None, 4, 0.5, "#000000", 0.75, 8, 0,
                GanttLabelPosition.Auto,
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto },
                EntityColourCapability.Fill));

    [Fact]
    public void Style_preset_rejects_a_negative_metric() =>
        Assert.Throws<ArgumentException>(
            () => _ = new GanttStylePreset("Bad", "Bad", string.Empty, string.Empty,
                GanttHatchPattern.None, -1, 0.5, "#000000", 0.75, 8, 0,
                GanttLabelPosition.Auto,
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto },
                EntityColourCapability.Fill));

    [Fact]
    public void Style_preset_rejects_a_null_style_key() =>
        Assert.Throws<ArgumentNullException>(
            () => _ = new GanttStylePreset(null!, "Bad", string.Empty, string.Empty,
                GanttHatchPattern.None, 4, 0.5, "#000000", 0.75, 8, 0,
                GanttLabelPosition.Auto,
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto },
                EntityColourCapability.Fill));

    [Fact]
    public void GetPreset_rejects_an_unknown_style_key() =>
        Assert.Throws<ArgumentOutOfRangeException>(
            () => _ = GanttCatalogues.GetPreset("NotABuiltInKey"));

    [Fact]
    public void Style_preset_rejects_null_allowed_label_positions()
    {
        Assert.Throws<ArgumentNullException>(() => _ = new GanttStylePreset(
            "Bad", "Bad", string.Empty, string.Empty,
            GanttHatchPattern.None, 4, 0.5, "#000000", 0.75, 8, 0,
            GanttLabelPosition.Auto,
            allowedLabelPositions: null!,
            EntityColourCapability.Fill));
    }

    [Fact]
    public void Style_preset_rejects_an_empty_allowed_label_position_set()
    {
        Assert.Throws<ArgumentException>(() => _ = NewStylePreset(
            allowedLabelPositions: new HashSet<GanttLabelPosition>()));
    }

    [Fact]
    public void Style_preset_rejects_an_undefined_allowed_label_position()
    {
        Assert.Throws<ArgumentException>(() => _ = NewStylePreset(
            allowedLabelPositions: new HashSet<GanttLabelPosition> { (GanttLabelPosition)999 }));
    }

    [Fact]
    public void Style_preset_rejects_an_undefined_default_label_position()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _ = NewStylePreset(
            defaultLabelPosition: (GanttLabelPosition)999));
    }

    [Fact]
    public void Style_preset_rejects_a_default_label_position_outside_the_allowed_set()
    {
        Assert.Throws<ArgumentException>(() => _ = NewStylePreset(
            defaultLabelPosition: GanttLabelPosition.Inside,
            allowedLabelPositions: new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto }));
    }

    [Fact]
    public void Style_preset_rejects_an_unknown_colour_capability()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _ = NewStylePreset(
            colourCapability: (EntityColourCapability)16));
    }

    private static GanttStylePreset NewStylePreset(
        GanttLabelPosition defaultLabelPosition = GanttLabelPosition.Auto,
        IReadOnlySet<GanttLabelPosition>? allowedLabelPositions = null,
        EntityColourCapability colourCapability = EntityColourCapability.Fill) =>
        new(
            "TestStyle",
            "Test Style",
            string.Empty,
            string.Empty,
            GanttHatchPattern.None,
            4,
            0.5,
            "#000000",
            0.75,
            8,
            0,
            defaultLabelPosition,
            allowedLabelPositions
                ?? new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto },
            colourCapability);

    [Fact]
    public void GetPreset_rejects_a_null_style_key() =>
        Assert.Throws<ArgumentNullException>(
            () => _ = GanttCatalogues.GetPreset(null!));

    // ------------------------------------------------------------------
    // Pinned first-release hash (ADR-0007 D6 drift detection)
    // ------------------------------------------------------------------

    /// <summary>
    /// The pinned hash of the current schema-v2 catalogue as transcribed from
    /// the entity guide and ADR-0014. The hash changes whenever any code-owned
    /// catalogue value changes — that is the drift-detection contract
    /// (ADR-0007 D6): a failing pin means a deliberate catalogue change,
    /// which must first be an entity-guide change and a schema-version
    /// decision, then a deliberate pin update in the same commit.
    /// </summary>
    private const string PinnedFirstReleaseHash =
        "55aa653f5f0d54b5811140c9d2b07341fc4b86d663ad10987dc42180d1b75e20";

    [Fact]
    public void The_first_release_catalogue_hash_is_pinned() =>
        Assert.Equal(PinnedFirstReleaseHash, GanttCatalogues.ComputeCatalogueHash());
}
