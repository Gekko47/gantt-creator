using GanttCreator.Core;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Pins the R5.6 theme validator: catalogue ranges, colour format, and
/// non-blocking contrast advisories. Every refusal has its positive test.
/// </summary>
public sealed class ThemeEditRequestTests
{
    [Fact]
    public void A_valid_chart_chrome_colour_is_accepted()
    {
        // The positive control for the chrome-colour validator: a key the
        // catalogue defines, carrying an uppercase #RRGGBB, is accepted.
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            ChromeColourEdits: new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["ChartBackgroundColour"] = "#F0F0F0",
            });

        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }

    [Fact]
    public void An_unknown_chart_chrome_colour_key_is_refused()
    {
        // A key the catalogue does not define is refused: the writer would
        // otherwise write a row no reader could interpret.
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            ChromeColourEdits: new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["NotAChartColour"] = "#F0F0F0",
            });

        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.UnknownStyleKey, refusal);
    }

    [Fact]
    public void A_malformed_chart_chrome_colour_is_refused()
    {
        // A malformed colour is refused rather than substituted, so the chart
        // never renders a colour the workbook did not ask for.
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            ChromeColourEdits: new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["ChartBackgroundColour"] = "not-a-colour",
            });

        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.BadColourFormat, refusal);
    }

    [Fact]
    public void Empty_request_is_valid_with_no_warnings()
    {
        var request = new ThemeEditRequest([], new Dictionary<string, double>());
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
        Assert.Empty(ThemeEditRequest.ContrastWarnings(request));
    }

    [Fact]
    public void Empty_text_colour_is_accepted_as_preset_default()
    {
        // TextColour, like FillColour and StrokeColour, is optional: empty
        // means "use the preset default" rather than a format error.
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("DelayEvent", "#FF0000", "#C00000", "")],
            new Dictionary<string, double>());
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }

    [Fact]
    public void Unknown_style_key_is_refused()
    {
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("Nope", "#FF0000", "#C00000", "#FFFFFF")],
            new Dictionary<string, double>());
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.UnknownStyleKey, refusal);
    }

    [Fact]
    public void Lowercase_colour_is_refused()
    {
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("DelayEvent", "#ff0000", "#C00000", "#FFFFFF")],
            new Dictionary<string, double>());
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.BadColourFormat, refusal);
    }

    [Fact]
    public void Metric_outside_catalogue_range_is_refused_with_range_in_message()
    {
        // An editable metric, so the refusal is about the VALUE and not about
        // the token being off-limits. The bar height is no longer editable:
        // the scene resolves it from the style row, not from this metric.
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double> { ["DelineatorLinePt"] = 999 });
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Equal(ThemeEditRefusal.MetricOutOfRange, refusal);

        GanttMetricToken token = GanttCatalogues.Metrics
            .First(candidate => candidate.Name == "DelineatorLinePt");
        Assert.Contains(
            token.Minimum.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            message,
            StringComparison.Ordinal);
        Assert.Contains(
            token.Maximum.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void An_undefined_line_type_is_refused()
    {
        // The positive test for the guard: an undefined dash would serialise to
        // a name the reader cannot parse, and the reader refuses the whole
        // catalogue rather than skipping the row.
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("DefaultDelineator", "#FF0000", "#C00000", "#FFFFFF", (GanttStrokeDash)99)],
            new Dictionary<string, double>());
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
        Assert.Contains("line type", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_defined_line_type_is_accepted()
    {
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("DefaultDelineator", "#FF0000", "#C00000", "#FFFFFF", GanttStrokeDash.Dash)],
            new Dictionary<string, double>());
        Assert.True(ThemeEditRequest.TryValidate(request, out _, out _));
    }

    [Fact]
    public void Non_editable_metric_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double> { ["TitleBandHeightPt"] = 24 });
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.MetricOutOfRange, refusal);
    }

    [Fact]
    public void Low_contrast_edit_warns_without_blocking()
    {
        var request = new ThemeEditRequest(
            [new ThemeStyleEdit("DelayEvent", "#FF0000", "#C00000", "#FFFFFF")],
            new Dictionary<string, double>());
        Assert.True(ThemeEditRequest.TryValidate(request, out _, out _));
        Assert.NotEmpty(ThemeEditRequest.ContrastWarnings(request));
    }

    [Fact]
    public void Null_request_is_refused()
    {
        Assert.False(ThemeEditRequest.TryValidate(null, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.UnknownStyleKey, refusal);
    }

    [Fact]
    public void Custom_style_blank_key_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "  ")]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_builtin_key_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "DelayEvent")]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_duplicate_key_in_same_request_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [
                new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE"),
                new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE"),
            ]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_rename_builtin_target_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Rename, "MYSTYLE", NewStyleKey: "DelayEvent")]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_rename_missing_new_key_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Rename, "MYSTYLE", NewStyleKey: "  ")]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_bad_colour_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE", FillColour: "#ff0000")]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.BadColourFormat, refusal);
    }

    [Fact]
    public void Custom_style_undefined_hatch_pattern_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE", HatchPattern: (GanttHatchPattern)99)]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_undefined_default_label_position_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE", DefaultLabelPosition: (GanttLabelPosition)99)]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_empty_allowed_label_set_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [
                new CustomStyleEdit(
                    CustomStyleOperation.Create,
                    "MYSTYLE",
                    AllowedLabelPositions: new HashSet<GanttLabelPosition>()),
            ]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_default_label_position_outside_allowed_set_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [
                new CustomStyleEdit(
                    CustomStyleOperation.Create,
                    "MYSTYLE",
                    DefaultLabelPosition: GanttLabelPosition.Inside,
                    AllowedLabelPositions: new HashSet<GanttLabelPosition>
                    {
                        GanttLabelPosition.Auto,
                        GanttLabelPosition.Left,
                    }),
            ]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_size_outside_catalogue_range_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE", ActivityHeightPt: 10_000)]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.MetricOutOfRange, refusal);
    }

    [Fact]
    public void Custom_style_negative_size_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE", MilestoneSizePt: -1)]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.MetricOutOfRange, refusal);
    }

    [Fact]
    public void Custom_style_zero_size_is_accepted_as_not_applicable()
    {
        // Zero is the style contract's "not applicable to this entity kind"
        // value: an activity style carries a zero milestone size and a
        // milestone style carries a zero bar height. The catalogue's minimum
        // for both tokens is above zero, so this case must be handled
        // explicitly rather than by the range check.
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(
                CustomStyleOperation.Create,
                "MYSTYLE",
                ActivityHeightPt: 0,
                MilestoneSizePt: 0,
                StandardOutlinePt: 0)]);
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }

    [Fact]
    public void Custom_style_size_below_the_catalogue_minimum_is_refused()
    {
        // A non-zero size must be inside the token's range. ActivityHeightPt's
        // minimum is 2, so 1 is refused.
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE", ActivityHeightPt: 1)]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.MetricOutOfRange, refusal);
    }

    [Fact]
    public void Custom_style_undefined_stroke_dash_is_refused()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE", StrokeDash: (GanttStrokeDash)99)]);
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out _));
        Assert.Equal(ThemeEditRefusal.InvalidCustomStyle, refusal);
    }

    [Fact]
    public void Custom_style_with_a_stroke_dash_passes()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(
                CustomStyleOperation.Create,
                "MYSTYLE",
                FillColour: "#FF0000",
                StrokeDash: GanttStrokeDash.LongDash)]);
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }

    [Fact]
    public void Custom_style_with_full_property_set_passes()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [
                new CustomStyleEdit(
                    CustomStyleOperation.Create,
                    "MYSTYLE",
                    FillColour: "#FF0000",
                    StrokeColour: "#C00000",
                    TextColour: "#FFFFFF",
                    HatchPattern: GanttHatchPattern.Cross,
                    StandardOutlinePt: 0.75,
                    ActivityHeightPt: 10,
                    MilestoneSizePt: 12,
                    DefaultLabelPosition: GanttLabelPosition.Inside,
                    AllowedLabelPositions: new HashSet<GanttLabelPosition>
                    {
                        GanttLabelPosition.Auto,
                        GanttLabelPosition.Inside,
                        GanttLabelPosition.None,
                    }),
            ]);
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }

    [Fact]
    public void Custom_style_create_with_valid_key_and_colours_passes()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Create, "MYSTYLE", FillColour: "#FF0000", StrokeColour: "#C00000", TextColour: "#FFFFFF")]);
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }

    [Fact]
    public void Custom_style_rename_with_valid_keys_passes()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Rename, "OLDSTYLE", NewStyleKey: "NEWSTYLE")]);
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }

    [Fact]
    public void Custom_style_update_with_valid_key_passes()
    {
        // An upsert by key keeps the key, so it can never collide with itself:
        // a colour-only change on an existing custom style validates cleanly.
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Update, "MYSTYLE", FillColour: "#FF0000")]);
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }

    [Fact]
    public void Custom_style_delete_with_valid_key_passes()
    {
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(CustomStyleOperation.Delete, "MYSTYLE")]);
        Assert.True(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Null(refusal);
        Assert.Null(message);
    }
}
