using GanttCreator.Core;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Pins the R5.6 theme validator: catalogue ranges, colour format, and
/// non-blocking contrast advisories. Every refusal has its positive test.
/// </summary>
public sealed class ThemeEditRequestTests
{
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
        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double> { ["ActivityHeightPt"] = 999 });
        Assert.False(ThemeEditRequest.TryValidate(request, out var refusal, out var message));
        Assert.Equal(ThemeEditRefusal.MetricOutOfRange, refusal);
        Assert.Contains("2", message, StringComparison.Ordinal);
        Assert.Contains("36", message, StringComparison.Ordinal);
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
