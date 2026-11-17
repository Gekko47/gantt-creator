using GanttCreator.Core;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Pins the R5.6b chart-title validator: non-blank trimmed title, exact two-key
/// projection, and the blank-title positive refusal.
/// </summary>
public sealed class ChartTitleEditRequestTests
{
    [Fact]
    public void Blank_title_is_refused()
    {
        Assert.False(ChartTitleEditRequest.TryCreate(
            "   ", true, out _, out var refusal, out _));
        Assert.Equal(ChartTitleEditRefusal.BlankTitle, refusal);
    }

    [Fact]
    public void Null_title_is_refused()
    {
        Assert.False(ChartTitleEditRequest.TryCreate(
            null, true, out _, out var refusal, out _));
        Assert.Equal(ChartTitleEditRefusal.BlankTitle, refusal);
    }

    [Fact]
    public void Leading_and_trailing_spaces_are_trimmed_and_valid()
    {
        Assert.True(ChartTitleEditRequest.TryCreate(
            "  As-Built with Critical Path  ", false, out var request, out _, out _));
        Assert.NotNull(request);
        Assert.Equal("As-Built with Critical Path", request.Title);
        Assert.False(request.ShowTitle);
    }

    [Fact]
    public void No_length_cap_is_imposed()
    {
        var longTitle = new string('A', 500);
        Assert.True(ChartTitleEditRequest.TryCreate(
            longTitle, true, out var request, out _, out _));
        Assert.NotNull(request);
        Assert.Equal(longTitle, request.Title);
    }

    [Fact]
    public void Projection_writes_exactly_the_two_intended_keys()
    {
        var request = new ChartTitleEditRequest("Gantt Chart", true);
        IReadOnlyDictionary<string, string> settings = ChartTitleEditRequest.ToSettings(request);
        Assert.Equal(2, settings.Count);
        Assert.Equal("Gantt Chart", settings["ChartTitle"]);
        Assert.Equal("TRUE", settings["ShowTitle"]);
    }

    [Fact]
    public void Hidden_title_still_requires_a_nonblank_retained_value()
    {
        Assert.False(ChartTitleEditRequest.TryCreate(
            string.Empty, false, out _, out var refusal, out _));
        Assert.Equal(ChartTitleEditRefusal.BlankTitle, refusal);
    }
}
