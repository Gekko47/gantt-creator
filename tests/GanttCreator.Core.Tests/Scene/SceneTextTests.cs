using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Contracts for <see cref="SceneText.LeftInsetPt"/>: the indent the editable export
/// consumes as the text frame's left margin.
/// </summary>
public sealed class SceneTextTests
{
    private static readonly SceneOwnerId _owner = SceneOwnerId.Chart;
    private static readonly SceneStyle _style = new("Text");
    private static readonly RectD _bounds = new(0, 0, 20, 10);

    [Fact]
    public void The_default_left_inset_is_zero()
    {
        // Existing call sites pass no inset, and they must stay byte-identical.
        SceneText text = new("chart:text", _owner, ZLayer.Label, "t", _bounds, _style, GanttTextAlignment.Left);

        Assert.Equal(0, text.LeftInsetPt);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void A_negative_or_non_finite_left_inset_is_refused(double inset)
    {
        // Positive test for the new validator: an inset outside the text box is not a
        // value the renderer can honour, so it is refused at construction.
        _ = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new SceneText(
                "chart:text",
                _owner,
                ZLayer.Label,
                "t",
                _bounds,
                _style,
                GanttTextAlignment.Left,
                leftInsetPt: inset));
    }

    [Fact]
    public void A_positive_left_inset_is_kept()
    {
        SceneText text = new(
            "chart:text",
            _owner,
            ZLayer.Label,
            "t",
            _bounds,
            _style,
            GanttTextAlignment.Left,
            leftInsetPt: 6);

        Assert.Equal(6, text.LeftInsetPt);
    }
}
