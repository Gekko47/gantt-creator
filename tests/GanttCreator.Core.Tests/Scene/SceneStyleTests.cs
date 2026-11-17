using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for <see cref="SceneStyle"/>'s own validation. A style is the resolved
/// appearance every renderer consumes, so a value the host cannot express has to be
/// refused here rather than written to a shape and silently rounded by Office.
/// </summary>
public sealed class SceneStyleTests
{
    [Theory]
    [InlineData(-1.0)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void An_unusable_outline_width_is_refused(double outlineWidthPt)
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => new SceneStyle("Line", outlineWidthPt: outlineWidthPt));
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.0)]
    [InlineData(double.NaN)]
    public void An_unusable_font_size_is_refused(double fontSizePt)
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => new SceneStyle("Label", fontSizePt: fontSizePt));
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void A_blank_style_key_is_refused(string styleKey)
    {
        Assert.ThrowsAny<ArgumentException>(() => new SceneStyle(styleKey));
    }

    [Fact]
    public void A_usable_style_keeps_every_supplied_value()
    {
        SceneStyle style = new(
            "Activity",
            fillColour: ColourHex.Parse("#112233"),
            outlineWidthPt: 0,
            fontSizePt: 0.5,
            bold: true);

        Assert.Equal("Activity", style.StyleKey);
        Assert.Equal("#112233", style.FillColour!.ToString());
        Assert.Equal(0, style.OutlineWidthPt);
        Assert.Equal(0.5, style.FontSizePt);
        Assert.True(style.Bold);
    }
}
