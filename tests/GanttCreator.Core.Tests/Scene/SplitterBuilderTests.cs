using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the §10 splitter builder. Every refusal carries a positive case here,
/// because a validator without one is how an error path stops working silently.
/// </summary>
public sealed class SplitterBuilderTests
{
    private static readonly RectD _plot = new(200, 60, 300, 140);
    private static readonly ITextMetrics _metrics = new FakeTextMetrics(_ => 4.0, 10.0);
    private static readonly LaneGeometry _lane = new("row:splitter", 0, 0, 18, [], [], true, false);
    private static readonly SceneStyle _style = new("Splitter", fillColour: ColourHex.Parse("#FFE699"));
    private static readonly SceneStyle _borderStyle =
        new("MajorGrid", strokeColour: ColourHex.Parse("#808080"), outlineWidthPt: 99);
    private static readonly SceneStyle _labelStyle = new("DefaultText", strokeColour: ColourHex.Parse("#000000"));

    private static GanttEvent Splitter(string description = "Work Package A") =>
        new(
            1,
            GanttRowId.New(),
            null,
            null,
            GanttEntityType.Splitter,
            description,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            true,
            null);

    private static SplitterRequest Request(
        GanttLabelPosition position = GanttLabelPosition.PlotCentre,
        double panelLeft = 0,
        double borderWidthPt = 1) =>
        new(Splitter(), _style, _borderStyle, _lane, panelLeft, _plot, borderWidthPt, _labelStyle, _metrics, position);

    private static SplitterResult Build(SplitterRequest request)
    {
        SplitterCreationOutcome outcome = SplitterBuilder.TryBuild(request);
        Assert.True(outcome.Succeeded, "Splitter build refused: " + outcome.Refusal);
        return outcome.Result!;
    }

    [Fact]
    public void The_band_spans_the_plot_only()
    {
        SplitterResult result = Build(Request());

        // Owner ruling 2026-11-19: §10's fill marks the splitter's lane across the PLOT.
        // The band used to start at the data panel's left edge, which is invisible on the
        // live sheet and painted a full-width stripe across the table the moment the
        // export composition drew the panel.
        Assert.Equal(_plot.Left, result.BandBounds.X);
        Assert.Equal(_plot.Right, result.BandBounds.Right);
        Assert.Equal(18, result.BandBounds.Height);
    }

    [Fact]
    public void The_band_sits_at_the_section_layer_with_its_borders_at_the_frame_layer()
    {
        SplitterResult result = Build(Request());

        SceneRect band = Assert.IsType<SceneRect>(result.Primitives[0]);
        Assert.Equal(ZLayer.Section, band.ZLayer);
        Assert.Equal(ColourHex.Parse("#FFE699"), band.Style.FillColour);

        SceneLine[] borders = [.. result.Primitives.OfType<SceneLine>()];
        Assert.Equal(2, borders.Length);
        Assert.All(borders, border => Assert.Equal(ZLayer.Frame, border.ZLayer));
        Assert.Equal(result.BandBounds.Y, borders[0].From.Y);
        Assert.Equal(result.BandBounds.Bottom, borders[1].From.Y);
    }

    [Fact]
    public void Both_borders_are_stroked_with_the_border_style_at_the_major_boundary_width()
    {
        // §10: "Its top and bottom borders are `MajorBoundaryPt` lines". The border style
        // supplies only the stroke token, so a line carries a colour and a width but no
        // fill. `_borderStyle` deliberately declares a width of 99 to prove the structural
        // `BorderWidthPt` wins over whatever the style happens to carry - reusing the
        // style wholesale would silently render a 99pt border.
        SplitterResult result = Build(Request(borderWidthPt: 1.5));

        SceneLine[] borders = [.. result.Primitives.OfType<SceneLine>()];
        Assert.Equal(2, borders.Length);
        Assert.All(borders, border =>
        {
            Assert.Equal(1.5, border.Style.OutlineWidthPt);
            Assert.NotNull(border.Style.StrokeColour);
            Assert.Equal(ColourHex.Parse("#808080"), border.Style.StrokeColour);
        });

        // The band is a fill and keeps the Splitter preset; the border is a stroke and must
        // not have inherited it. Sharing one style across both is what this asserts against.
        SceneRect band = Assert.IsType<SceneRect>(result.Primitives[0]);
        Assert.Equal(ColourHex.Parse("#FFE699"), band.Style.FillColour);
        Assert.All(borders, border => Assert.NotEqual(band.Style, border.Style));
    }

    [Fact]
    public void A_null_border_style_is_refused_rather_than_defaulting_the_stroke()
    {
        // Without a stroke token the builder would have to invent a colour, which is the
        // exact dual-source problem the band preset caused. It refuses instead.
        Assert.Equal(
            SplitterRefusal.NullDependency,
            SplitterBuilder.TryBuild(Request() with { BorderStyle = null! }).Refusal);
    }

    [Fact]
    public void A_label_position_with_no_room_is_suppressed_rather_than_emitted_zero_width()
    {
        // With no data panel the band starts at the plot's own left edge, so the
        // panel-side half of `Both` has exactly zero width to sit in. A zero-width text
        // box is an invisible primitive a renderer must still place and reconciliation
        // must track forever, so it is suppressed and only the plot-centre label remains.
        SplitterCreationOutcome outcome = SplitterBuilder.TryBuild(
            new SplitterRequest(
                Splitter(),
                _style,
                _borderStyle,
                _lane,
                200,   // PanelLeftPt == Plot.Left: no panel to its left
                _plot,
                1,
                _labelStyle,
                _metrics,
                GanttLabelPosition.Both));

        Assert.True(outcome.Succeeded, "Splitter build refused: " + outcome.Refusal);
        _ = Assert.Single(outcome.Result!.Primitives.OfType<SceneText>());
    }

    /// <summary>
    /// `DataPanelLeft` is no longer a splitter capability (R5.18l, D4), and the builder
    /// refuses it rather than drawing the band and silently dropping the label.
    /// </summary>
    [Fact]
    public void A_direct_left_alignment_is_refused_because_the_splitter_no_longer_offers_it()
    {
        SplitterCreationOutcome outcome = SplitterBuilder.TryBuild(
            new SplitterRequest(
                Splitter(),
                _style,
                _borderStyle,
                _lane,
                0,
                _plot,
                1,
                _labelStyle,
                _metrics,
                GanttLabelPosition.DataPanelLeft));

        Assert.False(outcome.Succeeded);
        Assert.Equal(SplitterRefusal.UnsupportedLabelPosition, outcome.Refusal);
    }

    [Theory]
    [InlineData(GanttLabelPosition.PlotCentre, 1)]
    [InlineData(GanttLabelPosition.Both, 2)]
    [InlineData(GanttLabelPosition.None, 0)]
    public void Each_permitted_label_position_emits_the_guide_s_primitive_count(
        GanttLabelPosition position,
        int expected)
    {
        // §10 names exactly these three positions since R5.18l, and `Both` "deliberately
        // creates two scene text entities with stable role-derived IDs".
        SplitterResult result = Build(Request(position));

        Assert.Equal(expected, result.Primitives.OfType<SceneText>().Count());
    }

    [Fact]
    public void Both_emits_two_texts_with_distinct_role_derived_ids()
    {
        SplitterResult result = Build(Request(GanttLabelPosition.Both));

        SceneText[] labels = [.. result.Primitives.OfType<SceneText>()];
        Assert.Equal(2, labels.Length);
        string owner = labels[0].OwnerId.Value;
        Assert.Equal(
            [$"{owner}:{SplitterBuilder.LabelRole}", $"{owner}:{SplitterBuilder.PlotLabelRole}"],
            labels.Select(label => label.PrimitiveId));
    }

    [Fact]
    public void A_blank_description_emits_a_band_but_no_label()
    {
        // Blank is legal data, and an empty text primitive is noise a renderer has to
        // special-case — the same rule PanelBuilder applies to a blank cell.
        SplitterCreationOutcome outcome = SplitterBuilder.TryBuild(Request() with { Event = Splitter("   ") });

        Assert.True(outcome.Succeeded);
        Assert.Empty(outcome.Result!.Primitives.OfType<SceneText>());
        Assert.Equal(GanttLabelPosition.None, outcome.Result.LabelPosition);
    }

    [Fact]
    public void Every_refusal_has_a_positive_case()
    {
        Assert.Equal(SplitterRefusal.NullRequest, SplitterBuilder.TryBuild(null).Refusal);
        Assert.Equal(
            SplitterRefusal.NullDependency,
            SplitterBuilder.TryBuild(Request() with { Event = null! }).Refusal);
        Assert.Equal(
            SplitterRefusal.NotASplitter,
            SplitterBuilder.TryBuild(
                Request() with { Event = Splitter() with { Type = GanttEntityType.AsPlannedActivity } }).Refusal);

        // A position the catalogue does not permit for a splitter is refused rather
        // than silently reinterpreted.
        Assert.Equal(
            SplitterRefusal.UnsupportedLabelPosition,
            SplitterBuilder.TryBuild(Request(GanttLabelPosition.Right)).Refusal);
        Assert.Equal(
            SplitterRefusal.UnsupportedLabelPosition,
            SplitterBuilder.TryBuild(Request((GanttLabelPosition)99)).Refusal);

        // A non-finite edge, a negative border, and a zero-height lane are all refused
        // before any primitive is emitted.
        Assert.Equal(
            SplitterRefusal.InvalidGeometry,
            SplitterBuilder.TryBuild(Request(panelLeft: double.NaN)).Refusal);
        Assert.Equal(
            SplitterRefusal.InvalidGeometry,
            SplitterBuilder.TryBuild(Request(borderWidthPt: -1)).Refusal);
        Assert.Equal(
            SplitterRefusal.InvalidGeometry,
            SplitterBuilder.TryBuild(Request() with { Lane = _lane with { Height = 0 } }).Refusal);
    }
}
