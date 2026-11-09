using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class HostOwnedGeometryTests
{
    private static readonly RectD _plot = new(100, 50, 400, 200);
    private static readonly SceneStyle _style = new("Test");
    private static readonly SceneOwnerId _owner = SceneOwnerId.ForRow(GanttRowId.New());

    [Fact]
    public void A_vertical_line_spanning_the_plot_is_host_owned()
    {
        SceneLine line = new(
            "G-1:delineator",
            _owner,
            ZLayer.Delineator,
            new PointD(150, _plot.Top),
            new PointD(150, _plot.Bottom),
            _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(line, _plot));
    }

    [Fact]
    public void A_horizontal_line_is_not_host_owned()
    {
        SceneLine line = new(
            "G-1:rule",
            _owner,
            ZLayer.Grid,
            new PointD(_plot.Left, 100),
            new PointD(_plot.Right, 100),
            _style);

        Assert.False(HostOwnedGeometry.IsHostOwned(line, _plot));
    }

    [Fact]
    public void A_vertical_line_shorter_than_the_plot_is_not_host_owned()
    {
        SceneLine line = new(
            "G-1:short",
            _owner,
            ZLayer.Delineator,
            new PointD(150, _plot.Top),
            new PointD(150, _plot.Bottom - 10),
            _style);

        Assert.False(HostOwnedGeometry.IsHostOwned(line, _plot));
    }

    [Fact]
    public void A_rect_spanning_the_plot_height_is_host_owned()
    {
        SceneRect rect = new(
            "G-1:band",
            _owner,
            ZLayer.AlternateBand,
            new RectD(120, _plot.Top, 40, _plot.Height),
            _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(rect, _plot));
    }

    [Fact]
    public void A_bar_rect_is_not_host_owned()
    {
        SceneRect rect = new(
            "G-1:bar",
            _owner,
            ZLayer.ActivityBody,
            new RectD(120, 80, 40, 12),
            _style);

        Assert.False(HostOwnedGeometry.IsHostOwned(rect, _plot));
    }

    [Fact]
    public void A_text_primitive_is_not_host_owned()
    {
        SceneText text = new(
            "G-1:label",
            _owner,
            ZLayer.Label,
            "hello",
            new RectD(120, 80, 40, 12),
            _style,
            GanttTextAlignment.Left);

        Assert.False(HostOwnedGeometry.IsHostOwned(text, _plot));
    }

    [Fact]
    public void A_full_height_line_within_the_geometry_tolerance_is_host_owned()
    {
        SceneLine line = new(
            "G-1:delineator",
            _owner,
            ZLayer.Delineator,
            new PointD(150, _plot.Top + 0.005),
            new PointD(150, _plot.Bottom - 0.005),
            _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(line, _plot));
    }

    /// <summary>
    /// A real delineator runs from the plot top LIFTED into the header to the plot bottom
    /// EXTENDED through the reserved anchor row, so it is host-owned.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the case that made the live chart wrong while every test passed. The
    /// builders do not emit plot-bounds geometry: <see cref="PlotSpanGeometry"/>
    /// lifts the top by <c>PlotBandHeaderOverlapPt</c> (ADR-0037) and extends the bottom
    /// by <c>anchorRow + boundaryWidth/2</c> (ADR-0038), and
    /// <c>SceneBuilder</c> hands both to the delineator. A predicate that demands
    /// EQUALITY with <c>PlotBounds</c> therefore rejects every real delineator and every
    /// real band, leaving them on <c>xlMove</c> — no live stretch at all.
    /// </para>
    /// <para>
    /// The contract is coverage, not equality: a primitive that spans at least the plot's
    /// vertical extent is host-owned, whatever the builders added at either end.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_delineator_extended_into_the_anchor_row_is_host_owned()
    {
        const double overlapPt = 0.5;
        const double anchorRowHeightPt = 18;
        const double boundaryWidthPt = 1;

        double top = PlotSpanGeometry.LiftTopIntoHeader(_plot, overlapPt).Top;
        double bottom = PlotSpanGeometry.ClosingLineBottomPt(
            _plot.Bottom,
            anchorRowHeightPt,
            boundaryWidthPt);

        SceneLine line = new(
            "G-1:delineator",
            _owner,
            ZLayer.Delineator,
            new PointD(150, top),
            new PointD(150, bottom),
            _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(line, _plot));
    }

    /// <summary>
    /// A real band is the plot rectangle lifted into the header, so it is host-owned.
    /// </summary>
    [Fact]
    public void A_band_lifted_into_the_header_is_host_owned()
    {
        RectD lifted = PlotSpanGeometry.LiftTopIntoHeader(_plot, overlapPt: 0.5);

        SceneRect band = new("G-1:band", _owner, ZLayer.AlternateBand, lifted, _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(band, _plot));
    }

    /// <summary>
    /// A band extended through the anchor row is host-owned.
    /// </summary>
    [Fact]
    public void A_band_extended_into_the_anchor_row_is_host_owned()
    {
        RectD extended = PlotSpanGeometry.ExtendBottomIntoAnchorRow(
            _plot,
            anchorRowHeightPt: 18,
            majorBoundaryPt: 1);

        SceneRect band = new("G-1:band", _owner, ZLayer.AlternateBand, extended, _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(band, _plot));
    }
}
