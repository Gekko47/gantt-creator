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

    /// <summary>
    /// A horizontal rule spanning the plot's width is host-owned (ADR-0052 D6, amended
    /// 2026-11-13), so Excel's <c>xlMoveAndSize</c> can stretch it when a column widens.
    /// </summary>
    /// <remarks>
    /// The live defect this pins: the chart's upper and lower frame rules stayed at their
    /// old width after a column resize, because <c>xlMove</c> translates a shape and never
    /// resizes it, and Excel has no two-edge anchor. The class used to exclude horizontal
    /// lines outright, so the exclusion itself was the bug.
    /// </remarks>
    [Fact]
    public void A_horizontal_rule_spanning_the_plot_width_is_host_owned()
    {
        SceneLine line = new(
            "chart:frame:top",
            _owner,
            ZLayer.Grid,
            new PointD(_plot.Left, 100),
            new PointD(_plot.Right, 100),
            _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(line, _plot));
    }

    [Fact]
    public void A_horizontal_rule_that_stops_short_of_the_plot_is_not_host_owned()
    {
        // A bar-length rule or a label-width rule has nothing the host could stretch with
        // the worksheet, so it stays on xlMove translation.
        SceneLine line = new(
            "G-1:short-rule",
            _owner,
            ZLayer.Grid,
            new PointD(_plot.Left, 100),
            new PointD(_plot.Right - 10, 100),
            _style);

        Assert.False(HostOwnedGeometry.IsHostOwned(line, _plot));
    }

    [Fact]
    public void A_horizontal_rule_within_the_geometry_tolerance_is_host_owned()
    {
        // Same coverage contract as the vertical class: the frame rules are built on the
        // content bounds, which can be wider than the plot by a rounding fraction.
        SceneLine line = new(
            "chart:frame:top",
            _owner,
            ZLayer.Grid,
            new PointD(_plot.Left + 0.005, 100),
            new PointD(_plot.Right - 0.005, 100),
            _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(line, _plot));
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
    public void A_line_exactly_on_the_geometry_tolerance_still_counts_as_covering_the_plot()
    {
        // The coverage comparisons are inclusive AT the tolerance, not strictly inside it: a
        // builder that lands a frame one tolerance short of the plot edge has still covered it,
        // and a strict comparison would drop that rule back to translation with nothing visible
        // changing in the scene.
        const double tolerancePt = HostOwnedGeometry.GeometryMatchTolerancePt;

        SceneLine horizontal = new(
            "chart:frame:top",
            _owner,
            ZLayer.Grid,
            new PointD(_plot.Left + tolerancePt, 100),
            new PointD(_plot.Right - tolerancePt, 100),
            _style);

        SceneLine vertical = new(
            "G-1:delineator",
            _owner,
            ZLayer.Delineator,
            new PointD(150, _plot.Top + tolerancePt),
            new PointD(150, _plot.Bottom - tolerancePt),
            _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(horizontal, _plot));
        Assert.True(HostOwnedGeometry.IsHostOwned(vertical, _plot));
    }

    [Fact]
    public void A_rule_sloped_exactly_by_the_geometry_tolerance_is_still_horizontal()
    {
        // "Horizontal" is decided within the geometry tolerance, so a rule drawn with a
        // sub-display-precision slope is still a horizontal rule and still host-sized. A strict
        // comparison would silently leave it on xlMove.
        const double tolerancePt = HostOwnedGeometry.GeometryMatchTolerancePt;

        // Anchored at y=0 so the difference the predicate takes is exactly the tolerance
        // rather than `100.01 - 100`, which is not representable and would make this a test
        // of binary floating point instead of the comparison under test.
        SceneLine line = new(
            "chart:frame:top",
            _owner,
            ZLayer.Grid,
            new PointD(_plot.Left, 0),
            new PointD(_plot.Right, tolerancePt),
            _style);

        Assert.True(HostOwnedGeometry.IsHostOwned(line, _plot));
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

    [Fact]
    public void A_null_primitive_is_refused_rather_than_classified()
    {
        // Trust boundary. Without the guard the switch's discard arm answers `false` for a
        // null, so a caller that lost its primitive would be told the shape is simply "not
        // host-owned" and would write it as an ordinary xlMove shape instead of stopping. The
        // guard is the difference between a silent mis-classification and a failure.
        _ = Assert.Throws<ArgumentNullException>(() => HostOwnedGeometry.IsHostOwned(null!, _plot));
    }
}
