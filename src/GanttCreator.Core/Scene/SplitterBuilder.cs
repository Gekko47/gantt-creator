namespace GanttCreator.Core.Scene;

/// <summary>One splitter request.</summary>
/// <param name="Event">The validated Splitter event. No date geometry is read.</param>
/// <param name="Style">The resolved named style, supplying <c>SplitterFill</c> (§10).</param>
/// <param name="BorderStyle">
/// The resolved major-boundary style supplying the borders' stroke token. Its own width is
/// deliberately not read; see <see cref="BorderWidthPt"/>.
/// </param>
/// <param name="Lane">The fixed-height lane the splitter occupies.</param>
/// <param name="PanelLeftPt">The data panel's left edge, so the band spans panel and plot.</param>
/// <param name="PlotBounds">The plot rectangle; the band's right edge is its right edge.</param>
/// <param name="BorderWidthPt">
/// The major top/bottom border width, from <c>MajorBoundaryPt</c>. §10's border is a
/// <em>major</em> boundary, so this is a structural token rather than a style-derived width
/// (ADR-0022 D4), and it is the value written to the borders' <c>OutlineWidthPt</c>.
/// </param>
/// <param name="LabelStyle">The resolved label style, or <see langword="null"/> for no label.</param>
/// <param name="LabelMetrics">The injected deterministic text-measuring seam.</param>
/// <param name="LabelPosition">The resolved label position, or <c>None</c> for no label.</param>
public sealed record SplitterRequest(
    GanttEvent Event,
    SceneStyle Style,
    SceneStyle BorderStyle,
    LaneGeometry Lane,
    double PanelLeftPt,
    RectD PlotBounds,
    double BorderWidthPt,
    SceneStyle? LabelStyle,
    ITextMetrics LabelMetrics,
    GanttLabelPosition LabelPosition = GanttLabelPosition.DataPanelLeft
);

/// <summary>The result of building one splitter.</summary>
/// <param name="Primitives">
/// The band, its two borders, and zero, one, or two labels in deterministic order.
/// </param>
/// <param name="BandBounds">The resolved band rectangle.</param>
/// <param name="LabelPosition">The position actually used, after the <c>None</c> collapse.</param>
public sealed record SplitterResult(
    IReadOnlyList<ScenePrimitive> Primitives,
    RectD BandBounds,
    GanttLabelPosition LabelPosition
);

/// <summary>The reason a splitter could not be built.</summary>
public enum SplitterRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The event, style, lane, or text-metrics seam was null.</summary>
    NullDependency = 1,

    /// <summary>A bound, the border width, or the lane height was not usable.</summary>
    InvalidGeometry = 2,

    /// <summary>The event is not a splitter type.</summary>
    NotASplitter = 3,

    /// <summary>The label position is not one the entity catalogue permits for a splitter.</summary>
    UnsupportedLabelPosition = 4,
}

/// <summary>The typed result of attempting to build one splitter.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record SplitterCreationOutcome(SplitterResult? Result, SplitterRefusal? Refusal)
{
    /// <summary>Gets whether construction succeeded.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Builds the §10 splitter band, its major borders, and its label(s), with no
/// Office dependency.
/// </summary>
/// <remarks>
/// <para>
/// §10 geometry: the splitter "occupies a complete lane across the included data
/// panel and plot", so the band spans <c>PanelLeftPt</c> to the plot's right edge
/// rather than the plot alone. It has no date geometry, so no date is read and the
/// horizontal extent is structural rather than time-scaled.
/// </para>
/// <para>
/// The band sits at <see cref="ZLayer.Section"/> (30), which the guide's z-order table
/// already assigns to "splitter/section backgrounds", and the two major borders sit at
/// <see cref="ZLayer.Frame"/> so they frame the band rather than being covered by it.
/// </para>
/// <para>
/// The band and the borders take their styles from different sources, because §10
/// specifies them differently. The band is a <em>fill</em> and takes the resolved
/// <c>SplitterFill</c> preset; each border is a <em>stroke</em> whose width is the
/// structural <c>MajorBoundaryPt</c> token (ADR-0022 D4), so the border takes its colour
/// from the supplied <c>BorderStyle</c> and its width from <c>BorderWidthPt</c>. Carrying
/// one style across both would have forced the border to be either unfillable or wrongly
/// filled, and would have left the structural width unused.
/// </para>
/// <para>
/// The label positions belong to this builder rather than to <see cref="LabelPlanner"/>:
/// the planner's §22 cascade places a label beside an entity's own bounds, which is a
/// different problem from "left edge of the data panel" or "centre of the plot". This
/// mirrors how <see cref="DelineatorBuilder"/> owns its corner labels, and the shared
/// <see cref="ITextMetrics"/> seam is still used so measurement is identical across
/// every label in the scene.
/// </para>
/// </remarks>
public static class SplitterBuilder
{
    /// <summary>The §10 role suffix for the band rectangle.</summary>
    public const string BandRole = "splitter-band";

    /// <summary>The §10 role suffix for the data-panel label.</summary>
    public const string LabelRole = "splitter-label";

    /// <summary>The §10 role suffix for the plot-centre label emitted by <c>Both</c>.</summary>
    public const string PlotLabelRole = "splitter-label-plot";

    /// <summary>Attempts to build one splitter band, its borders, and its labels.</summary>
    /// <param name="request">The typed splitter request.</param>
    /// <returns>A typed result or refusal.</returns>
    public static SplitterCreationOutcome TryBuild(SplitterRequest? request)
    {
        if (request is null)
        {
            return Refused(SplitterRefusal.NullRequest);
        }

        if (request.Event is null
            || request.Style is null
            || request.BorderStyle is null
            || request.Lane is null
            || request.LabelMetrics is null)
        {
            return Refused(SplitterRefusal.NullDependency);
        }

        if (request.Event.Type != GanttEntityType.Splitter)
        {
            return Refused(SplitterRefusal.NotASplitter);
        }

        if (!Enum.IsDefined(request.LabelPosition)
            || !EntityTypeCatalog.GetDefinition(GanttEntityType.Splitter)!.AllowedLabelPositions
                .Contains(request.LabelPosition))
        {
            return Refused(SplitterRefusal.UnsupportedLabelPosition);
        }

        LaneGeometry lane = request.Lane;
        if (!double.IsFinite(request.PanelLeftPt)
            || !double.IsFinite(request.BorderWidthPt)
            || request.BorderWidthPt < 0
            || !double.IsFinite(lane.Top)
            || !double.IsFinite(lane.Height)
            || lane.Height <= 0
            || !double.IsFinite(request.PlotBounds.Right))
        {
            return Refused(SplitterRefusal.InvalidGeometry);
        }

        var owner = SceneOwnerId.ForRow(request.Event.Id);
        var bounds = new RectD(
            request.PanelLeftPt,
            lane.Top,
            request.PlotBounds.Right - request.PanelLeftPt,
            lane.Height);

        // The border is a *stroke*, not the band's fill. Reusing `request.Style` gave both
        // border lines the Splitter preset's fill with `StrokeColour` and `OutlineWidthPt`
        // both null, so the scene described a §10 major boundary that no renderer could draw
        // and silently discarded the `BorderWidthPt` the caller had already supplied and this
        // method had already validated. The style now supplies the stroke token and the
        // structural `BorderWidthPt` supplies the width, which is exactly the split ADR-0022
        // D4 draws: the colour is a style decision, the width is a structural measurement.
        // Only the stroke members are carried across; a line has no fill, and copying the
        // band's fill onto it would re-introduce the same dual meaning.
        SceneStyle borderStyle = new(
            request.BorderStyle.StyleKey,
            strokeColour: request.BorderStyle.StrokeColour,
            outlineWidthPt: request.BorderWidthPt);

        List<ScenePrimitive> primitives =
        [
            new SceneRect(
                ScenePrimitive.CreateId(owner, BandRole),
                owner,
                ZLayer.Section,
                bounds,
                request.Style),
            new SceneLine(
                ScenePrimitive.CreateId(owner, $"{BandRole}:top"),
                owner,
                ZLayer.Frame,
                new PointD(bounds.X, bounds.Y),
                new PointD(bounds.Right, bounds.Y),
                borderStyle,
                request.Event.Type,
                lane.LaneOrder),
            new SceneLine(
                ScenePrimitive.CreateId(owner, $"{BandRole}:bottom"),
                owner,
                ZLayer.Frame,
                new PointD(bounds.X, bounds.Bottom),
                new PointD(bounds.Right, bounds.Bottom),
                borderStyle,
                request.Event.Type,
                lane.LaneOrder),
        ];

        GanttLabelPosition position = AddLabels(request, owner, lane, bounds, primitives);

        return new SplitterCreationOutcome(
            new SplitterResult(primitives, bounds, position),
            null);
    }

    /// <summary>Appends the §10 labels for the requested position.</summary>
    /// <param name="request">The typed splitter request.</param>
    /// <param name="owner">The row owner every emitted primitive carries.</param>
    /// <param name="lane">The fixed-height lane, supplying the ordering value.</param>
    /// <param name="bounds">The resolved band rectangle.</param>
    /// <param name="primitives">The primitive list to append to.</param>
    /// <returns>The position actually rendered, after the <c>None</c> collapse.</returns>
    /// <remarks>
    /// A blank description renders no label rather than an empty text primitive, which
    /// is the same rule <see cref="PanelBuilder"/> applies to a blank cell: blank is
    /// legal data and an empty text entity is noise a renderer has to special-case.
    /// </remarks>
    private static GanttLabelPosition AddLabels(
        SplitterRequest request,
        SceneOwnerId owner,
        LaneGeometry lane,
        RectD bounds,
        List<ScenePrimitive> primitives)
    {
        if (request.LabelPosition == GanttLabelPosition.None || request.LabelStyle is not { } labelStyle)
        {
            return GanttLabelPosition.None;
        }

        var text = request.Event.Description;
        if (string.IsNullOrWhiteSpace(text)
            || !request.LabelMetrics.TryMeasure(text, out TextMeasurement? measured)
            || measured is null)
        {
            return GanttLabelPosition.None;
        }

        var height = measured.HeightPt;
        var top = bounds.Y + ((bounds.Height - height) / 2);

        // `Both` deliberately emits two scene text entities with distinct
        // role-derived IDs, so a reconciling renderer can update one without
        // disturbing the other.
        //
        // A label with no room is suppressed rather than emitted at zero width. This
        // is not hypothetical: in a scene with no data panel, `DataPanelLeft` has
        // nowhere to sit - the band starts at the plot's own left edge, so the space
        // between them is exactly zero. A zero-width text box is an invisible
        // primitive that a renderer would still have to place, and it would occupy a
        // scene slot that reconciliation then tracks forever. A blank/degenerate
        // primitive is the same class of defect this repository treats as a bug.
        if (request.LabelPosition is GanttLabelPosition.DataPanelLeft or GanttLabelPosition.Both)
        {
            var width = Math.Min(measured.WidthPt, request.PlotBounds.Left - bounds.X);
            if (width > 0)
            {
                primitives.Add(
                    new SceneText(
                        ScenePrimitive.CreateId(owner, LabelRole),
                        owner,
                        ZLayer.Label,
                        text,
                        new RectD(bounds.X, top, width, height),
                        labelStyle,
                        GanttTextAlignment.Left,
                        request.Event.Type,
                        lane.LaneOrder));
            }
        }

        if (request.LabelPosition is GanttLabelPosition.PlotCentre or GanttLabelPosition.Both)
        {
            var width = Math.Min(measured.WidthPt, request.PlotBounds.Width);
            if (width <= 0)
            {
                return primitives.Count > 0 ? request.LabelPosition : GanttLabelPosition.None;
            }

            primitives.Add(
                new SceneText(
                    ScenePrimitive.CreateId(owner, PlotLabelRole),
                    owner,
                    ZLayer.Label,
                    text,
                    new RectD(
                        request.PlotBounds.X + ((request.PlotBounds.Width - width) / 2),
                        top,
                        width,
                        height),
                    labelStyle,
                    GanttTextAlignment.Centre,
                    request.Event.Type,
                    lane.LaneOrder));
        }

        return request.LabelPosition;
    }

    private static SplitterCreationOutcome Refused(SplitterRefusal refusal) => new(null, refusal);
}
