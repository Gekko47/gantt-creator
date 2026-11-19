using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.8 label parity: the editable copy places the same description and date labels, at the same
/// offsets, as the live chart.
/// </summary>
/// <remarks>
/// <para>
/// <b>What this isolates.</b> Both profiles share <see cref="SceneBuilder"/>; the copy diverges
/// only in the request it hands the builder (zero-origin bounds, chrome padding, carried anchor
/// row) and in the <see cref="SceneCompositionProfile"/> it selects. So a label that moves or
/// disappears between the two scenes is caused by geometry, not by the label builders.
/// </para>
/// <para>
/// <b>Why the boxes are compared SHAPE-relative, not absolutely.</b> The live composition is
/// laid out on the worksheet's measured coordinates and the copy on a zero-origin canvas with a
/// uniform outer chrome, so absolute boxes differ by design. Measured on this fixture: the copy's
/// plot is 0.4516pt narrower/offset because its left outer padding is part of the width budget,
/// which moves every date-scaled X by the same fraction. That is composition geometry, not label
/// placement. Every label is anchored to its own shape's edge (a bar's end, a delineator's line),
/// so the offset from that edge is invariant to both origin and scale while still changing
/// completely if the copy resolves a different corner. That offset is what these facts compare.
/// </para>
/// <para>
/// <b>Why the style default is a different position.</b> Each fixture row names an explicit
/// position while its named style defaults to another one. A composition that silently falls back
/// to the style default (the current copy defect for style-key rows) moves the label and fails
/// these facts, so they are the scene-level half of the copy label parity regression.
/// </para>
/// </remarks>
public class ExportLiveLabelParityTests
{
    private const double _tolerance = 0.001;

    private static readonly DateOnly _start = new(2024, 1, 8);
    private static readonly DateOnly _finish = new(2024, 1, 19);
    private static readonly GanttRowId _spanId = GanttRowId.New();
    private static readonly GanttRowId _delineatorId = GanttRowId.New();

    /// <summary>A span row's explicit Right label is placed Right by the copy, exactly as live.</summary>
    [Fact]
    public void The_copy_places_an_explicit_span_label_like_the_live_chart()
    {
        GanttEvent row = SpanEvent();

        GanttScene live = BuildLiveScene([row]);
        GanttScene export = BuildExportScene([row]);
        static PointD BarLeftOrigin(GanttScene scene)
        {
            RectD bar = Bar(scene, _spanId);
            return new PointD(bar.Left, bar.Top);
        }

        static PointD BarRightOrigin(GanttScene scene)
        {
            RectD bar = Bar(scene, _spanId);
            return new PointD(bar.Right, bar.Top);
        }

        SceneText liveDescription = AssertSameLabelOffset(
            live, export, $"{_spanId.Value}:label", BarRightOrigin);

        // The explicit Right is the whole point: the style default is Inside. A copy that fell
        // back to the style would centre the box on the bar and fail the offset comparison above.
        RectD bar = Bar(live, _spanId);
        Assert.True(
            liveDescription.TextBounds.Left >= bar.Right - _tolerance,
            $"The live description label is not right of its bar: {liveDescription.TextBounds} vs {bar}.");

        _ = AssertSameLabelOffset(
            live, export, $"{_spanId.Value}:{DateLabelBuilder.StartRole}", BarLeftOrigin);
    }

    /// <summary>
    /// A delineator row's explicit corner labels are placed at the same corners by the copy.
    /// </summary>
    /// <remarks>
    /// The description takes <c>TopLeft</c> and the date label <c>BottomRight</c> — opposite sides
    /// of the line, because one 18pt row cannot hold two 10pt boxes on the same side — while the
    /// named style defaults to <c>TopRight</c>. The copy's tighter bounds are the observed trigger
    /// for a corner label being dropped, so this fact measures the copy against the live chart
    /// rather than asserting a rule the guide does not yet state.
    /// </remarks>
    [Fact]
    public void The_copy_places_explicit_delineator_corners_like_the_live_chart()
    {
        GanttEvent row = DelineatorEvent();

        GanttScene live = BuildLiveScene([row]);
        GanttScene export = BuildExportScene([row]);
        static PointD LineTopOrigin(GanttScene scene)
        {
            return Line(scene, _delineatorId).From;
        }

        static PointD LineBottomOrigin(GanttScene scene)
        {
            SceneLine line = Line(scene, _delineatorId);
            return new PointD(line.From.X, line.To.Y);
        }

        SceneText liveDescription = AssertSameLabelOffset(
            live, export, $"{_delineatorId.Value}:delineator-label", LineTopOrigin);
        _ = AssertSameLabelOffset(
            live, export, $"{_delineatorId.Value}:{DateLabelBuilder.StartRole}", LineBottomOrigin);

        Assert.True(
            liveDescription.TextBounds.Right <= LineTopOrigin(live).X + _tolerance,
            $"The live TopLeft description label is not left of its line: "
                + $"{liveDescription.TextBounds} vs x={LineTopOrigin(live).X}.");
    }

    /// <summary>
    /// Asserts the two scenes carry the same label primitive box RELATIVE to the shape it is
    /// anchored to, and returns the live primitive for follow-up assertions.
    /// </summary>
    /// <param name="live">The scene the live request built.</param>
    /// <param name="export">The scene the export request built.</param>
    /// <param name="primitiveId">The role-derived label identifier.</param>
    /// <param name="origin">Selects each scene's own anchor point (bar corner, line start).</param>
    /// <returns>The live scene's label primitive.</returns>
    private static SceneText AssertSameLabelOffset(
        GanttScene live,
        GanttScene export,
        string primitiveId,
        Func<GanttScene, PointD> origin)
    {
        SceneText liveLabel = Label(live, primitiveId);
        SceneText exportLabel = Label(export, primitiveId);

        Assert.Equal(liveLabel.Text, exportLabel.Text);
        Assert.Equal(liveLabel.Alignment, exportLabel.Alignment);
        Assert.Equal(
            OffsetFrom(origin(live), liveLabel.TextBounds),
            OffsetFrom(origin(export), exportLabel.TextBounds),
            new RectValueComparer());

        return liveLabel;
    }

    /// <summary>Returns the one label primitive the identifier names.</summary>
    /// <param name="scene">The scene to search.</param>
    /// <param name="primitiveId">The role-derived label identifier.</param>
    /// <returns>The label primitive.</returns>
    private static SceneText Label(GanttScene scene, string primitiveId) =>
        Assert.Single(
            scene.Primitives.OfType<SceneText>(),
            primitive => primitive.PrimitiveId == primitiveId);

    /// <summary>Returns the one span bar the row identifier owns.</summary>
    /// <param name="scene">The scene to search.</param>
    /// <param name="rowId">The row whose bar is wanted.</param>
    /// <returns>The bar's bounds.</returns>
    private static RectD Bar(GanttScene scene, GanttRowId rowId) =>
        Assert.Single(
            scene.Primitives.OfType<SceneRect>(),
            primitive => primitive.PrimitiveId == $"{rowId.Value}:bar").Bounds;

    /// <summary>Returns the one delineator line the row identifier owns.</summary>
    /// <param name="scene">The scene to search.</param>
    /// <param name="rowId">The delineator row.</param>
    /// <returns>The line primitive.</returns>
    private static SceneLine Line(GanttScene scene, GanttRowId rowId) =>
        Assert.Single(
            scene.Primitives.OfType<SceneLine>(),
            primitive => primitive.PrimitiveId == $"{rowId.Value}:delineator");

    /// <summary>Rebases a label box on its anchor point, so profiles are comparable.</summary>
    /// <param name="origin">The shape edge the label is anchored to.</param>
    /// <param name="bounds">The absolute label box.</param>
    /// <returns>The anchor-relative box.</returns>
    private static RectD OffsetFrom(PointD origin, RectD bounds) =>
        new(
            bounds.Left - origin.X,
            bounds.Top - origin.Y,
            bounds.Width,
            bounds.Height);

    /// <summary>The live composition the Refresh engine builds for the events.</summary>
    /// <param name="events">The events to compose.</param>
    /// <returns>The built live scene.</returns>
    private static GanttScene BuildLiveScene(IReadOnlyList<GanttEvent> events)
    {
        SceneBuildRequestOutcome request = new ExcelSceneBuildRequestFactory(new FakeTextMetrics()).Create(
            events,
            new Dictionary<string, string>(),
            Registry(),
            Grid(),
            new Dictionary<string, double>());

        Assert.True(request.Succeeded, "the live request was refused: " + request.Message);

        SceneBuildOutcome built = SceneBuilder.TryBuild(request.Request!);
        Assert.True(built.Succeeded, "the live scene was refused: " + built.Refusal);
        return built.Result!.Scene;
    }

    /// <summary>The editable-export composition the Copy Editable command builds.</summary>
    /// <param name="events">The events to compose.</param>
    /// <returns>The built export scene.</returns>
    private static GanttScene BuildExportScene(IReadOnlyList<GanttEvent> events)
    {
        SceneBuildRequestOutcome request = new ExcelSceneBuildRequestFactory(new FakeTextMetrics()).CreateExport(
            events,
            new Dictionary<string, string>(),
            Registry(),
            Grid(),
            new Dictionary<string, double>());

        Assert.True(request.Succeeded, "the export request was refused: " + request.Message);

        ExportBoundsOutcome resolved = ExportBoundsResolver.TryResolve(
            request.Request!,
            ExportPanelTheme.Resolve(),
            ExportSelection.TablePlusChart);

        Assert.True(resolved.Succeeded, "the export bounds were refused: " + resolved.Refusal);
        return resolved.Scene!;
    }

    /// <summary>A measured grid wide enough to leave a readable plot.</summary>
    /// <returns>The measured grid.</returns>
    private static PanelCellGrid Grid()
    {
        double rowHeight = GanttCatalogues.MetricDefault("GanttRowHeightPt");
        return PanelCellGrid.TryCreate(
            [new PanelColumn("Id", 120)],
            [rowHeight],
            rowHeight,
            ["Id"],
            originTopPt: 58).Grid!;
    }

    /// <summary>An As-Planned Activity whose explicit Right differs from its style default.</summary>
    /// <returns>The validated span event.</returns>
    /// <remarks>
    /// <c>Right</c> is chosen deliberately. It anchors to the bar's right edge, so its offset is
    /// invariant to the plot's scale, and the row asks for the START date label only: a finish
    /// date label would claim that same right-hand edge and one of the two would be suppressed.
    /// <c>Inside</c> is not used because a centred box re-centres when the plot width differs.
    /// </remarks>
    private static GanttEvent SpanEvent() =>
        Event(_spanId, GanttEntityType.AsPlannedActivity, "Design", start: _start, finish: _finish) with
        {
            StyleKey = "AsPlannedActivity",
            LabelPosition = GanttLabelPosition.Right,
            DateLabelPosition = GanttDateLabelPosition.Start,
        };

    /// <summary>
    /// A Delineator whose explicit corners differ from its style default. A delineator reads
    /// <c>Start</c> only, so no finish date is supplied. The description takes the top-left
    /// corner and the date label the bottom-right one, so the two are on opposite sides of the
    /// line and never collide (one 18pt row cannot hold two 10pt boxes on the same side); both
    /// differ from the style's top-right default.
    /// </summary>
    /// <returns>The validated delineator event.</returns>
    private static GanttEvent DelineatorEvent() =>
        Event(_delineatorId, GanttEntityType.Delineator, "Gate", start: _start, finish: null) with
        {
            StyleKey = "GateDelineator",
            LabelPosition = GanttLabelPosition.TopLeft,
            DateLabelPlacement = GanttLabelPosition.BottomRight,
        };

    /// <summary>Builds one event with the non-default members the fixtures need.</summary>
    /// <param name="id">The row identifier.</param>
    /// <param name="type">The entity type.</param>
    /// <param name="description">The row description.</param>
    /// <param name="start">The start date.</param>
    /// <param name="finish">The finish date, or null for a start-only type.</param>
    /// <returns>The event.</returns>
    private static GanttEvent Event(
        GanttRowId id,
        GanttEntityType type,
        string description,
        DateOnly start,
        DateOnly? finish) =>
        new(
            RowNumber: 1,
            Id: id,
            LaneId: GanttRowId.New(),
            StackIndex: 0,
            Type: type,
            Description: description,
            Start: start,
            Finish: finish,
            ParentId: null,
            StyleKey: null,
            LabelPosition: null,
            FillColour: null,
            StrokeColour: null,
            Visible: true,
            SortOrder: 0);

    /// <summary>
    /// A registry whose styles deliberately default to a DIFFERENT position than the rows name.
    /// </summary>
    /// <returns>The style registry.</returns>
    private static GanttStyleRegistry Registry() =>
        new([
            new GanttStyleDefinition(
                "AsPlannedActivity",
                new HashSet<GanttLabelPosition>
                {
                    GanttLabelPosition.Auto,
                    GanttLabelPosition.Left,
                    GanttLabelPosition.Right,
                    GanttLabelPosition.Inside,
                    GanttLabelPosition.None,
                },
                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                DefaultLabelPosition: GanttLabelPosition.Inside,
                FillColour: "#DDEBF7",
                StrokeColour: "#2E75B6",
                TextColour: "#000000",
                HatchPattern: GanttHatchPattern.None,
                StandardOutlinePt: 1,
                ActivityHeightPt: 12),
            new GanttStyleDefinition(
                "GateDelineator",
                new HashSet<GanttLabelPosition>
                {
                    GanttLabelPosition.Auto,
                    GanttLabelPosition.TopLeft,
                    GanttLabelPosition.TopRight,
                    GanttLabelPosition.BottomLeft,
                    GanttLabelPosition.BottomRight,
                    GanttLabelPosition.None,
                },
                EntityColourCapability.Stroke,
                DefaultLabelPosition: GanttLabelPosition.TopRight,
                StrokeColour: "#C00000",
                HatchPattern: GanttHatchPattern.None,
                StandardOutlinePt: 1),
        ]);

    /// <summary>Compares two boxes member by member within the label tolerance.</summary>
    private sealed class RectValueComparer : IEqualityComparer<RectD>
    {
        public bool Equals(RectD x, RectD y) =>
            Math.Abs(x.Left - y.Left) <= _tolerance
            && Math.Abs(x.Top - y.Top) <= _tolerance
            && Math.Abs(x.Width - y.Width) <= _tolerance
            && Math.Abs(x.Height - y.Height) <= _tolerance;

        public int GetHashCode(RectD obj) => HashCode.Combine(obj.Left, obj.Top, obj.Width, obj.Height);
    }
}
