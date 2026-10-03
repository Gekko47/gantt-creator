using System.Globalization;

namespace GanttCreator.Core.Scene;

/// <summary>
/// A width-only view over the single <see cref="ITextMetrics"/> seam, kept so the
/// R3.5 frame/band builder reads naturally without measuring a label box it
/// never uses.
/// </summary>
/// <remarks>
/// <c>docs/02-ARCHITECTURE.md</c> allows exactly one text-metrics service, so
/// this is an adapter over <see cref="ITextMetrics"/>, not a second measuring
/// implementation. R3.6 widened the port to return a height as well as a width.
/// </remarks>
public interface ITextWidthMeasurer
{
    /// <summary>Attempts to measure text width in points.</summary>
    /// <param name="text">The exact text to measure.</param>
    /// <param name="widthPt">The measured width when successful.</param>
    /// <returns><see langword="true"/> when a finite non-negative width was measured.</returns>
    bool TryMeasure(string text, out double widthPt);
}

/// <summary>
/// Presents an <see cref="ITextWidthMeasurer"/> as an <see cref="ITextMetrics"/>
/// so the width-only seam can reach the shared truncation helper.
/// </summary>
/// <param name="measurer">The width-only seam to adapt.</param>
/// <remarks>
/// The height is reported as zero because no caller of the shared truncation
/// helper needs it: the helper compares widths only.
/// </remarks>
public sealed class WidthOnlyMetrics(ITextWidthMeasurer measurer) : ITextMetrics
{
    private readonly ITextWidthMeasurer _measurer =
        measurer ?? throw new ArgumentNullException(nameof(measurer));

    /// <inheritdoc />
    public bool TryMeasure(string text, out TextMeasurement? measurement)
    {
        if (_measurer.TryMeasure(text, out var widthPt))
        {
            measurement = new TextMeasurement(widthPt, 0);
            return true;
        }

        measurement = null;
        return false;
    }
}

/// <summary>Adapts an <see cref="ITextMetrics"/> to the width-only seam.</summary>
/// <param name="metrics">The injected deterministic text-metrics seam.</param>
public sealed class TextWidthMeasurer(ITextMetrics metrics) : ITextWidthMeasurer
{
    private readonly ITextMetrics _metrics =
        metrics ?? throw new ArgumentNullException(nameof(metrics));

    /// <inheritdoc />
    public bool TryMeasure(string text, out double widthPt)
    {
        if (_metrics.TryMeasure(text, out TextMeasurement? measurement) && measurement is not null)
        {
            widthPt = measurement.WidthPt;
            return true;
        }

        widthPt = 0;
        return false;
    }
}

/// <summary>Builds chart-owned frame, header, band, title, and grid primitives.</summary>
public static class FrameBandsBuilder
{
    private const string _titleOverflowCode = "TitleOverflow";
    private const string _periodLabelsSuppressedCode = "AllPeriodLabelsSuppressed";

    /// <summary>Attempts to build the frame and time bands.</summary>
    /// <param name="request">The typed frame/band request.</param>
    /// <param name="textMeasurer">The injected deterministic text-width seam.</param>
    /// <returns>A typed result or refusal.</returns>
    public static FrameBandsCreationOutcome TryBuild(FrameBandsRequest? request, ITextWidthMeasurer? textMeasurer)
    {
        if (request is null)
        {
            return Refused(FrameBandsRefusal.NullRequest);
        }

        if (request.TimeScale is null)
        {
            return Refused(FrameBandsRefusal.InvalidTimeScale);
        }

        if (!GanttChartSettings.IsCompatible(request.Scale, request.PeriodLabelFormat))
        {
            return Refused(FrameBandsRefusal.IncompatibleSettings);
        }

        if (!ValidGeometry(request))
        {
            return Refused(FrameBandsRefusal.InvalidGeometry);
        }

        if (request.ShowTitle && string.IsNullOrWhiteSpace(request.ChartTitle))
        {
            return Refused(FrameBandsRefusal.BlankVisibleTitle);
        }

        if (textMeasurer is null)
        {
            return Refused(FrameBandsRefusal.TextMeasurementUnavailable);
        }

        if (request.Theme is null || !ValidTheme(request.Theme))
        {
            return Refused(FrameBandsRefusal.InvalidTheme);
        }

        if (request.ShowTitle && !textMeasurer.TryMeasure(request.ChartTitle, out _))
        {
            return Refused(FrameBandsRefusal.TextMeasurementUnavailable);
        }

        BandSequenceCreationOutcome sequenceOutcome = BandSequence.TryCreate(
            request.TimeScale,
            request.Scale,
            request.PeriodLabelFormat,
            request.PlotBounds,
            request.MinimumHeaderLabelWidthPt
        );
        if (sequenceOutcome.Sequence is not { } sequence)
        {
            return Refused(
                sequenceOutcome.Refusal == BandSequenceRefusal.IncompatibleSettings
                    ? FrameBandsRefusal.IncompatibleSettings
                    : FrameBandsRefusal.InvalidGeometry
            );
        }

        // Entity guide §5/§6: the YEAR band is the upper strip and the PERIOD band sits
        // BELOW it, immediately above the plot. Stacked upward from the plot's top
        // edge that is: period first, then year, then (for export only) the title.
        //
        // These were previously built the other way round - year directly above the
        // plot, period above that - so a live chart drew the year band over the
        // header row and pushed the period band onto the reserved title row. It was
        // invisible in the fixtures because the golden builds from Core with its own
        // band heights, and it only appeared once the bands were finally positioned
        // against real measured rows.
        RectD periodBounds = new(
            request.PlotBounds.X,
            request.PlotBounds.Y - request.PeriodBandHeightPt,
            request.PlotBounds.Width,
            request.PeriodBandHeightPt
        );
        RectD yearBounds = new(
            request.PlotBounds.X,
            periodBounds.Y - request.YearBandHeightPt,
            request.PlotBounds.Width,
            request.YearBandHeightPt
        );
        RectD contentBounds = Union(request.PanelBounds, yearBounds, periodBounds, request.PlotBounds);
        RectD? titleBounds = request.ShowTitle
            ? new RectD(contentBounds.X, contentBounds.Y - request.TitleBandHeightPt, contentBounds.Width, request.TitleBandHeightPt)
            : null;
        RectD unionBounds = titleBounds is { } title ? Union(contentBounds, title) : contentBounds;
        // Each side is padded independently (ADR-0031 D1). A live chart's top and
        // bottom margins are whole measured rows while its left margin is zero, so
        // the plot sits flush against the data table instead of being pushed one
        // margin-width away from it. Export passes ChartPaddingPt.Uniform and is
        // unaffected.
        ChartPaddingPt padding = request.Padding;
        RectD chartBounds = new(
            unionBounds.X - padding.LeftPt,
            unionBounds.Y - padding.TopPt,
            unionBounds.Width + padding.LeftPt + padding.RightPt,
            unionBounds.Height + padding.TopPt + padding.BottomPt
        );
        ChartFrameGeometry geometry = new(chartBounds, contentBounds, titleBounds, yearBounds, periodBounds);

        // ADR-0037 D1: the vertical extent every PLOT-SPANNING shape shares --
        // background, alternate bands, vertical grid lines.
        //
        // Its top is lifted PlotBandHeaderOverlapPt ABOVE the plot's own top edge,
        // into the header row, and its height grows by the same amount so the bottom
        // is unchanged. Excel resizes a shape on a row insert only when the insertion
        // point is strictly below the shape's TopLeftCell row; with the top exactly on
        // the header/body boundary, a row added at the TOP of the body left the shape
        // sliding instead of stretching (measured 2026-10-03,
        // scripts/probe-frame-stretch.ps1) and the plot stayed unpainted there. Half a
        // point resolves TopLeftCell to the HEADER row, so every body insert is now
        // strictly below the anchor. The header paints at ZLayer.Frame, over these, so
        // the overlap is invisible; and export, which has no rows to anchor to, is
        // unaffected because the geometry it reads is unchanged.
        //
        // Zero reproduces the previous behaviour exactly, so this is a retunable
        // mechanism rather than an irreversible one.
        RectD plotSpanBounds = PlotSpanGeometry.LiftTopIntoHeader(request.PlotBounds, request.PlotBandHeaderOverlapPt);

        // ADR-0038 D2: the vertical extent every PLOT-SPANNING shape shares --
        // background, alternate bands, vertical grid lines -- now also extends DOWN
        // through the reserved anchor row and half the closing line.
        //
        // The top is lifted (ADR-0037 D1) because a body insert at or above a shape's
        // TopLeftCell row makes Excel SLIDE it; the bottom is extended because every
        // insert the add-in performs targets one row past the body, so the bottom edge
        // needs a cell anchor strictly below that point or the same shape slides
        // instead of stretching. Measured 2026-10-03
        // (scripts/probe-anchor-row-height.ps1 Q2): an insert at the anchor row gives
        // HeightDelta = 18, a full body row.
        //
        // The two are NOT symmetric, and this is the asymmetry: the top is a sub-row
        // overlap into the header, the bottom is a reserved row of its own. "Symmetrising"
        // them -- mirroring one onto the other -- re-breaks the bottom, because no
        // overlap can create a row for an edge to anchor into.
        //
        // Zero reproduces the previous behaviour exactly, so this is retunable rather
        // than baked in.
        plotSpanBounds = PlotSpanGeometry.ExtendBottomIntoAnchorRow(
            plotSpanBounds,
            request.ChartAnchorRowHeightPt,
            request.MajorBoundaryPt);

        List<ScenePrimitive> primitives =
        [
            // The background stays chartBounds, NOT plotSpanBounds. It spans the whole
            // chart including the title and panel, so its top already sits above the
            // header row and every body insert is strictly below its anchor -- it needs
            // no overlap, and shrinking it to the plot would leave the headers unpainted.
            new SceneRect("chart:background", SceneOwnerId.Chart, ZLayer.Background, chartBounds, request.Theme.Background),
        ];
        List<SceneWarning> warnings = [];
        AddBands(primitives, request, sequence, plotSpanBounds);
        AddHeaders(primitives, request, sequence, geometry);
        AddTitle(primitives, warnings, request, geometry, textMeasurer);
        AddGrid(primitives, request, sequence, plotSpanBounds);
        AddClosingLine(primitives, request);
        AddOuterFrame(primitives, request, chartBounds);
        if (sequence.Periods.Count > 0 && sequence.Periods.All(period => !period.ShowLabel))
        {
            warnings.Add(new SceneWarning(SceneOwnerId.Chart, _periodLabelsSuppressedCode, "All period labels were suppressed."));
        }

        return new FrameBandsCreationOutcome(new FrameBandsResult(geometry, primitives, warnings), null);
    }

    private static void AddBands(
        List<ScenePrimitive> primitives,
        FrameBandsRequest request,
        BandSequence sequence,
        RectD plotSpan)
    {
        if (!request.AlternateBanding)
        {
            return;
        }

        for (var index = 1; index < sequence.Periods.Count; index += 2)
        {
            BandInterval period = sequence.Periods[index];
            primitives.Add(
                new SceneRect(
                    $"chart:band:{index}",
                    SceneOwnerId.Chart,
                    ZLayer.AlternateBand,
                    new RectD(period.Left, plotSpan.Y, period.Width, plotSpan.Height),
                    request.Theme.AlternateBand
                )
            );
        }
    }

    /// <summary>
    /// Emits the year and period header bands, taking each band's VERTICAL extent
    /// from the already-resolved <see cref="ChartFrameGeometry"/>.
    /// </summary>
    /// <param name="primitives">The primitive list to append to.</param>
    /// <param name="request">The frame/band request.</param>
    /// <param name="sequence">The resolved band sequence.</param>
    /// <param name="geometry">
    /// The resolved geometry, whose <see cref="ChartFrameGeometry.YearBounds"/> and
    /// <see cref="ChartFrameGeometry.PeriodBounds"/> are the single authority for
    /// where each band sits.
    /// </param>
    /// <remarks>
    /// <para>
    /// This method previously recomputed both bands from
    /// <c>PlotBounds.Y</c> and the two band heights, and it computed them in the
    /// OPPOSITE order to <c>TryBuild</c>: it put the year band directly above the
    /// plot and the period band above that, while <c>TryBuild</c> puts the period
    /// band directly above the plot and the year band above it. So the
    /// <see cref="ChartFrameGeometry"/> this method's own caller returned described
    /// one arrangement and the primitives it emitted drew the other.
    /// </para>
    /// <para>
    /// The live symptom was entity guide sections 5 and 6 inverted on screen - the
    /// month labels sat in the upper row and the year in the lower one - and nothing
    /// caught it because the two computations each looked locally correct. It was
    /// invisible in the golden fixtures, which build from Core with their own band
    /// heights, and appeared only once the bands were positioned against real
    /// measured rows.
    /// </para>
    /// <para>
    /// Consuming the geometry makes the two incapable of disagreeing. The horizontal
    /// extents still come from the sequence, because each band is one rectangle PER
    /// year or per period; only the shared vertical placement is taken from here.
    /// </para>
    /// </remarks>
    private static void AddHeaders(
        List<ScenePrimitive> primitives,
        FrameBandsRequest request,
        BandSequence sequence,
        ChartFrameGeometry geometry)
    {
        for (var index = 0; index < sequence.Years.Count; index++)
        {
            BandInterval year = sequence.Years[index];
            var id = $"chart:year:{year.Start.Year}";
            RectD bounds = new(year.Left, geometry.YearBounds.Y, year.Width, geometry.YearBounds.Height);
            primitives.Add(
                new SceneRect(id, SceneOwnerId.Chart, ZLayer.Frame, bounds, request.Theme.YearHeader)
            );
            if (year.ShowLabel)
            {
                primitives.Add(
                    new SceneText(
                        // A header label names its parent by appending ":label" to
                        // the parent's identifier, the same convention the period
                        // header below and the row description label already use.
                        // The year header previously used a "chart:year-label:"
                        // prefix, so the two header kinds disagreed about how a
                        // child names its parent - and a renderer reconciles on
                        // exactly this text (R3.17).
                        $"{id}:label",
                        SceneOwnerId.Chart,
                        ZLayer.Frame,
                        year.Label,
                        bounds,
                        request.Theme.YearHeader,
                        GanttTextAlignment.Centre
                    )
                );
            }
        }

        for (var index = 0; index < sequence.Periods.Count; index++)
        {
            BandInterval period = sequence.Periods[index];
            var id = $"chart:period:{period.Start.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)}";
            RectD bounds = new(
                period.Left,
                geometry.PeriodBounds.Y,
                period.Width,
                geometry.PeriodBounds.Height
            );
            primitives.Add(new SceneRect(id, SceneOwnerId.Chart, ZLayer.Frame, bounds, request.Theme.PeriodHeader));
            if (period.ShowLabel)
            {
                primitives.Add(
                    new SceneText(
                        $"{id}:label",
                        SceneOwnerId.Chart,
                        ZLayer.Frame,
                        period.Label,
                        bounds,
                        request.Theme.PeriodHeader,
                        GanttTextAlignment.Centre
                    )
                );
            }
        }
    }

    private static void AddTitle(
        List<ScenePrimitive> primitives,
        List<SceneWarning> warnings,
        FrameBandsRequest request,
        ChartFrameGeometry geometry,
        ITextWidthMeasurer textMeasurer
    )
    {
        if (geometry.TitleBounds is not { } title)
        {
            return;
        }

        var text = request.ChartTitle;
        if (textMeasurer.TryMeasure(text, out var width) && width > title.Width)
        {
            warnings.Add(new SceneWarning(SceneOwnerId.Chart, _titleOverflowCode, "Chart title was truncated with an ellipsis."));
            text = Ellipsize(text, title.Width, textMeasurer);
        }

        primitives.Add(new SceneRect("chart:title-band", SceneOwnerId.Chart, ZLayer.Title, title, request.Theme.Title));
        primitives.Add(
            new SceneText("chart:title-text", SceneOwnerId.Chart, ZLayer.Title, text, title, request.Theme.Title, GanttTextAlignment.Centre)
        );
    }

    private static void AddGrid(
        List<ScenePrimitive> primitives,
        FrameBandsRequest request,
        BandSequence sequence,
        RectD plotSpan)
    {
        Dictionary<double, bool> lines = [];
        if (request.ShowMajorGrid)
        {
            AddGridLine(lines, request.PlotBounds.Left, true);
            AddGridLine(lines, request.PlotBounds.Right, true);
        }

        foreach (BandInterval period in sequence.Periods)
        {
            if (period.Right < request.PlotBounds.Right - GeometryMath.Epsilon)
            {
                var major = request.ShowMajorGrid && period.Finish.Month == 12;
                if (request.ShowMinorGrid || major)
                {
                    AddGridLine(lines, period.Right, major);
                }
            }
        }

        foreach (KeyValuePair<double, bool> line in lines.OrderBy(pair => pair.Key))
        {
            var key = line.Key.ToString("R", CultureInfo.InvariantCulture);
            primitives.Add(
                new SceneLine(
                    line.Key == request.PlotBounds.Left || line.Key == request.PlotBounds.Right
                        ? $"chart:grid:edge:{key}"
                        : $"chart:grid:{key}",
                    SceneOwnerId.Chart,
                    ZLayer.Grid,
                    new PointD(line.Key, plotSpan.Y),
                    new PointD(line.Key, plotSpan.Bottom),
                    line.Value ? request.Theme.MajorGrid : request.Theme.MinorGrid
                )
            );
        }
    }

    /// <summary>
    /// Emits the closing line at the plot's data boundary (ADR-0038 D3).
    /// </summary>
    /// <param name="primitives">The primitive list to append to.</param>
    /// <param name="request">The frame/band request.</param>
    /// <remarks>
    /// <para>
    /// <b>It is not the bottom frame.</b> <c>chart:frame:bottom</c> closes the chart at
    /// the bottom of its padding row and stays exactly where ADR-0031 D2 put it; this
    /// line closes the plot at the boundary of the data. Two lines are the intent, not
    /// an artefact of two code paths writing to the same edge: the rejected
    /// alternative was a single frame line at <c>plotBottom + extension</c>, which
    /// would have made the chart's bottom margin 0.75pt instead of the padding row's
    /// 5.75pt and broken ADR-0031 D2 in the process.
    /// </para>
    /// <para>
    /// <b>Its job is also to cover the overhang.</b> The bands now reach
    /// <c>plotBottom + anchor + w/2</c>, which is the line's own centre, so the band
    /// bottom lands under the stroke rather than beside it. <see cref="ZLayer.Frame"/>
    /// (80) puts it above the bands (10) and the grid (20), so the join reads as one
    /// terminated stack instead of bands fading out mid-cell.
    /// </para>
    /// </remarks>
    private static void AddClosingLine(List<ScenePrimitive> primitives, FrameBandsRequest request)
    {
        // With no anchor row there is no overhang and nothing to terminate, so no line
        // is emitted at all. Emitting one would draw a stroke across the plot's own
        // bottom edge with no band behind it -- a visible rule that closes nothing.
        if (!PlotSpanGeometry.HasAnchorRow(request.ChartAnchorRowHeightPt))
        {
            return;
        }

        var closingPt = PlotSpanGeometry.ClosingLineBottomPt(
            request.PlotBounds.Bottom,
            request.ChartAnchorRowHeightPt,
            request.MajorBoundaryPt);

        primitives.Add(
            new SceneLine(
                "chart:plot-closing",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new PointD(request.PlotBounds.Left, closingPt),
                new PointD(request.PlotBounds.Right, closingPt),
                request.Theme.MajorGrid
            )
        );
    }

    private static void AddGridLine(Dictionary<double, bool> lines, double x, bool major)
    {
        if (lines.TryGetValue(x, out var existingMajor))
        {
            lines[x] = existingMajor || major;
        }
        else
        {
            lines.Add(x, major);
        }
    }

    private static void AddOuterFrame(List<ScenePrimitive> primitives, FrameBandsRequest request, RectD bounds)
    {
        SceneStyle style = request.Theme.MajorGrid;
        primitives.Add(
            new SceneLine(
                "chart:frame:top",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new PointD(bounds.Left, bounds.Top),
                new PointD(bounds.Right, bounds.Top),
                style
            )
        );
        primitives.Add(
            new SceneLine(
                "chart:frame:bottom",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new PointD(bounds.Left, bounds.Bottom),
                new PointD(bounds.Right, bounds.Bottom),
                style
            )
        );
        primitives.Add(
            new SceneLine(
                "chart:frame:left",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new PointD(bounds.Left, bounds.Top),
                new PointD(bounds.Left, bounds.Bottom),
                style
            )
        );
        primitives.Add(
            new SceneLine(
                "chart:frame:right",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new PointD(bounds.Right, bounds.Top),
                new PointD(bounds.Right, bounds.Bottom),
                style
            )
        );
    }

    private static bool ValidGeometry(FrameBandsRequest request) =>
        request.PanelBounds.Width > 0
        && request.PanelBounds.Height > 0
        && request.PlotBounds.Width > 0
        && request.PlotBounds.Height > 0
        && GeometryMath.ApproximatelyEqual(request.TimeScale.PlotLeftPt, request.PlotBounds.Left)
        && GeometryMath.ApproximatelyEqual(request.TimeScale.PlotRightPt, request.PlotBounds.Right)
        && IsValidPadding(request.Padding)
        && IsFinitePositive(request.TitleBandHeightPt)
        && IsFinitePositive(request.YearBandHeightPt)
        && IsFinitePositive(request.PeriodBandHeightPt)
        && IsFiniteNonNegative(request.MinimumHeaderLabelWidthPt)
        && IsFiniteNonNegative(request.PlotBandHeaderOverlapPt)
        && IsFiniteNonNegative(request.ChartAnchorRowHeightPt)
        && IsFinitePositive(request.GridLinePt)
        && IsFinitePositive(request.MajorBoundaryPt);

    /// <summary>
    /// Every side of the padding must be a finite, non-negative number.
    /// </summary>
    /// <remarks>
    /// <b>All four sides are checked, not the maximum.</b> ADR-0031 D1 made the
    /// padding per-side, and a check that only looked at one member would accept a
    /// request whose <c>BottomPt</c> was negative while its other three were sound -
    /// producing chart bounds that extended <em>above</em> their own content, a
    /// rectangle no caller asked for and nothing downstream would report.
    /// </remarks>
    private static bool IsValidPadding(ChartPaddingPt padding) =>
        IsFiniteNonNegative(padding.LeftPt)
        && IsFiniteNonNegative(padding.TopPt)
        && IsFiniteNonNegative(padding.RightPt)
        && IsFiniteNonNegative(padding.BottomPt);

    private static bool ValidTheme(FrameBandsTheme theme) =>
        theme.Background is not null
        && theme.AlternateBand is not null
        && theme.MinorGrid is not null
        && theme.MajorGrid is not null
        && theme.YearHeader is not null
        && theme.PeriodHeader is not null
        && theme.Title is not null;

    private static bool IsFiniteNonNegative(double value) => double.IsFinite(value) && value >= 0;

    private static bool IsFinitePositive(double value) => double.IsFinite(value) && value > 0;

    /// <summary>
    /// Truncates the title to the available width through the shared
    /// <see cref="LabelText"/> helper, so the chart title and the R3.6 description
    /// labels truncate by one implementation and one marker.
    /// </summary>
    private static string Ellipsize(string text, double availableWidth, ITextWidthMeasurer measurer) =>
        measurer.TryMeasure(text, out var width) && width <= availableWidth
            ? text
            : LabelText.Ellipsize(text, availableWidth, new WidthOnlyMetrics(measurer));

    private static RectD Union(params RectD[] rectangles)
    {
        var left = rectangles.Min(rectangle => rectangle.Left);
        var top = rectangles.Min(rectangle => rectangle.Top);
        var right = rectangles.Max(rectangle => rectangle.Right);
        var bottom = rectangles.Max(rectangle => rectangle.Bottom);
        return new RectD(left, top, right - left, bottom - top);
    }

    private static FrameBandsCreationOutcome Refused(FrameBandsRefusal refusal) => new(null, refusal);
}

/// <summary>
/// Live-anchoring arithmetic for the plot-spanning shapes (ADR-0037 D1).
/// </summary>
public static class PlotSpanGeometry
{
    /// <summary>
    /// Lifts a plot rectangle's top edge <paramref name="overlapPt"/> into the header
    /// row and grows its height by the same amount, leaving its bottom unchanged.
    /// </summary>
    /// <param name="plot">The plot rectangle as measured.</param>
    /// <param name="overlapPt">The sub-row overlap, from the catalogue token.</param>
    /// <returns>The adjusted rectangle; unchanged when the overlap is zero.</returns>
    /// <remarks>
    /// <para>
    /// <b>Why the top edge moves up and not the bottom edge down.</b> The bottom is
    /// the plot's own boundary and several entity contracts are stated against it; the
    /// top is the only edge with no such meaning, so it is the only one that can absorb
    /// the overlap without moving anything the product has an opinion about.
    /// </para>
    /// <para>
    /// <b>Why not the row above.</b> A whole-row overlap would work too, but it would
    /// paint over whatever the user replaced that row with. Half a point is enough to
    /// resolve the shape's cell anchor to the header and cannot reach visibly.
    /// </para>
    /// </remarks>
    public static RectD LiftTopIntoHeader(RectD plot, double overlapPt)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(overlapPt);
        return overlapPt == 0
            ? plot
            : new RectD(plot.X, plot.Y - overlapPt, plot.Width, plot.Height + overlapPt);
    }

    /// <summary>
    /// Extends a plot rectangle's bottom edge down through the reserved anchor row and
    /// half the closing line, growing its height and leaving its top unchanged
    /// (ADR-0038 D2).
    /// </summary>
    /// <param name="plot">The plot rectangle as measured.</param>
    /// <param name="anchorRowHeightPt">
    /// The reserved anchor row's measured height, from the <c>ChartAnchorRowHeightPt</c>
    /// token. Zero reproduces the pre-ADR-0038 geometry exactly.
    /// </param>
    /// <param name="majorBoundaryPt">The closing line's width, from <c>MajorBoundaryPt</c>.</param>
    /// <returns>The adjusted rectangle.</returns>
    /// <remarks>
    /// <para>
    /// <b>This is NOT the mirror of <see cref="LiftTopIntoHeader"/>, and must not be
    /// written as one.</b> Excel resizes a shape on a row insert only when the
    /// insertion point is strictly below the relevant cell anchor, and the two ends
    /// behave differently: an insert at a shape's <c>TopLeftCell</c> row SLIDES it
    /// (<c>TopDelta</c> +18, <c>HeightDelta</c> 0), while an insert at its
    /// <c>BottomRightCell</c> row STRETCHES it (<c>HeightDelta</c> +18). Measured
    /// 2026-10-03, <c>scripts/probe-anchor-row-height.ps1</c> and
    /// <c>scripts/probe-frame-stretch.ps1</c>.
    /// </para>
    /// <para>
    /// <b>Why the top could be an overlap and the bottom must be a row.</b> Every
    /// insert the add-in performs targets one row past the body, so the bottom edge
    /// needs a cell anchor strictly BELOW that point, and only a reserved row can
    /// create one -- lifting the edge alone cannot. That is the asymmetry this method
    /// exists to encode, and the reason a "symmetrise the two edges" cleanup re-breaks
    /// the bottom.
    /// </para>
    /// <para>
    /// <b>Why half the line is included in the extension.</b> The closing line
    /// (D3) is centred on the boundary with width <paramref name="majorBoundaryPt"/>,
    /// so it covers <c>boundary - w/2 .. boundary + w/2</c>. The bands must reach that
    /// boundary -- and stop there -- so the extension is the anchor row plus half a
    /// line. At the defaults that is 0.25 + 0.5 = <b>0.75pt</b>, the line covers
    /// <c>plotBottom + 0.25 .. plotBottom + 1.25</c>, and the band bottom lands on
    /// <c>plotBottom + 0.75</c>: covered, with nothing spare.
    /// </para>
    /// </remarks>
    public static RectD ExtendBottomIntoAnchorRow(
        RectD plot,
        double anchorRowHeightPt,
        double majorBoundaryPt)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(anchorRowHeightPt);
        ArgumentOutOfRangeException.ThrowIfNegative(majorBoundaryPt);

        var extensionPt = BottomExtensionPt(anchorRowHeightPt, majorBoundaryPt);
        return extensionPt == 0
            ? plot
            : new RectD(plot.X, plot.Y, plot.Width, plot.Height + extensionPt);
    }

    /// <summary>
    /// Whether this composition has an anchor row, and therefore an overhang and a
    /// closing line at all (ADR-0038 D1).
    /// </summary>
    /// <param name="anchorRowHeightPt">The reserved anchor row's measured height.</param>
    /// <returns><see langword="true"/> when the plot bottom is extended.</returns>
    /// <remarks>
    /// <para>
    /// <b>Why this is a predicate and not merely an arithmetic outcome.</b> The
    /// extension is <c>anchor + w/2</c>, so with an anchor row of <b>zero</b> it would
    /// still be half a line -- 0.5pt -- and the token's "zero restores the previous
    /// geometry exactly" claim would be false by half a point. The closing line exists
    /// only to terminate stacks that overhang into the anchor row; with no anchor row
    /// there is nothing to overhang and nothing to terminate, so the whole mechanism
    /// switches off rather than leaving half of itself applied.
    /// </para>
    /// <para>
    /// The alternative -- treating half a line as a legitimate zero-anchor extension --
    /// would make the retunable claim untrue and leave a closing line drawn across the
    /// plot's own bottom edge with no band under it.
    /// </para>
    /// </remarks>
    public static bool HasAnchorRow(double anchorRowHeightPt)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(anchorRowHeightPt);
        return anchorRowHeightPt > 0;
    }

    /// <summary>
    /// The Y coordinate of the closing line: the plot's bottom edge plus the anchor
    /// row and half the line (ADR-0038 D3).
    /// </summary>
    /// <param name="plotBottomPt">The plot's own bottom edge.</param>
    /// <param name="anchorRowHeightPt">The reserved anchor row's measured height.</param>
    /// <param name="majorBoundaryPt">The closing line's width.</param>
    /// <returns>The closing line's centre Y.</returns>
    /// <remarks>
    /// <para>
    /// <b>The same arithmetic as the extension, read once.</b> The line's centre must
    /// equal the bands' bottom or the overhang is either uncovered or painted past the
    /// line, and two independent expressions of "plotBottom + anchor + w/2" are two
    /// chances to disagree. The extension and this coordinate therefore both call
    /// <see cref="BottomExtensionPt"/>.
    /// </para>
    /// <para>
    /// <b>Coverage is arithmetic, never measurement.</b> An Excel <c>AddLine</c> has a
    /// degenerate bounding box -- its <c>Top + Height</c> is a point, not the stroke --
    /// so a test that measures whether the band is "covered" by reading the line's
    /// bounds reports <c>False</c> for a correctly drawn chart. Coverage comes from
    /// <c>Line.Weight</c>, which is why the assertion is this subtraction.
    /// </para>
    /// </remarks>
    public static double ClosingLineBottomPt(
        double plotBottomPt,
        double anchorRowHeightPt,
        double majorBoundaryPt) =>
        plotBottomPt + BottomExtensionPt(anchorRowHeightPt, majorBoundaryPt);

    /// <summary>
    /// How far the plot-spanning shapes extend below the plot's bottom edge: the anchor
    /// row plus half the closing line.
    /// </summary>
    /// <param name="anchorRowHeightPt">The reserved anchor row's measured height.</param>
    /// <param name="majorBoundaryPt">The closing line's width.</param>
    /// <returns>The extension in points.</returns>
    /// <remarks>
    /// Private because both callers above are the contract. Publishing it would invite
    /// a fourth expression of the same sum to be written somewhere it can drift.
    /// </remarks>
    private static double BottomExtensionPt(double anchorRowHeightPt, double majorBoundaryPt) =>
        HasAnchorRow(anchorRowHeightPt) ? anchorRowHeightPt + (majorBoundaryPt / 2) : 0d;
}
