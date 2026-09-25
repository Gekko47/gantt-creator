namespace GanttCreator.Core.Scene;

/// <summary>Injected deterministic text-width measurement seam for scene construction.</summary>
public interface ITextWidthMeasurer
{
    /// <summary>Attempts to measure text width in points.</summary>
    /// <param name="text">The exact text to measure.</param>
    /// <param name="widthPt">The measured width when successful.</param>
    /// <returns><see langword="true"/> when a finite non-negative width was measured.</returns>
    bool TryMeasure(string text, out double widthPt);
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

        RectD yearBounds = new(
            request.PlotBounds.X,
            request.PlotBounds.Y - request.YearBandHeightPt,
            request.PlotBounds.Width,
            request.YearBandHeightPt
        );
        RectD periodBounds = new(
            request.PlotBounds.X,
            yearBounds.Y - request.PeriodBandHeightPt,
            request.PlotBounds.Width,
            request.PeriodBandHeightPt
        );
        RectD contentBounds = Union(request.PanelBounds, yearBounds, periodBounds, request.PlotBounds);
        RectD? titleBounds = request.ShowTitle
            ? new RectD(contentBounds.X, contentBounds.Y - request.TitleBandHeightPt, contentBounds.Width, request.TitleBandHeightPt)
            : null;
        RectD unionBounds = titleBounds is { } title ? Union(contentBounds, title) : contentBounds;
        RectD chartBounds = new(
            unionBounds.X - request.ChartOuterPaddingPt,
            unionBounds.Y - request.ChartOuterPaddingPt,
            unionBounds.Width + (request.ChartOuterPaddingPt * 2),
            unionBounds.Height + (request.ChartOuterPaddingPt * 2)
        );
        ChartFrameGeometry geometry = new(chartBounds, contentBounds, titleBounds, yearBounds, periodBounds);
        List<ScenePrimitive> primitives =
        [
            new SceneRect("chart:background", SceneOwnerId.Chart, ZLayer.Background, chartBounds, request.Theme.Background),
        ];
        List<SceneWarning> warnings = [];
        AddBands(primitives, request, sequence);
        AddHeaders(primitives, request, sequence);
        AddTitle(primitives, warnings, request, geometry, textMeasurer);
        AddGrid(primitives, request, sequence);
        AddOuterFrame(primitives, request, chartBounds);
        if (sequence.Periods.Count > 0 && sequence.Periods.All(period => !period.ShowLabel))
        {
            warnings.Add(new SceneWarning(SceneOwnerId.Chart, _periodLabelsSuppressedCode, "All period labels were suppressed."));
        }

        return new FrameBandsCreationOutcome(new FrameBandsResult(geometry, primitives, warnings), null);
    }

    private static void AddBands(List<ScenePrimitive> primitives, FrameBandsRequest request, BandSequence sequence)
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
                    new RectD(period.Left, request.PlotBounds.Y, period.Width, request.PlotBounds.Height),
                    request.Theme.AlternateBand
                )
            );
        }
    }

    private static void AddHeaders(List<ScenePrimitive> primitives, FrameBandsRequest request, BandSequence sequence)
    {
        for (var index = 0; index < sequence.Years.Count; index++)
        {
            BandInterval year = sequence.Years[index];
            RectD bounds = new(year.Left, request.PlotBounds.Y - request.YearBandHeightPt, year.Width, request.YearBandHeightPt);
            primitives.Add(
                new SceneRect($"chart:year:{year.Start.Year}", SceneOwnerId.Chart, ZLayer.Frame, bounds, request.Theme.YearHeader)
            );
            if (year.ShowLabel)
            {
                primitives.Add(
                    new SceneText(
                        $"chart:year-label:{year.Start.Year}",
                        SceneOwnerId.Chart,
                        ZLayer.Frame,
                        year.Label,
                        bounds,
                        request.Theme.YearHeader,
                        GanttLabelPosition.Auto
                    )
                );
            }
        }

        for (var index = 0; index < sequence.Periods.Count; index++)
        {
            BandInterval period = sequence.Periods[index];
            var id = $"chart:period:{period.Start:yyyy-MM-dd}";
            RectD bounds = new(
                period.Left,
                request.PlotBounds.Y - request.YearBandHeightPt - request.PeriodBandHeightPt,
                period.Width,
                request.PeriodBandHeightPt
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
                        GanttLabelPosition.Auto
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
            new SceneText("chart:title-text", SceneOwnerId.Chart, ZLayer.Title, text, title, request.Theme.Title, GanttLabelPosition.Auto)
        );
    }

    private static void AddGrid(List<ScenePrimitive> primitives, FrameBandsRequest request, BandSequence sequence)
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
            primitives.Add(
                new SceneLine(
                    line.Key == request.PlotBounds.Left || line.Key == request.PlotBounds.Right
                        ? $"chart:grid:edge:{line.Key:R}"
                        : $"chart:grid:{line.Key:R}",
                    SceneOwnerId.Chart,
                    ZLayer.Grid,
                    new PointD(line.Key, request.PlotBounds.Y),
                    new PointD(line.Key, request.PlotBounds.Bottom),
                    line.Value ? request.Theme.MajorGrid : request.Theme.MinorGrid
                )
            );
        }
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
        && IsFiniteNonNegative(request.ChartOuterPaddingPt)
        && IsFinitePositive(request.TitleBandHeightPt)
        && IsFinitePositive(request.YearBandHeightPt)
        && IsFinitePositive(request.PeriodBandHeightPt)
        && IsFiniteNonNegative(request.MinimumHeaderLabelWidthPt)
        && IsFinitePositive(request.GridLinePt)
        && IsFinitePositive(request.MajorBoundaryPt);

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

    private static string Ellipsize(string text, double availableWidth, ITextWidthMeasurer measurer)
    {
        for (var length = text.Length - 1; length >= 0; length--)
        {
            var candidate = string.Concat(text.AsSpan(0, length), "…");
            if (measurer.TryMeasure(candidate, out var width) && width <= availableWidth)
            {
                return candidate;
            }
        }

        return "…";
    }

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
