namespace GanttCreator.Core.Scene;

/// <summary>One clipped year or scale-period interval.</summary>
/// <param name="Start">The inclusive visible start date.</param>
/// <param name="Finish">The inclusive visible finish date.</param>
/// <param name="Left">The exact left X coordinate.</param>
/// <param name="Right">The exact right X coordinate.</param>
/// <param name="Label">The invariant-culture label.</param>
/// <param name="ShowLabel">Whether the label clears the minimum visible width.</param>
public sealed record BandInterval(DateOnly Start, DateOnly Finish, double Left, double Right, string Label, bool ShowLabel)
{
    /// <summary>Gets the visible width in points.</summary>
    public double Width => Right - Left;
}

/// <summary>The resolved rectangle layout for the chart frame and headers.</summary>
/// <param name="ChartBounds">The final padded chart bounds.</param>
/// <param name="ContentBounds">The unpadded union of panel, headers, and plot.</param>
/// <param name="TitleBounds">The title rectangle when visible.</param>
/// <param name="YearBounds">The year-header strip.</param>
/// <param name="PeriodBounds">The period-header strip.</param>
public sealed record ChartFrameGeometry(RectD ChartBounds, RectD ContentBounds, RectD? TitleBounds, RectD YearBounds, RectD PeriodBounds);

/// <summary>Resolved scene styles consumed by the frame/band builder.</summary>
/// <param name="Background">The chart background style.</param>
/// <param name="AlternateBand">The alternate-band style.</param>
/// <param name="MinorGrid">The minor-grid style.</param>
/// <param name="MajorGrid">The major-grid/frame style.</param>
/// <param name="YearHeader">The year-header style.</param>
/// <param name="PeriodHeader">The period-header style.</param>
/// <param name="Title">The title style.</param>
public sealed record FrameBandsTheme(
    SceneStyle Background,
    SceneStyle AlternateBand,
    SceneStyle MinorGrid,
    SceneStyle MajorGrid,
    SceneStyle YearHeader,
    SceneStyle PeriodHeader,
    SceneStyle Title
);

/// <summary>Typed inputs for pure chart frame and band construction.</summary>
/// <param name="TimeScale">The validated date-to-X scale.</param>
/// <param name="Scale">The selected calendar scale.</param>
/// <param name="PeriodLabelFormat">The selected period label format.</param>
/// <param name="PanelBounds">The measured data-panel bounds supplied by the caller.</param>
/// <param name="PlotBounds">The measured plot bounds supplied by the caller.</param>
/// <param name="ChartOuterPaddingPt">Padding applied once around the union.</param>
/// <param name="TitleBandHeightPt">The title strip height.</param>
/// <param name="YearBandHeightPt">The year-header height.</param>
/// <param name="PeriodBandHeightPt">The period-header height.</param>
/// <param name="MinimumHeaderLabelWidthPt">Minimum visible label width.</param>
/// <param name="GridLinePt">Minor grid width.</param>
/// <param name="MajorBoundaryPt">Major boundary/frame width.</param>
/// <param name="ShowTitle">Whether the populated title band is visible.</param>
/// <param name="ChartTitle">The nonblank stored title when visible.</param>
/// <param name="AlternateBanding">Whether alternating plot bands are emitted.</param>
/// <param name="ShowMinorGrid">Whether minor period grid lines are emitted.</param>
/// <param name="ShowMajorGrid">Whether major year/plot grid lines are emitted.</param>
/// <param name="Theme">Resolved frame/band styles.</param>
public sealed record FrameBandsRequest(
    TimeScale TimeScale,
    GanttTimeScale Scale,
    GanttPeriodLabelFormat PeriodLabelFormat,
    RectD PanelBounds,
    RectD PlotBounds,
    double ChartOuterPaddingPt,
    double TitleBandHeightPt,
    double YearBandHeightPt,
    double PeriodBandHeightPt,
    double MinimumHeaderLabelWidthPt,
    double GridLinePt,
    double MajorBoundaryPt,
    bool ShowTitle,
    string ChartTitle,
    bool AlternateBanding,
    bool ShowMinorGrid,
    bool ShowMajorGrid,
    FrameBandsTheme Theme
);

/// <summary>The result of a successful frame/band build.</summary>
/// <param name="Geometry">The resolved chart geometry.</param>
/// <param name="Primitives">The chart-owned scene primitives.</param>
/// <param name="Warnings">Non-blocking chart warnings.</param>
public sealed record FrameBandsResult(
    ChartFrameGeometry Geometry,
    IReadOnlyList<ScenePrimitive> Primitives,
    IReadOnlyList<SceneWarning> Warnings
);

/// <summary>The reason a frame/band build was refused.</summary>
public enum FrameBandsRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The time scale was invalid.</summary>
    InvalidTimeScale = 1,

    /// <summary>The scale and period format were incompatible.</summary>
    IncompatibleSettings = 2,

    /// <summary>One or more geometry/metric values were invalid.</summary>
    InvalidGeometry = 3,

    /// <summary>The visible title was blank.</summary>
    BlankVisibleTitle = 4,

    /// <summary>The text measurement seam was absent or failed.</summary>
    TextMeasurementUnavailable = 5,

    /// <summary>One or more resolved styles were null.</summary>
    InvalidTheme = 6,
}

/// <summary>The typed result of attempting to build frame/bands.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record FrameBandsCreationOutcome(FrameBandsResult? Result, FrameBandsRefusal? Refusal)
{
    /// <summary>Gets whether construction succeeded.</summary>
    public bool Succeeded => Result is not null;
}
