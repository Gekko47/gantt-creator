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
/// <param name="Padding">
/// The margin between the content union and the chart frame, per side
/// (ADR-0031 D1). A live chart pads top and bottom by whole measured rows and left
/// by nothing, so this cannot be a single scalar.
/// </param>
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
/// <param name="PlotBandHeaderOverlapPt">
/// How far the plot-spanning shapes extend UP into the header row, from the
/// <c>PlotBandHeaderOverlapPt</c> token (ADR-0037 D1). Zero restores the prior
/// behaviour exactly.
/// </param>
/// <param name="ChartAnchorRowHeightPt">
/// The reserved anchor row's measured height, from the <c>ChartAnchorRowHeightPt</c>
/// token (ADR-0038 D1). It is how far the plot-spanning shapes extend DOWN through
/// that row and half the closing line, and where the closing line sits. Zero
/// reproduces the pre-ADR-0038 geometry exactly.
/// </param>
/// <param name="Theme">Resolved frame/band styles.</param>
/// <param name="SheetCoordinates">
/// Whether the request's coordinates are absolute worksheet points, which is
/// true of the LIVE composition and false of the EXPORT ones (ADR-0047).
/// <para>
/// A live chart's coordinates are the sheet's own, so nothing can be drawn
/// above row 1. Excel clamps a shape's negative top to the top of the sheet,
/// and a clamped top with an unchanged height moves the shape's BOTTOM down,
/// so the chart's background and right edge end below the row they belong to -
/// the sheet and the model then disagree about where the chart ends.
/// </para>
/// <para>
/// Two rules follow, and both are applied by the frame builder. The CONTENT
/// must be on the sheet: a band above row 1 cannot be drawn at all, so the
/// build is refused rather than silently misplaced. The FRAME's own top
/// margin may reach above the sheet, and is then clamped - the frame is
/// shortened at the top so that its bottom edge stays exactly where it
/// belongs.
/// </para>
/// <para>
/// The EXPORT compositions build in their own coordinate space and are
/// translated to a zero origin by the renderer, so a negative origin is
/// legitimate there and this flag is off.
/// </para>
/// </param>
public sealed record FrameBandsRequest(
    TimeScale TimeScale,
    GanttTimeScale Scale,
    GanttPeriodLabelFormat PeriodLabelFormat,
    RectD PanelBounds,
    RectD PlotBounds,
    ChartPaddingPt Padding,
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
    double PlotBandHeaderOverlapPt,
    double ChartAnchorRowHeightPt,
    FrameBandsTheme Theme,
    bool SheetCoordinates = false
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

    /// <summary>
    /// The chart's content would be drawn above the top of the sheet.
    /// </summary>
    /// <remarks>
    /// A live chart's coordinates are the sheet's own, so a band above row 1
    /// cannot be drawn where the model says. Excel clamps it instead, which
    /// moves the band down over the row it belongs to, so the build is
    /// refused rather than silently misplaced (ADR-0047).
    /// </remarks>
    ContentAboveSheet = 7,
}

/// <summary>The typed result of attempting to build frame/bands.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record FrameBandsCreationOutcome(FrameBandsResult? Result, FrameBandsRefusal? Refusal)
{
    /// <summary>Gets whether construction succeeded.</summary>
    public bool Succeeded => Result is not null;
}
