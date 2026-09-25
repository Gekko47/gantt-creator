namespace GanttCreator.Core.Scene;

/// <summary>Resolved metric values used by lane layout.</summary>
/// <param name="LaneHeightPt">The minimum lane height.</param>
/// <param name="LanePaddingTopPt">Space above the first stack slot.</param>
/// <param name="LanePaddingBottomPt">Space below the last stack slot.</param>
/// <param name="StackGapPt">Space between distinct effective slots.</param>
/// <param name="SplitterHeightPt">The fixed splitter lane height.</param>
/// <param name="SpacerHeightPt">The fixed spacer lane height.</param>
public sealed record LaneLayoutMetrics(
    double LaneHeightPt,
    double LanePaddingTopPt,
    double LanePaddingBottomPt,
    double StackGapPt,
    double SplitterHeightPt,
    double SpacerHeightPt
);

/// <summary>One deterministic visual lane and its vertical slots.</summary>
/// <param name="LaneKey">The stable lane key used by this layout.</param>
/// <param name="LaneOrder">The deterministic lane order.</param>
/// <param name="Top">The lane top in points.</param>
/// <param name="Height">The lane height in points.</param>
/// <param name="Slots">The ordered vertical slots.</param>
/// <param name="EventIds">The stable event IDs in the lane.</param>
/// <param name="IsSplitter">Whether this is a splitter lane.</param>
/// <param name="IsSpacer">Whether this is a spacer lane.</param>
public sealed record LaneGeometry(
    string LaneKey,
    int LaneOrder,
    double Top,
    double Height,
    IReadOnlyList<SlotGeometry> Slots,
    IReadOnlyList<GanttRowId> EventIds,
    bool IsSplitter,
    bool IsSpacer
);

/// <summary>The result of a successful pure lane layout.</summary>
/// <param name="Lanes">The deterministic lane collection.</param>
/// <param name="Warnings">The deterministic non-blocking layout warnings.</param>
public sealed record LaneLayoutResult(IReadOnlyList<LaneGeometry> Lanes, IReadOnlyList<SceneWarning> Warnings);
