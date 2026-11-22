using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Export;

/// <summary>
/// What the validator expects one <c>ppPasteShape</c> result to report (R7.3): the staged
/// child count, the staged group envelope, and the tolerance the bounds comparison uses.
/// </summary>
/// <remarks>
/// <para>
/// The expectation is assembled by the caller from the verified staged group —
/// <c>StagedGroupOutcome.ChildOrder.Count</c> and <c>StagedGroupOutcome.Bounds</c> — so the
/// validator never re-derives either. <paramref name="TolerancePt"/> defaults to the
/// repository's host read-back tolerance
/// (<see cref="HostOwnedGeometry.GeometryMatchTolerancePt"/>); it is a parameter only so a
/// test can pin the comparison at a stated value, never so a production path can loosen it.
/// </para>
/// <para>
/// A non-positive child count or a non-finite or non-positive expected bounds is a caller
/// error that no host result can satisfy; <see cref="PowerPointPasteValidator"/> refuses it
/// with its own typed members before it reads the facts.
/// </para>
/// </remarks>
/// <param name="ExpectedChildCount">The staged child count N; the expected range is <c>1 + N</c>.</param>
/// <param name="ExpectedWidthPt">The staged group envelope width in points.</param>
/// <param name="ExpectedHeightPt">The staged group envelope height in points.</param>
/// <param name="TolerancePt">The per-axis tolerance in points; defaults to the repository's host read-back tolerance.</param>
public sealed record PowerPointPasteExpectation(
    int ExpectedChildCount,
    double ExpectedWidthPt,
    double ExpectedHeightPt,
    double TolerancePt = HostOwnedGeometry.GeometryMatchTolerancePt);
