namespace GanttCreator.Core.Export;

/// <summary>
/// The primitive facts one <c>Shapes.PasteSpecial(ppPasteShape)</c> result reported (R7.3).
/// </summary>
/// <remarks>
/// <para>
/// Exactly the members the work item names — count, group count, bounds and group type — so
/// the validator is testable without a host and the AddIn compilation never names the
/// PowerPoint PIA. The group's identity (<c>Id</c>, <c>Name</c>) and its own child count are
/// Office-side observation members: the stage result carries them forward for R7.4, not the
/// Core validator.
/// </para>
/// <para>
/// A range with no <c>msoGroup</c> reports <paramref name="GroupCount"/> 0 and a
/// <see langword="null"/> <paramref name="GroupShapeType"/> — never a sentinel that could be
/// mistaken for a type.
/// </para>
/// </remarks>
/// <param name="Count">The returned <c>ShapeRange.Count</c>.</param>
/// <param name="GroupCount">How many <c>msoGroup</c> shapes the range contains.</param>
/// <param name="GroupShapeType">
/// The located group's <c>Shape.Type</c> value, or <see langword="null"/> when the range holds
/// no group. Core cannot name the interop enum, so this is the documented <c>PpShapeType</c>
/// value pinned by <see cref="PowerPointPasteValidator.MsoGroupShapeType"/>.
/// </param>
/// <param name="WidthPt">The returned range width in points.</param>
/// <param name="HeightPt">The returned range height in points.</param>
public sealed record PowerPointPasteFacts(
    int Count,
    int GroupCount,
    int? GroupShapeType,
    double WidthPt,
    double HeightPt);
