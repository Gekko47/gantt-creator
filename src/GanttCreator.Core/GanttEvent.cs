namespace GanttCreator.Core;

/// <summary>
/// One validated <c>tblGanttData</c> body row, ready for scene construction.
/// Only rows with no blocking <see cref="GanttValidationSeverity.Error"/>
/// become events; rows carrying only warnings are included. Immutable.
/// </summary>
/// <param name="RowNumber">The one-based body-row index; an ordering aid, never identity.</param>
/// <param name="Id">The stable row identifier.</param>
/// <param name="LaneId">The stable visual-lane identifier, or <see langword="null"/> when the type is not lane-bound or the cell is blank and optional.</param>
/// <param name="StackIndex">The non-negative vertical-band order, or <see langword="null"/> when the type is not lane-bound or the cell is blank and optional.</param>
/// <param name="Type">The parsed entity type.</param>
/// <param name="Description">The trimmed description, or <see langword="null"/> when blank. Blank is permitted for every type (R2.5 U2).</param>
/// <param name="Start">The inclusive start date, or <see langword="null"/> when the type reads no dates.</param>
/// <param name="Finish">The inclusive finish date for spans, or <see langword="null"/> for point events and dateless types.</param>
/// <param name="ParentId">The owning activity for critical intervals, or <see langword="null"/> otherwise.</param>
/// <param name="StyleKey">The trimmed named style key, or <see langword="null"/> when blank (the type default applies).</param>
/// <param name="LabelPosition">The explicit label position, or <see langword="null"/> when blank (the type default or <c>Auto</c> applies).</param>
/// <param name="FillColour">The normalised uppercase <c>#RRGGBB</c> fill override, or <see langword="null"/> when blank.</param>
/// <param name="StrokeColour">The normalised uppercase <c>#RRGGBB</c> stroke override, or <see langword="null"/> when blank.</param>
/// <param name="Visible">Whether the entity renders; blank input resolves to <see langword="true"/>.</param>
/// <param name="SortOrder">The explicit ordering key, or <see langword="null"/> when blank.</param>
public sealed record GanttEvent(
    int RowNumber,
    GanttRowId Id,
    GanttRowId? LaneId,
    int? StackIndex,
    GanttEntityType Type,
    string? Description,
    DateOnly? Start,
    DateOnly? Finish,
    GanttRowId? ParentId,
    string? StyleKey,
    GanttLabelPosition? LabelPosition,
    string? FillColour,
    string? StrokeColour,
    bool Visible,
    int? SortOrder);
