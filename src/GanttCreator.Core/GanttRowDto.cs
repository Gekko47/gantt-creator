namespace GanttCreator.Core;

/// <summary>
/// One body row of the visible <c>tblGanttData</c> table in neutral,
/// culture-invariant form. Carries raw cell values without domain validation:
/// unknown types, missing IDs, and reversed dates are R2.5 scope. Immutable.
/// </summary>
/// <param name="RowNumber">The one-based body-row index; an ordering aid, never identity.</param>
/// <param name="Id">The trimmed <c>Id</c> text, or <see langword="null"/> when blank.</param>
/// <param name="LaneId">The trimmed <c>LaneId</c> text, or <see langword="null"/> when blank.</param>
/// <param name="StackIndex">The <c>StackIndex</c> value, or <see langword="null"/> when blank.</param>
/// <param name="TypeText">The trimmed <c>Type</c> text, or <see langword="null"/> when blank. Parsing to <see cref="GanttEntityType"/> is R2.5 scope.</param>
/// <param name="Description">The trimmed <c>Description</c> text, or <see langword="null"/> when blank.</param>
/// <param name="Start">The <c>Start</c> date, or <see langword="null"/> when blank or unreadable.</param>
/// <param name="Finish">The <c>Finish</c> date, or <see langword="null"/> when blank or unreadable.</param>
/// <param name="ParentId">The trimmed <c>ParentId</c> text, or <see langword="null"/> when blank.</param>
/// <param name="StyleKey">The trimmed <c>StyleKey</c> text, or <see langword="null"/> when blank.</param>
/// <param name="LabelPositionText">The trimmed <c>LabelPosition</c> text, or <see langword="null"/> when blank.</param>
/// <param name="FillColourText">The trimmed <c>FillColour</c> text, or <see langword="null"/> when blank.</param>
/// <param name="StrokeColourText">The trimmed <c>StrokeColour</c> text, or <see langword="null"/> when blank.</param>
/// <param name="Visible">The <c>Visible</c> value, or <see langword="null"/> when blank.</param>
/// <param name="SortOrder">The raw trimmed <c>SortOrder</c> text, or <see langword="null"/> when blank. Coercion is R2.5 scope.</param>
public sealed record GanttRowDto(
    int RowNumber,
    string? Id,
    string? LaneId,
    int? StackIndex,
    string? TypeText,
    string? Description,
    DateOnly? Start,
    DateOnly? Finish,
    string? ParentId,
    string? StyleKey,
    string? LabelPositionText,
    string? FillColourText,
    string? StrokeColourText,
    bool? Visible,
    string? SortOrder);
