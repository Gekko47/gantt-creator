namespace GanttCreator.Core;

/// <summary>
/// One body row of the visible <c>tblGanttData</c> table in neutral form.
/// Every workbook column has a typed cell state; normalized views preserve the
/// pre-R2.4b caller contract until R2.5a consumes raw states.
/// </summary>
public sealed record GanttRowDto
{
    /// <summary>Creates a row from typed neutral cells.</summary>
    public GanttRowDto(
        int rowNumber,
        GanttCell<string> idCell,
        GanttCell<string> laneIdCell,
        GanttCell<int?> stackIndexCell,
        GanttCell<string> typeCell,
        GanttCell<string> descriptionCell,
        GanttCell<DateOnly?> startCell,
        GanttCell<DateOnly?> finishCell,
        GanttCell<string> parentIdCell,
        GanttCell<string> styleKeyCell,
        GanttCell<string> labelPositionCell,
        GanttCell<string> fillColourCell,
        GanttCell<string> strokeColourCell,
        GanttCell<bool?> visibleCell,
        GanttCell<string> sortOrderCell)
    {
        RowNumber = rowNumber;
        IdCell = idCell;
        LaneIdCell = laneIdCell;
        StackIndexCell = stackIndexCell;
        TypeCell = typeCell;
        DescriptionCell = descriptionCell;
        StartCell = startCell;
        FinishCell = finishCell;
        ParentIdCell = parentIdCell;
        StyleKeyCell = styleKeyCell;
        LabelPositionCell = labelPositionCell;
        FillColourCell = fillColourCell;
        StrokeColourCell = strokeColourCell;
        VisibleCell = visibleCell;
        SortOrderCell = sortOrderCell;
    }

    /// <summary>Creates a normalized compatibility row from the pre-R2.4b shape.</summary>
    public GanttRowDto(
        int rowNumber,
        string? id,
        string? laneId,
        int? stackIndex,
        string? typeText,
        string? description,
        DateOnly? start,
        DateOnly? finish,
        string? parentId,
        string? styleKey,
        string? labelPositionText,
        string? fillColourText,
        string? strokeColourText,
        bool? visible,
        string? sortOrder)
        : this(
            rowNumber,
            Cell(id),
            Cell(laneId),
            Cell(stackIndex),
            Cell(typeText),
            Cell(description),
            Cell(start),
            Cell(finish),
            Cell(parentId),
            Cell(styleKey),
            Cell(labelPositionText),
            Cell(fillColourText),
            Cell(strokeColourText),
            Cell(visible),
            Cell(sortOrder))
    {
    }

    /// <summary>The one-based body-row index.</summary>
    public int RowNumber { get; init; }

    /// <summary>Raw typed Id cell state.</summary>
    public GanttCell<string> IdCell { get; init; }

    /// <summary>Raw typed LaneId cell state.</summary>
    public GanttCell<string> LaneIdCell { get; init; }

    /// <summary>Raw typed StackIndex cell state.</summary>
    public GanttCell<int?> StackIndexCell { get; init; }

    /// <summary>Raw typed Type cell state.</summary>
    public GanttCell<string> TypeCell { get; init; }

    /// <summary>Raw typed Description cell state.</summary>
    public GanttCell<string> DescriptionCell { get; init; }

    /// <summary>Raw typed Start cell state.</summary>
    public GanttCell<DateOnly?> StartCell { get; init; }

    /// <summary>Raw typed Finish cell state.</summary>
    public GanttCell<DateOnly?> FinishCell { get; init; }

    /// <summary>Raw typed ParentId cell state.</summary>
    public GanttCell<string> ParentIdCell { get; init; }

    /// <summary>Raw typed StyleKey cell state.</summary>
    public GanttCell<string> StyleKeyCell { get; init; }

    /// <summary>Raw typed LabelPosition cell state.</summary>
    public GanttCell<string> LabelPositionCell { get; init; }

    /// <summary>Raw typed FillColour cell state.</summary>
    public GanttCell<string> FillColourCell { get; init; }

    /// <summary>Raw typed StrokeColour cell state.</summary>
    public GanttCell<string> StrokeColourCell { get; init; }

    /// <summary>Raw typed Visible cell state.</summary>
    public GanttCell<bool?> VisibleCell { get; init; }

    /// <summary>Raw typed SortOrder cell state.</summary>
    public GanttCell<string> SortOrderCell { get; init; }
    /// <summary>Normalized Id compatibility view.</summary>
    public string? Id => IdCell.Value;

    /// <summary>Normalized LaneId compatibility view.</summary>
    public string? LaneId => LaneIdCell.Value;

    /// <summary>Normalized StackIndex compatibility view.</summary>
    public int? StackIndex => StackIndexCell.Value;

    /// <summary>Normalized Type compatibility view.</summary>
    public string? TypeText => TypeCell.Value;

    /// <summary>Normalized Description compatibility view.</summary>
    public string? Description => DescriptionCell.Value;

    /// <summary>Normalized Start compatibility view.</summary>
    public DateOnly? Start => StartCell.Value;

    /// <summary>Normalized Finish compatibility view.</summary>
    public DateOnly? Finish => FinishCell.Value;

    /// <summary>Normalized ParentId compatibility view.</summary>
    public string? ParentId => ParentIdCell.Value;

    /// <summary>Normalized StyleKey compatibility view.</summary>
    public string? StyleKey => StyleKeyCell.Value;

    /// <summary>Normalized LabelPosition compatibility view.</summary>
    public string? LabelPositionText => LabelPositionCell.Value;

    /// <summary>Normalized FillColour compatibility view.</summary>
    public string? FillColourText => FillColourCell.Value;

    /// <summary>Normalized StrokeColour compatibility view.</summary>
    public string? StrokeColourText => StrokeColourCell.Value;

    /// <summary>Normalized Visible compatibility view.</summary>
    public bool? Visible => VisibleCell.Value;

    /// <summary>Normalized SortOrder compatibility view.</summary>
    public string? SortOrder => SortOrderCell.Value;

    private static GanttCell<T?> Cell<T>(T? value) where T : struct
        => value is null ? GanttCells.Empty<T?>() : GanttCells.Value(value);

    private static GanttCell<string> Cell(string? value)
        => value is null ? GanttCells.Empty<string>() : GanttCells.Value(value);
}
