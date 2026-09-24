namespace GanttCreator.Core;

/// <summary>
/// Stable machine-readable codes for every <see cref="GanttValidationIssue"/>
/// produced by <see cref="GanttRowValidator"/>. Codes are part of the
/// validation contract: R2.6 error reporting keys off them.
/// </summary>
public static class GanttValidationCodes
{
    /// <summary>The <c>Id</c> cell is blank or not a well-formed <see cref="GanttRowId"/>.</summary>
    public const string IdMissingOrMalformed = nameof(IdMissingOrMalformed);

    /// <summary>A later row repeats an <c>Id</c> already seen; the first occurrence is canonical.</summary>
    public const string DuplicateId = nameof(DuplicateId);

    /// <summary>The <c>Type</c> cell is blank or not an exact <see cref="EntityTypeCatalog"/> display name.</summary>
    public const string UnknownType = nameof(UnknownType);

    /// <summary>A span row is missing <c>Start</c>.</summary>
    public const string StartRequired = nameof(StartRequired);

    /// <summary>A span row is missing <c>Finish</c>.</summary>
    public const string FinishRequired = nameof(FinishRequired);

    /// <summary>A span row has <c>Start</c> after <c>Finish</c>.</summary>
    public const string StartAfterFinish = nameof(StartAfterFinish);

    /// <summary>A point-event row is missing <c>Start</c>, its single date.</summary>
    public const string StartRequiredForPoint = nameof(StartRequiredForPoint);

    /// <summary>A lane-bound row is missing <c>LaneId</c> or it is malformed.</summary>
    public const string LaneIdMissingOrMalformed = nameof(LaneIdMissingOrMalformed);

    /// <summary>A lane-bound row is missing <c>StackIndex</c>.</summary>
    public const string StackIndexRequired = nameof(StackIndexRequired);

    /// <summary>A <c>StackIndex</c> value is negative.</summary>
    public const string StackIndexNegative = nameof(StackIndexNegative);

    /// <summary>A relevant cell contains an Excel error value.</summary>
    public const string CellContainsExcelError = nameof(CellContainsExcelError);

    /// <summary>A relevant cell contains a payload unsupported by its column type.</summary>
    public const string CellValueUnsupported = nameof(CellValueUnsupported);

    /// <summary>A field carries a value the row's type never reads for geometry; the value is retained.</summary>
    public const string NotUsedByType = nameof(NotUsedByType);

    /// <summary>A <c>Critical Interval</c> row is missing <c>ParentId</c> or it is malformed.</summary>
    public const string ParentMissingOrMalformed = nameof(ParentMissingOrMalformed);

    /// <summary>A <c>Critical Interval</c> references an <c>Id</c> absent from the same input batch.</summary>
    public const string ParentUnknown = nameof(ParentUnknown);

    /// <summary>A <c>Critical Interval</c> references a parent that is not a span event.</summary>
    public const string ParentNotSpan = nameof(ParentNotSpan);

    /// <summary>A <c>Critical Interval</c> references a parent row that is not valid.</summary>
    public const string ParentInvalid = nameof(ParentInvalid);

    /// <summary>A <c>Critical Interval</c> parent relationship contains a cycle.</summary>
    public const string ParentCycle = nameof(ParentCycle);

    /// <summary>A <c>Custom Activity</c> row is missing its required <c>StyleKey</c>.</summary>
    public const string StyleKeyRequired = nameof(StyleKeyRequired);

    /// <summary>A <c>LabelPosition</c> value is not a known position name.</summary>
    public const string UnknownLabelPosition = nameof(UnknownLabelPosition);

    /// <summary>A <c>LabelPosition</c> value is not permitted for the row's type.</summary>
    public const string LabelNotAllowedForType = nameof(LabelNotAllowedForType);

    /// <summary>A colour override is not <c>#RRGGBB</c>.</summary>
    public const string BadColourFormat = nameof(BadColourFormat);

    /// <summary>A colour override is not permitted for the row's type.</summary>
    public const string ColourNotAllowedForType = nameof(ColourNotAllowedForType);

    /// <summary>A <c>SortOrder</c> value is not an invariant non-negative integer.</summary>
    public const string BadSortOrder = nameof(BadSortOrder);
}
