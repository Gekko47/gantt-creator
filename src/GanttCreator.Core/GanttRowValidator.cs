using System.Globalization;

namespace GanttCreator.Core;

/// <summary>
/// Maps neutral <see cref="GanttRowDto"/>s into validated <see cref="GanttEvent"/>s
/// with all-errors reporting. Pure and Office-free: operates only on the DTO
/// values produced by the R2.4 table reader. Never throws for routine bad user
/// input; every row-level fault becomes a <see cref="GanttValidationIssue"/>.
/// Issues sort deterministically by severity, row, field, then code.
/// </summary>
/// <remarks>
/// <para>
/// Rule summary (R2.5 work item): <c>Id</c> required and well-formed with
/// first-canonical duplicate policy; <c>Type</c> must be an exact catalogue
/// display name; dates follow the type's <see cref="EntityDateMode"/>
/// (spans require <c>Start ≤ Finish</c>, point events read <c>Start</c> only);
/// lane/stack are required for spans and critical intervals (lane optional for
/// critical intervals when the parent supplies it), optional for milestones,
/// and warned-not-read for delineators, splitters, and spacers;
/// <c>Critical Interval</c> requires a <c>ParentId</c> present in the same
/// batch whose target is a span; <c>Custom Activity</c> requires a
/// <c>StyleKey</c> whose existence is deferred to R2.9; label positions and
/// colour overrides are checked against the type's catalogue capabilities
/// (skipped for Custom); <c>Description</c> is optional for every type;
/// <c>SortOrder</c> is blank or an invariant non-negative integer; blank
/// <c>Visible</c> resolves to <see langword="true"/>.
/// </para>
/// </remarks>
public static class GanttRowValidator
{
    /// <summary>
    /// Validates every row and maps rows without blocking errors to events.
    /// </summary>
    /// <param name="rows">The neutral body rows in table order.</param>
    /// <returns>The valid events plus every error and warning in deterministic order.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rows"/> is <see langword="null"/>.</exception>
    public static GanttValidationOutcome Validate(IReadOnlyList<GanttRowDto> rows)
    {
        ArgumentNullException.ThrowIfNull(rows);

        var issues = new List<GanttValidationIssue>();
        var perRow = new ValidatedRow?[rows.Count];

        for (var i = 0; i < rows.Count; i++)
        {
            perRow[i] = ValidateFields(rows[i], issues);
        }

        // Index first-canonical IDs: the first row carrying well-formed text wins,
        // even when that row has other blocking errors. Later rows with the same
        // trimmed text are duplicates.
        var canonicalById = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < rows.Count; i++)
        {
            ValidatedRow? parsed = perRow[i];
            if (parsed?.Id is null)
            {
                continue;
            }

            var key = parsed.Id.Value;
            if (!canonicalById.ContainsKey(key))
            {
                canonicalById[key] = i;
            }
            else
            {
                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "Id",
                        GanttValidationCodes.DuplicateId,
                        GanttValidationSeverity.Error,
                        $"Duplicate Id '{key}'; the first row carrying it is canonical."
                    )
                );
                perRow[i] = perRow[i]! with { HasBlockingError = true };
            }
        }

        CheckCriticalParents(rows, perRow, canonicalById, issues);

        var events = new List<GanttEvent>();
        for (var i = 0; i < rows.Count; i++)
        {
            ValidatedRow? parsed = perRow[i];
            if (parsed is not null && !parsed.HasBlockingError)
            {
                events.Add(parsed.Event!);
            }
        }

        GanttValidationIssue[] ordered =
        [
            .. issues
                .OrderBy(issue => issue.Severity)
                .ThenBy(issue => issue.RowNumber)
                .ThenBy(issue => issue.Field, StringComparer.Ordinal)
                .ThenBy(issue => issue.Code, StringComparer.Ordinal),
        ];

        return new GanttValidationOutcome(
            Array.AsReadOnly([.. events]),
            Array.AsReadOnly(ordered),
            !ordered.Any(issue => issue.Severity == GanttValidationSeverity.Error)
        );
    }

    /// <summary>Per-row state threaded between the field pass and the cross-row pass.</summary>
    private sealed record ValidatedRow(
        GanttRowId? Id,
        GanttEntityType? Type,
        EntityTypeDefinition? Definition,
        GanttEvent? Event,
        bool HasBlockingError
    );

    private static ValidatedRow? ValidateFields(GanttRowDto row, List<GanttValidationIssue> issues)
    {
        var rowNumber = row.RowNumber;
        var hasError = false;

        void Add(string field, string code, GanttValidationSeverity severity, string message)
        {
            issues.Add(new GanttValidationIssue(rowNumber, field, code, severity, message));
            if (severity == GanttValidationSeverity.Error)
            {
                hasError = true;
            }
        }

        // Id: required, well-formed. Duplicate detection is cross-row.
        GanttRowId? id = null;
        if (!GanttRowId.TryParse(row.Id, out GanttRowId? parsedId) || parsedId is null)
        {
            Add(
                "Id",
                GanttValidationCodes.IdMissingOrMalformed,
                GanttValidationSeverity.Error,
                "Id is blank or malformed; expected 'G-' plus 32 lowercase hex digits."
            );
        }
        else
        {
            id = parsedId;
        }

        // Type: required, exact catalogue display name.
        GanttEntityType? type = null;
        EntityTypeDefinition? definition = null;
        if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType parsedType))
        {
            Add(
                "Type",
                GanttValidationCodes.UnknownType,
                GanttValidationSeverity.Error,
                $"Unknown Type '{row.TypeText}'; use one of the 16 catalogue display names."
            );
        }
        else
        {
            type = parsedType;
            definition = EntityTypeCatalog.GetDefinition(parsedType);
        }

        // Lane/Stack presence by kind. Splitter/Spacer are never lane-bound;
        // delineators carry no lane geometry; milestones treat both as optional.
        var isSplitterOrSpacer = type is GanttEntityType.Splitter or GanttEntityType.Spacer;
        var isDelineator = type == GanttEntityType.Delineator;
        var isCriticalInterval = type == GanttEntityType.CriticalInterval;
        var isMilestone =
            type
            is GanttEntityType.AsBuiltMilestone
                or GanttEntityType.AsPlannedMilestone
                or GanttEntityType.BaselineMilestone
                or GanttEntityType.CriticalMilestone;
        var laneRequired = definition is not null && !isSplitterOrSpacer && !isDelineator && !isMilestone && !isCriticalInterval;
        var stackRequired = laneRequired;

        GanttRowId? laneId = null;
        if (row.LaneId is null)
        {
            if (laneRequired)
            {
                Add(
                    "LaneId",
                    GanttValidationCodes.LaneIdMissingOrMalformed,
                    GanttValidationSeverity.Error,
                    "LaneId is required for this Type."
                );
            }
        }
        else if (isSplitterOrSpacer || isDelineator)
        {
            Add(
                "LaneId",
                GanttValidationCodes.NotUsedByType,
                GanttValidationSeverity.Warning,
                "LaneId is not used by this Type and is ignored for geometry."
            );
        }
        else if (!GanttRowId.TryParse(row.LaneId, out GanttRowId? laneParsed) || laneParsed is null)
        {
            Add(
                "LaneId",
                GanttValidationCodes.LaneIdMissingOrMalformed,
                GanttValidationSeverity.Error,
                $"LaneId '{row.LaneId}' is malformed; expected 'G-' plus 32 lowercase hex digits."
            );
        }
        else
        {
            laneId = laneParsed;
        }

        int? stackIndex = null;
        if (row.StackIndex is null)
        {
            if (stackRequired)
            {
                Add(
                    "StackIndex",
                    GanttValidationCodes.StackIndexRequired,
                    GanttValidationSeverity.Error,
                    "StackIndex is required for this Type."
                );
            }
        }
        else if (isSplitterOrSpacer || isDelineator)
        {
            Add(
                "StackIndex",
                GanttValidationCodes.NotUsedByType,
                GanttValidationSeverity.Warning,
                "StackIndex is not used by this Type and is ignored for geometry."
            );
        }
        else if (row.StackIndex.Value < 0)
        {
            Add(
                "StackIndex",
                GanttValidationCodes.StackIndexNegative,
                GanttValidationSeverity.Error,
                "StackIndex must be zero or greater."
            );
        }
        else
        {
            stackIndex = row.StackIndex.Value;
        }

        // Dates by EntityDateMode.
        DateOnly? start = row.Start;
        DateOnly? finish = row.Finish;
        if (definition is not null)
        {
            switch (definition.DateMode)
            {
                case EntityDateMode.None:
                    if (start is not null)
                    {
                        Add(
                            "Start",
                            GanttValidationCodes.NotUsedByType,
                            GanttValidationSeverity.Warning,
                            "Start is not used by this Type and is ignored for geometry."
                        );
                        start = null;
                    }

                    if (finish is not null)
                    {
                        Add(
                            "Finish",
                            GanttValidationCodes.NotUsedByType,
                            GanttValidationSeverity.Warning,
                            "Finish is not used by this Type and is ignored for geometry."
                        );
                        finish = null;
                    }

                    break;
                case EntityDateMode.StartFinish:
                    if (start is null)
                    {
                        Add("Start", GanttValidationCodes.StartRequired, GanttValidationSeverity.Error, "Start is required for this Type.");
                    }

                    if (finish is null)
                    {
                        Add(
                            "Finish",
                            GanttValidationCodes.FinishRequired,
                            GanttValidationSeverity.Error,
                            "Finish is required for this Type."
                        );
                    }
                    else if (start is not null && start.Value > finish.Value)
                    {
                        Add(
                            "Finish",
                            GanttValidationCodes.StartAfterFinish,
                            GanttValidationSeverity.Error,
                            "Start must not be after Finish."
                        );
                    }

                    break;
                case EntityDateMode.StartOnly:
                    if (start is null)
                    {
                        Add(
                            "Start",
                            GanttValidationCodes.StartRequiredForPoint,
                            GanttValidationSeverity.Error,
                            "Start is the single date for this Type and is required."
                        );
                    }

                    if (finish is not null)
                    {
                        Add(
                            "Finish",
                            GanttValidationCodes.NotUsedByType,
                            GanttValidationSeverity.Warning,
                            "Finish is not used by this Type and is ignored for geometry."
                        );
                        finish = null;
                    }

                    break;
                default:
                    break;
            }
        }

        // ParentId: required only for Critical Interval; warned elsewhere.
        GanttRowId? parentId = null;
        var parentText = row.ParentId;
        if (isCriticalInterval)
        {
            if (!GanttRowId.TryParse(parentText, out GanttRowId? parentParsed) || parentParsed is null)
            {
                Add(
                    "ParentId",
                    GanttValidationCodes.ParentMissingOrMalformed,
                    GanttValidationSeverity.Error,
                    "ParentId is required for Critical Interval and must be a well-formed row Id."
                );
            }
            else
            {
                parentId = parentParsed;
            }
        }
        else if (parentText is not null)
        {
            Add(
                "ParentId",
                GanttValidationCodes.NotUsedByType,
                GanttValidationSeverity.Warning,
                "ParentId is not used by this Type and is ignored for geometry."
            );
            if (GanttRowId.TryParse(parentText, out GanttRowId? parentParsed) && parentParsed is not null)
            {
                parentId = parentParsed;
            }
        }

        // StyleKey: Custom Activity requires one (existence deferred to R2.9).
        var styleKey = row.StyleKey;
        if (type == GanttEntityType.CustomActivity && styleKey is null)
        {
            Add(
                "StyleKey",
                GanttValidationCodes.StyleKeyRequired,
                GanttValidationSeverity.Error,
                "Custom Activity requires a StyleKey; named-style resolution lands in R2.9."
            );
        }

        // LabelPosition: blank resolves later; otherwise exact-enum plus catalogue capability (skipped for Custom).
        GanttLabelPosition? labelPosition = null;
        if (row.LabelPositionText is not null && type is not null && type != GanttEntityType.CustomActivity && definition is not null)
        {
            if (
                !Enum.TryParse(row.LabelPositionText, ignoreCase: false, out GanttLabelPosition parsedLabel)
                || !Enum.IsDefined(parsedLabel)
            )
            {
                Add(
                    "LabelPosition",
                    GanttValidationCodes.UnknownLabelPosition,
                    GanttValidationSeverity.Error,
                    $"Unknown LabelPosition '{row.LabelPositionText}'."
                );
            }
            else if (!definition.AllowedLabelPositions.Contains(parsedLabel))
            {
                Add(
                    "LabelPosition",
                    GanttValidationCodes.LabelNotAllowedForType,
                    GanttValidationSeverity.Error,
                    $"LabelPosition '{row.LabelPositionText}' is not allowed for Type '{definition.DisplayName}'."
                );
            }
            else
            {
                labelPosition = parsedLabel;
            }
        }

        // Colours: format plus capability (skipped for Custom; deferred to R2.9).
        var fill = ValidateColour(row.FillColourText, isFill: true, type, definition, Add);
        var stroke = ValidateColour(row.StrokeColourText, isFill: false, type, definition, Add);

        // SortOrder: blank or invariant non-negative int.
        int? sortOrder = null;
        if (row.SortOrder is not null)
        {
            if (!int.TryParse(row.SortOrder, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedSort) || parsedSort < 0)
            {
                Add(
                    "SortOrder",
                    GanttValidationCodes.BadSortOrder,
                    GanttValidationSeverity.Error,
                    $"SortOrder '{row.SortOrder}' must be a non-negative integer."
                );
            }
            else
            {
                sortOrder = parsedSort;
            }
        }

        if (id is null || type is null)
        {
            return new ValidatedRow(id, type, definition, null, HasBlockingError: true);
        }

        GanttRowId? laneForEvent = (isSplitterOrSpacer || isDelineator) ? null : laneId;
        var stackForEvent = (isSplitterOrSpacer || isDelineator) ? null : stackIndex;

        var @event = new GanttEvent(
            rowNumber,
            id,
            laneForEvent,
            stackForEvent,
            type.Value,
            row.Description,
            start,
            finish,
            parentId,
            styleKey,
            labelPosition,
            fill,
            stroke,
            row.Visible ?? true,
            sortOrder
        );

        return new ValidatedRow(id, type, definition, @event, hasError);
    }

    private static string? ValidateColour(
        string? text,
        bool isFill,
        GanttEntityType? type,
        EntityTypeDefinition? definition,
        Action<string, string, GanttValidationSeverity, string> add
    )
    {
        if (text is null || type is null || definition is null)
        {
            return null;
        }

        // R2.5 U4: Custom Activity defers all colour resolution to R2.9.
        if (type == GanttEntityType.CustomActivity)
        {
            return text.ToUpperInvariant();
        }

        var field = isFill ? "FillColour" : "StrokeColour";
        if (!IsHexColour(text))
        {
            add(field, GanttValidationCodes.BadColourFormat, GanttValidationSeverity.Error, $"{field} '{text}' must be '#RRGGBB'.");
            return null;
        }

        EntityColourCapability capability = definition.ColourCapability;
        var allowed = isFill
            ? capability.HasFlag(EntityColourCapability.Fill)
            : capability.HasFlag(EntityColourCapability.Stroke) || capability.HasFlag(EntityColourCapability.Hatch);
        if (!allowed)
        {
            add(
                field,
                GanttValidationCodes.ColourNotAllowedForType,
                GanttValidationSeverity.Error,
                $"{field} is not allowed for Type '{definition.DisplayName}'."
            );
            return null;
        }

        return text.ToUpperInvariant();
    }

    private static bool IsHexColour(string text)
    {
        if (text.Length != 7 || text[0] != '#')
        {
            return false;
        }

        for (var i = 1; i < text.Length; i++)
        {
            var c = text[i];
            var isHex = c is (>= '0' and <= '9') or (>= 'a' and <= 'f') or (>= 'A' and <= 'F');
            if (!isHex)
            {
                return false;
            }
        }

        return true;
    }

    private static void CheckCriticalParents(
        IReadOnlyList<GanttRowDto> rows,
        ValidatedRow?[] perRow,
        Dictionary<string, int> canonicalById,
        List<GanttValidationIssue> issues
    )
    {
        // Resolve each Critical Interval ParentId against the same batch.
        for (var i = 0; i < rows.Count; i++)
        {
            ValidatedRow? parsed = perRow[i];
            if (parsed?.Type != GanttEntityType.CriticalInterval || parsed.Event?.ParentId is null)
            {
                continue;
            }

            var key = parsed.Event.ParentId.Value;
            if (!canonicalById.TryGetValue(key, out var parentIndex))
            {
                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "ParentId",
                        GanttValidationCodes.ParentUnknown,
                        GanttValidationSeverity.Error,
                        $"ParentId '{key}' does not match any Id in the table."
                    )
                );
                perRow[i] = parsed with { HasBlockingError = true };
                continue;
            }

            GanttEntityType? parentType = perRow[parentIndex]?.Type;
            if (parentType is null)
            {
                // Parent row itself has an unknown type; the interval cannot attach.
                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "ParentId",
                        GanttValidationCodes.ParentNotSpan,
                        GanttValidationSeverity.Error,
                        $"ParentId '{key}' does not reference a span event."
                    )
                );
                perRow[i] = parsed with { HasBlockingError = true };
                continue;
            }

            EntityTypeDefinition? parentDefinition = EntityTypeCatalog.GetDefinition(parentType.Value);
            if (parentDefinition is null || parentDefinition.Kind != EntityKind.Span)
            {
                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "ParentId",
                        GanttValidationCodes.ParentNotSpan,
                        GanttValidationSeverity.Error,
                        $"ParentId '{key}' must reference a span event."
                    )
                );
                perRow[i] = parsed with { HasBlockingError = true };
            }
        }
    }
}
