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
/// <c>LaneId</c> and <c>StackIndex</c> are optional for every type because they
/// are engine-generated, but a supplied value is still checked (well-formed lane
/// id, non-negative stack) and a value on a type that never reads it is warned
/// rather than silently dropped;
/// <c>ParentId</c> is optional for every child-capable type and a supplied value
/// must name a row in the same batch whose target may own children; a cycle is
/// refused; <c>Custom Activity</c> requires a
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
    public static GanttValidationOutcome Validate(IReadOnlyList<GanttRowDto> rows) =>
        Validate(rows, null);

    /// <summary>
    /// Validates every row using the optional named-style registry. A null
    /// registry preserves the pre-R2.9 validation behaviour.
    /// </summary>
    /// <param name="rows">The neutral body rows in table order.</param>
    /// <param name="styleRegistry">The named-style capabilities, or null for compatibility.</param>
    /// <returns>The valid events plus every error and warning in deterministic order.</returns>
    public static GanttValidationOutcome Validate(
        IReadOnlyList<GanttRowDto> rows,
        GanttStyleRegistry? styleRegistry)
    {
        ArgumentNullException.ThrowIfNull(rows);

        var issues = new List<GanttValidationIssue>();
        var perRow = new ValidatedRow?[rows.Count];

        for (var i = 0; i < rows.Count; i++)
        {
            perRow[i] = ValidateFields(rows[i], issues, styleRegistry);
        }

        // Index first-canonical IDs: the first row carrying well-formed text wins,
        // even when that row has other blocking errors. Later rows with the same
        // trimmed text are duplicates.
        var canonicalById = new Dictionary<string, int>(StringComparer.Ordinal);
        var duplicateIds = new HashSet<string>(StringComparer.Ordinal);
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
                _ = duplicateIds.Add(key);
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

        CheckCriticalParents(rows, perRow, canonicalById, duplicateIds, issues);
        CheckChildCapacityAndDepth(rows, perRow, canonicalById, duplicateIds, issues);

        // Run propagation a second time. A parent can only become blocked in the
        // pass above -- the seven-child cap is reported against the PARENT's row, and
        // the depth rule is what blocks a parent that is itself a child -- and a
        // child of such a parent would otherwise survive as a valid event pointing at
        // a row that is not in `Events`. Without this second pass the outcome also
        // depended on row order for those rows: a child authored above its parent was
        // decided before the parent's blocking finding existed.
        PropagateBlockedParents(rows, perRow, canonicalById, duplicateIds, issues);

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

    /// <summary>Classifies a raw cell once and records its applicability decision.</summary>
    /// <param name="field">The workbook column name.</param>
    /// <param name="state">The raw cell state.</param>
    /// <param name="relevant">Whether the selected type reads the field.</param>
    /// <param name="add">The row issue sink.</param>
    /// <returns><see langword="true"/> when the field was non-empty and ignored or blocked.</returns>
    private static bool ClassifyCellState(
        string field,
        GanttCellState state,
        bool relevant,
        Action<string, string, GanttValidationSeverity, string> add
    )
    {
        if (state == GanttCellState.Empty)
        {
            return false;
        }

        if (!relevant)
        {
            add(
                field,
                GanttValidationCodes.NotUsedByType,
                GanttValidationSeverity.Warning,
                $"{field} is not used by this Type and is ignored for geometry."
            );
            return true;
        }

        if (state == GanttCellState.ExcelError)
        {
            add(
                field,
                GanttValidationCodes.CellContainsExcelError,
                GanttValidationSeverity.Error,
                $"{field} contains an Excel error value."
            );
            return true;
        }

        if (state == GanttCellState.Unsupported)
        {
            add(
                field,
                GanttValidationCodes.CellValueUnsupported,
                GanttValidationSeverity.Error,
                $"{field} contains a value unsupported by its column type."
            );
            return true;
        }

        return false;
    }

    private static ValidatedRow? ValidateFields(
        GanttRowDto row,
        List<GanttValidationIssue> issues,
        GanttStyleRegistry? styleRegistry)
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

        var idBlocked = ClassifyCellState("Id", row.IdCell.State, relevant: true, Add);
        var typeBlocked = ClassifyCellState("Type", row.TypeCell.State, relevant: true, Add);

        // Id: required, well-formed. Duplicate detection is cross-row.
        GanttRowId? id = null;
        GanttRowId? parsedId = null;
        if (!idBlocked && (!GanttRowId.TryParse(row.Id, out parsedId) || parsedId is null))
        {
            Add(
                "Id",
                GanttValidationCodes.IdMissingOrMalformed,
                GanttValidationSeverity.Error,
                "Id is blank or malformed; expected 'G-' plus 32 lowercase hex digits."
            );
        }
        else if (!idBlocked)
        {
            id = parsedId;
        }

        // Type: required, exact catalogue display name.
        GanttEntityType? type = null;
        EntityTypeDefinition? definition = null;
        GanttEntityType parsedType = default;
        if (!typeBlocked && !EntityTypeCatalog.TryParse(row.TypeText, out parsedType))
        {
            Add(
                "Type",
                GanttValidationCodes.UnknownType,
                GanttValidationSeverity.Error,
                $"Unknown Type '{row.TypeText}'; use one of the 16 catalogue display names."
            );
        }
        else if (!typeBlocked)
        {
            type = parsedType;
            definition = EntityTypeCatalog.GetDefinition(parsedType);
        }

        // Lane/Stack relevance by kind. Splitter/Spacer are never lane-bound;
        // delineators carry no lane geometry. There is deliberately NO
        // `laneRequired`/`stackRequired` counterpart: both columns are
        // `EngineHidden` (ADR-0029 D8) and R2.8's `GanttRowDefaults` scaffolds
        // them blank, so requiring either one made every Add-Row activity
        // unvalidatable with a blocking error on a cell the user cannot even see.
        // ADR-0012 already made the visible `StackIndex` untrusted and unrequired
        // for layout, and R4.7B derives render-lane ownership in Core
        // (`ProjectionResolver`) rather than reading the `LaneId` cell --
        // `LaneOrdering.LaneKey` already returns a row-scoped key when it is null.
        // A SUPPLIED value is still validated, so this is optional rather than
        // unvalidated.
        var isSplitterOrSpacer = type is GanttEntityType.Splitter or GanttEntityType.Spacer;
        var isDelineator = type == GanttEntityType.Delineator;
        var laneRelevant = !isSplitterOrSpacer && !isDelineator;
        var startRelevant = definition?.DateMode != EntityDateMode.None;
        var finishRelevant = definition?.DateMode == EntityDateMode.StartFinish;
        // R4.7A D9: ParentId is authoritative for general hierarchy, not only for
        // Critical Interval, so relevance is driven by the child-capability matrix
        // rather than by a hard-coded `isCriticalInterval`. `definition is not null`
        // is required first, so an unknown Type cannot make ParentId relevant and
        // report a second, misleading error on top of the UnknownType one. Types the
        // matrix does not classify as child-capable -- Splitter, Spacer, Delineator --
        // still fall through to NotUsedByType below, so nothing becomes silently
        // permitted.
        var parentRelevant = definition is not null && EntityHierarchyCatalog.ParentIdIsRelevant(parsedType);
        var styleRelevant = true;
        var labelRelevant = true;
        var fillRelevant = true;
        var strokeRelevant = true;
        const bool visibleRelevant = true;
        const bool sortOrderRelevant = true;

        var laneBlocked = ClassifyCellState("LaneId", row.LaneIdCell.State, laneRelevant, Add);
        var stackBlocked = ClassifyCellState("StackIndex", row.StackIndexCell.State, laneRelevant, Add);
        var startBlocked = ClassifyCellState("Start", row.StartCell.State, startRelevant, Add);
        var finishBlocked = ClassifyCellState("Finish", row.FinishCell.State, finishRelevant, Add);
        var parentBlocked = ClassifyCellState("ParentId", row.ParentIdCell.State, parentRelevant, Add);
        var styleBlocked = ClassifyCellState("StyleKey", row.StyleKeyCell.State, styleRelevant, Add);
        var labelBlocked = ClassifyCellState("LabelPosition", row.LabelPositionCell.State, labelRelevant, Add);
        var fillBlocked = ClassifyCellState("FillColour", row.FillColourCell.State, fillRelevant, Add);
        var strokeBlocked = ClassifyCellState("StrokeColour", row.StrokeColourCell.State, strokeRelevant, Add);
        _ = ClassifyCellState("Visible", row.VisibleCell.State, visibleRelevant, Add);
        _ = ClassifyCellState("SortOrder", row.SortOrderCell.State, sortOrderRelevant, Add);

        // SiblingOrder (R4.7A D2): engine-maintained, so it is always relevant and
        // is never required. A blank value means "the engine has not assigned one
        // yet" and is reported by the caller, not treated as ordering -- silently
        // reading blank as 0 would make every unassigned row a first child.
        _ = ClassifyCellState("SiblingOrder", row.SiblingOrderCell.State, true, Add);
        if (row.SiblingOrder is { } parsedSiblingOrder)
        {
            if (parsedSiblingOrder < 0)
            {
                Add(
                    "SiblingOrder",
                    GanttValidationCodes.BadSiblingOrder,
                    GanttValidationSeverity.Error,
                    "SiblingOrder must be blank or a non-negative integer."
                );
            }
        }

        GanttRowId? laneId = null;
        if (!laneBlocked)
        {
            if (row.LaneId is null)
            {
                // Blank is the scaffolded and normal state: the column is
                // engine-generated and hidden, and R4.7B derives the render lane in
                // Core. `LaneOrdering.LaneKey` gives a null lane a row-scoped key.
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
        }

        int? stackIndex = null;
        if (!stackBlocked)
        {
            if (row.StackIndex is null)
            {
                // ADR-0012: the visible cell is compatibility data that is neither
                // trusted nor required. `LaneLayoutBuilder` derives the effective
                // stack from row/parent-child position, so a blank cell is correct
                // rather than missing.
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
        }

        // Dates by EntityDateMode.
        DateOnly? start = row.Start;
        DateOnly? finish = row.Finish;
        if (definition is not null)
        {
            switch (definition.DateMode)
            {
                case EntityDateMode.None:
                    if (!startBlocked && start is not null)
                    {
                        Add(
                            "Start",
                            GanttValidationCodes.NotUsedByType,
                            GanttValidationSeverity.Warning,
                            "Start is not used by this Type and is ignored for geometry."
                        );
                        start = null;
                    }

                    if (!finishBlocked && finish is not null)
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
                    if (!startBlocked && start is null)
                    {
                        Add("Start", GanttValidationCodes.StartRequired, GanttValidationSeverity.Error, "Start is required for this Type.");
                    }

                    if (!finishBlocked && finish is null)
                    {
                        Add(
                            "Finish",
                            GanttValidationCodes.FinishRequired,
                            GanttValidationSeverity.Error,
                            "Finish is required for this Type."
                        );
                    }
                    else if (!finishBlocked && start is not null && finish is not null && start.Value > finish.Value)
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
                    if (!startBlocked && start is null)
                    {
                        Add(
                            "Start",
                            GanttValidationCodes.StartRequiredForPoint,
                            GanttValidationSeverity.Error,
                            "Start is the single date for this Type and is required."
                        );
                    }

                    if (!finishBlocked && finish is not null)
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

        if (!startRelevant)
        {
            start = null;
        }

        if (!finishRelevant)
        {
            finish = null;
        }

        // ParentId (R4.7A D9). Two cases, split by the child-capability matrix
        // rather than by `isCriticalInterval`:
        //   * a Critical Interval MAY carry a parent, but no longer REQUIRES one;
        //   * any other child-capable type may carry one -- a top-level row simply
        //     leaves it blank, which is what makes a promoted child valid again
        //     after its parent is deleted;
        //   * a type the matrix excludes still reports NotUsedByType.
        // The `isCriticalInterval` branch used to make ParentId mandatory for that
        // one type. That contradicted <see cref="EntityHierarchyCatalog"/>, which
        // lists CriticalInterval among the types that may own children precisely so
        // a TOP-LEVEL interval can own a level-2 child under the uniform depth
        // rule -- and it contradicted <c>EntityProjection</c>, which resolves such
        // a child onto the interval's own lane. The requirement therefore made the
        // catalogue's entry and the projection's boundary unreachable. A blank
        // ParentId is now legal for every child-capable type, and the DEPTH rule is
        // the single authority on what a Critical Interval may parent.
        // The previous `else if (parentText is not null)` branch was reached by any
        // non-critical row, so a child-capable span carrying a real ParentId emitted
        // a spurious "not used by this Type" warning *alongside* its legitimate
        // ParentUnknown error. `parentRelevant` is the same flag ClassifyCellState
        // used above, so the two decisions cannot disagree.
        GanttRowId? parentId = null;
        var parentText = row.ParentId;
        if (!parentBlocked && parentRelevant)
        {
            // Optional for every child-capable type: blank means top-level, which
            // is valid. A supplied value must parse, and is then resolved against
            // the batch by the cross-row parent pass.
            if (parentText is not null
                && (!GanttRowId.TryParse(parentText, out GanttRowId? parentParsed) || parentParsed is null))
            {
                Add(
                    "ParentId",
                    GanttValidationCodes.ParentMissingOrMalformed,
                    GanttValidationSeverity.Error,
                    "ParentId must be blank or a well-formed row Id."
                );
            }
            else if (parentText is not null)
            {
                _ = GanttRowId.TryParse(parentText, out parentId);
            }
        }
        else if (!parentBlocked && parentText is not null)
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

        // StyleKey: Custom Activity requires one; a supplied key must resolve
        // when the caller supplies the validated style registry.
        var styleKey = row.StyleKey;
        GanttStyleDefinition? styleDefinition = null;
        if (!styleBlocked && type == GanttEntityType.CustomActivity && styleKey is null)
        {
            Add(
                "StyleKey",
                GanttValidationCodes.StyleKeyRequired,
                GanttValidationSeverity.Error,
                "Custom Activity requires a StyleKey."
            );
        }
        else if (!styleBlocked && styleKey is not null && styleRegistry is not null)
        {
            if (!styleRegistry.TryGet(styleKey, out styleDefinition))
            {
                Add(
                    "StyleKey",
                    GanttValidationCodes.StyleKeyUnknown,
                    GanttValidationSeverity.Error,
                    $"StyleKey '{styleKey}' is not in the configured style catalogue."
                );
            }
        }

        // LabelPosition: blank resolves later; otherwise exact-enum plus the
        // selected type's or Custom Activity style's capability set.
        GanttLabelPosition? labelPosition = null;
        if (
            !labelBlocked
            && row.LabelPositionText is not null
            && type is not null
            && (type != GanttEntityType.CustomActivity || styleDefinition is not null)
        )
        {
            if (
                !Enum.TryParse(row.LabelPositionText, ignoreCase: false, out GanttLabelPosition parsedLabel) || !Enum.IsDefined(parsedLabel)
            )
            {
                Add(
                    "LabelPosition",
                    GanttValidationCodes.UnknownLabelPosition,
                    GanttValidationSeverity.Error,
                    $"Unknown LabelPosition '{row.LabelPositionText}'."
                );
            }
            else
            {
                IReadOnlySet<GanttLabelPosition> allowedLabels = type == GanttEntityType.CustomActivity
                    ? styleDefinition!.AllowedLabelPositions
                    : definition!.AllowedLabelPositions;
                if (!allowedLabels.Contains(parsedLabel))
                {
                    Add(
                        "LabelPosition",
                        GanttValidationCodes.LabelNotAllowedForType,
                        GanttValidationSeverity.Error,
                        $"LabelPosition '{row.LabelPositionText}' is not allowed for this Type."
                    );
                }
                else
                {
                    labelPosition = parsedLabel;
                }
            }
        }

        // Colours: format plus type capability, or the selected Custom style's
        // capability. A Custom row with an unresolved style remains blocked by
        // StyleKeyUnknown and receives no colour-capability finding.
        var fill = fillBlocked
            ? null
            : ValidateColour(row.FillColourText, isFill: true, type, definition, styleDefinition, styleRegistry is not null, Add);
        var stroke = strokeBlocked
            ? null
            : ValidateColour(row.StrokeColourText, isFill: false, type, definition, styleDefinition, styleRegistry is not null, Add);

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
        GanttStyleDefinition? styleDefinition,
        bool styleRegistryProvided,
        Action<string, string, GanttValidationSeverity, string> add)
    {
        if (text is null || type is null || definition is null)
        {
            return null;
        }

        var field = isFill ? "FillColour" : "StrokeColour";
        if (!IsHexColour(text))
        {
            add(field, GanttValidationCodes.BadColourFormat, GanttValidationSeverity.Error, $"{field} '{text}' must be '#RRGGBB'.");
            return null;
        }

        if (type == GanttEntityType.CustomActivity && styleDefinition is null)
        {
            return null;
        }

        // Preserve the pre-R2.9 compatibility path when no registry was
        // supplied. With a registry, Custom Activity resolves strictly from
        // its selected named style.
        if (type == GanttEntityType.CustomActivity && !styleRegistryProvided)
        {
            return text.ToUpperInvariant();
        }

        EntityColourCapability capability = type == GanttEntityType.CustomActivity
            ? styleDefinition?.ColourCapability ?? EntityColourCapability.None
            : definition.ColourCapability;
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
        HashSet<string> duplicateIds,
        List<GanttValidationIssue> issues
    )
    {
        HashSet<int> cycleRows = CheckParentCycles(rows, perRow, canonicalById, duplicateIds, issues);

        // Resolve each Critical Interval ParentId against the same batch.
        for (var i = 0; i < rows.Count; i++)
        {
            ValidatedRow? parsed = perRow[i];
            if (parsed is null
                || !EntityHierarchyCatalog.MayBeChild(parsed.Type ?? GanttEntityType.Spacer)
                || parsed.Event?.ParentId is null)
            {
                continue;
            }

            if (cycleRows.Contains(i))
            {
                continue;
            }

            var key = parsed.Event.ParentId.Value;
            if (duplicateIds.Contains(key))
            {
                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "ParentId",
                        GanttValidationCodes.ParentAmbiguous,
                        GanttValidationSeverity.Error,
                        $"ParentId '{key}' is ambiguous because multiple rows carry that Id."
                    )
                );
                perRow[i] = parsed with { HasBlockingError = true };
                continue;
            }

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

            ValidatedRow? parent = perRow[parentIndex];
            if (parent is not { Event: not null, HasBlockingError: false })
            {
                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "ParentId",
                        GanttValidationCodes.ParentInvalid,
                        GanttValidationSeverity.Error,
                        $"ParentId '{key}' references a row that did not validate as a usable event."
                    )
                );
                perRow[i] = parsed with { HasBlockingError = true };
                continue;
            }

            // R4.7A D9: a parent must be able to OWN children per the matrix, not
            // merely be a Span. A child milestone projecting onto a milestone parent
            // is a legal case (REV5's MilestoneOnParent), so the old Span-only test
            // would have rejected a hierarchy the matrix permits. The matrix is the
            // single authority; this check simply consults it.
            if (!EntityHierarchyCatalog.MayOwnChildren(parent.Type!.Value))
            {
                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "ParentId",
                        GanttValidationCodes.ParentNotSpan,
                        GanttValidationSeverity.Error,
                        $"ParentId '{key}' must reference a type that may own children."
                    )
                );
                perRow[i] = parsed with { HasBlockingError = true };
            }
        }

        // A Critical Interval can itself be the parent of another Critical
        // Interval, and the child is commonly authored above its parent. The
        // loop above therefore decides a child against a parent that has not
        // been decided yet, so the same rows in the opposite order would block
        // the child. Propagate each newly blocked parent to its dependents
        // until no further row changes, which makes the outcome independent of
        // the input order.
        PropagateBlockedParents(rows, perRow, canonicalById, duplicateIds, issues);
    }

    /// <summary>
    /// Blocks every still-undecided child row whose canonical parent row is already
    /// blocked, repeating until no further row is blocked so a chain of dependent
    /// children is covered.
    /// </summary>
    /// <param name="rows">The neutral body rows in table order.</param>
    /// <param name="perRow">The per-row state to update.</param>
    /// <param name="canonicalById">First-canonical row index by Id text.</param>
    /// <param name="duplicateIds">Ids carried by more than one row.</param>
    /// <param name="issues">The row issue sink.</param>
    /// <remarks>
    /// <para>
    /// <b>Every child-capable type, not only Critical Interval.</b> The predicate is
    /// <see cref="EntityHierarchyCatalog.MayBeChild"/> -- the single matrix -- so a
    /// child activity, child milestone or critical milestone is propagated exactly as
    /// a critical interval is. Keying on <c>Type == CriticalInterval</c> meant a child
    /// whose parent had just been refused for an unrelated reason (a bad date, a style
    /// key that does not resolve) survived as a valid event pointing at a parent that
    /// was not in <c>Events</c> at all, and the chart then had to resolve that edge
    /// itself. The matrix is the same authority the direct parent pass consults, so
    /// the two cannot disagree about who is a child.
    /// </para>
    /// <para>
    /// <b>Order independence.</b> The direct pass decides a child against whatever its
    /// parent row currently looks like, and an authoring surface may place a child
    /// above its parent. This pass is what makes the outcome independent of input
    /// order, and it runs again after <see cref="CheckChildCapacityAndDepth"/> so a
    /// parent that only becomes blocked THERE -- over the seven-child cap, or the depth
    /// rule -- propagates to its own children in the same validation.
    /// </para>
    /// </remarks>
    private static void PropagateBlockedParents(
        IReadOnlyList<GanttRowDto> rows,
        ValidatedRow?[] perRow,
        Dictionary<string, int> canonicalById,
        HashSet<string> duplicateIds,
        List<GanttValidationIssue> issues
    )
    {
        // Rows that already carry a blocking error were decided by the field
        // pass, the cycle pass, or the direct parent pass; re-reporting them
        // here would duplicate that finding.
        var decided = new HashSet<int>();
        for (var i = 0; i < rows.Count; i++)
        {
            if (perRow[i] is { HasBlockingError: true }
                && EntityHierarchyCatalog.MayBeChild(perRow[i]!.Type ?? GanttEntityType.Spacer))
            {
                _ = decided.Add(i);
            }
        }

        var changed = true;
        while (changed)
        {
            changed = false;
            for (var i = 0; i < rows.Count; i++)
            {
                if (
                    decided.Contains(i)
                    || perRow[i] is not
                    {
                        HasBlockingError: false,
                        Event.ParentId: { } parentId,
                        Type: { } rowType,
                    }
                    || !EntityHierarchyCatalog.MayBeChild(rowType)
                    || duplicateIds.Contains(parentId.Value)
                    || !canonicalById.TryGetValue(parentId.Value, out var parentIndex)
                    || perRow[parentIndex] is not { HasBlockingError: true }
                )
                {
                    continue;
                }

                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "ParentId",
                        GanttValidationCodes.ParentInvalid,
                        GanttValidationSeverity.Error,
                        $"ParentId '{parentId.Value}' references a row that did not validate as a usable event."
                    )
                );
                perRow[i] = perRow[i]! with { HasBlockingError = true };
                _ = decided.Add(i);
                changed = true;
            }
        }
    }

    /// <summary>
    /// Enforces the two hierarchy limits the product fixes: at most
    /// <see cref="EntityHierarchyCatalog.MaxChildrenPerParent"/> children per
    /// parent, and at most <see cref="EntityHierarchyCatalog.MaxDepth"/> levels of
    /// nesting.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This runs after <c>CheckCriticalParents</c> so a row whose parent is
    /// already known to be ambiguous, unknown or invalid is skipped here rather
    /// than counted as a child of a parent that does not exist. A child of a
    /// broken parent is already blocking; adding a second, unrelated error would
    /// be noise.
    /// </para>
    /// <para>
    /// <b>Determinism.</b> Children are grouped in ascending worksheet row order
    /// and the issue is reported against the parent's row, so the same table
    /// always produces the same findings regardless of enumeration order. Which
    /// children are "the excess" is therefore the ones furthest down the table,
    /// which is also what a user looking at their sheet would expect to be
    /// pointed at.
    /// </para>
    /// </remarks>
    private static void CheckChildCapacityAndDepth(
        IReadOnlyList<GanttRowDto> rows,
        ValidatedRow?[] perRow,
        Dictionary<string, int> canonicalById,
        HashSet<string> duplicateIds,
        List<GanttValidationIssue> issues
    )
    {
        var childrenByParent = new Dictionary<string, List<int>>(StringComparer.Ordinal);

        for (var i = 0; i < rows.Count; i++)
        {
            ValidatedRow? parsed = perRow[i];
            if (parsed is null
                || parsed.HasBlockingError
                || !EntityHierarchyCatalog.MayBeChild(parsed.Type ?? GanttEntityType.Spacer)
                || parsed.Event?.ParentId is not { } parentId
                || duplicateIds.Contains(parentId.Value)
                || !canonicalById.TryGetValue(parentId.Value, out var parentIndex))
            {
                continue;
            }

            if (!childrenByParent.TryGetValue(parentId.Value, out List<int>? siblings))
            {
                siblings = [];
                childrenByParent[parentId.Value] = siblings;
            }
            siblings.Add(i);

            // Depth: a grandchild is refused -- a child of a child. The rule is
            // uniform in the child's type: a parent that is itself a child puts
            // any child of it at depth 3, whichever type that child is. A
            // Critical Interval used to be exempted here, on the grounds that a
            // critical interval may parent another critical interval while
            // remaining a child of an activity. The owner has ruled that
            // Critical Interval is a level-2 child and cannot itself own a child,
            // so the exemption is withdrawn (R4.7 reconciliation, 2026-09-29) and
            // the depth rule now matches MaxDepth, EntityProjection, and the
            // EntityHierarchyCatalog contract comment, which all already said a
            // critical interval's own parent must be top-level.
            //
            // DEPTH IS DECIDED FROM THE PARENT'S ORIGINAL ParentId, never from the
            // parent's evolving HasBlockingError flag. The flag is order-dependent:
            // in a four-level chain authored top-down, the depth rule blocks row 4
            // while this loop is still running, so row 5 was refused by
            // PropagateBlockedParents as ParentInvalid; authored bottom-up, row 5 is
            // reached first, its parent is not yet blocked, and it was refused as
            // HierarchyTooDeep. The same table produced two different findings for
            // the same row purely from the order rows were supplied in.
            //
            // Whether a row is nested is a fact about the declared hierarchy and is
            // true whatever else is wrong, so it is read from the parent's own
            // ParentId. Classifying a row whose PARENT is blocked stays with
            // PropagateBlockedParents, which reports ParentInvalid; this loop no
            // longer guesses at it, and a parent refused for an unrelated fault (a
            // bad date, an unresolvable StyleKey) still yields exactly one finding
            // on its child, because a TOP-LEVEL parent's ParentId is null and so
            // never trips this rule.
            ValidatedRow? parent = perRow[parentIndex];
            if (parent?.Event?.ParentId is not null)
            {
                issues.Add(
                    new GanttValidationIssue(
                        rows[i].RowNumber,
                        "ParentId",
                        GanttValidationCodes.HierarchyTooDeep,
                        GanttValidationSeverity.Error,
                        $"ParentId '{parentId.Value}' is itself a child, which would nest beyond the supported depth of {EntityHierarchyCatalog.MaxDepth}."
                    )
                );
                parsed = parsed with { HasBlockingError = true };
                perRow[i] = parsed;

                // The row is not counted as a healthy child, so it cannot push a
                // parent over the capacity limit with a child that does not exist
                // in a renderable hierarchy.
                siblings.RemoveAt(siblings.Count - 1);
            }
        }

        foreach ((var parentId, List<int> children) in childrenByParent.OrderBy(static p => p.Key, StringComparer.Ordinal))
        {
            if (children.Count <= EntityHierarchyCatalog.MaxChildrenPerParent
                || !canonicalById.TryGetValue(parentId, out var parentRowIndex))
            {
                continue;
            }

            issues.Add(
                new GanttValidationIssue(
                    rows[parentRowIndex].RowNumber,
                    "ParentId",
                    GanttValidationCodes.TooManyChildren,
                    GanttValidationSeverity.Error,
                    $"This row has {children.Count} children; at most {EntityHierarchyCatalog.MaxChildrenPerParent} are permitted."
                )
            );
            perRow[parentRowIndex] = perRow[parentRowIndex]! with { HasBlockingError = true };
        }
    }

    /// <summary>
    /// Rejects self-references and longer directed cycles formed by <c>ParentId</c>
    /// values. Ordinary span parents are terminal nodes and are not traversed.
    /// </summary>
    /// <remarks>
    /// An edge whose target Id is duplicated is ambiguous rather than cyclic:
    /// which row the reference means is unknown, so claiming a cycle through it
    /// would replace the accurate <see cref="GanttValidationCodes.ParentAmbiguous"/>
    /// finding. Such an edge terminates the walk instead.
    /// </remarks>
    private static HashSet<int> CheckParentCycles(
        IReadOnlyList<GanttRowDto> rows,
        ValidatedRow?[] perRow,
        Dictionary<string, int> canonicalById,
        HashSet<string> duplicateIds,
        List<GanttValidationIssue> issues
    )
    {
        var cycleRows = new HashSet<int>();
        var path = new List<int>();
        var pathIndexByRow = new Dictionary<int, int>();

        for (var start = 0; start < rows.Count; start++)
        {
            if (!IsUsableChildCapableRow(start, perRow, canonicalById))
            {
                continue;
            }

            path.Clear();
            pathIndexByRow.Clear();
            var current = start;
            while (current >= 0 && IsUsableChildCapableRow(current, perRow, canonicalById))
            {
                if (pathIndexByRow.TryGetValue(current, out var cycleStart))
                {
                    for (var index = cycleStart; index < path.Count; index++)
                    {
                        _ = cycleRows.Add(path[index]);
                    }

                    break;
                }

                pathIndexByRow[current] = path.Count;
                path.Add(current);
                ValidatedRow? node = perRow[current];
                if (node?.Event?.ParentId is not { } parentId
                    || duplicateIds.Contains(parentId.Value)
                    || !canonicalById.TryGetValue(parentId.Value, out var parentIndex))
                {
                    break;
                }

                current = parentIndex;
            }
        }

        foreach (var rowIndex in cycleRows.OrderBy(index => rows[index].RowNumber))
        {
            issues.Add(
                new GanttValidationIssue(
                    rows[rowIndex].RowNumber,
                    "ParentId",
                    GanttValidationCodes.ParentCycle,
                    GanttValidationSeverity.Error,
                    "ParentId forms a parent cycle: a row names itself or an ancestor as its parent."
                )
            );
            perRow[rowIndex] = perRow[rowIndex]! with { HasBlockingError = true };
        }

        return cycleRows;
    }

    /// <summary>
    /// Returns whether a row is a canonical, error-free, child-capable event that can
    /// participate in the parent graph.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Every child-capable type, not only Critical Interval.</b> The predicate is
    /// <see cref="EntityHierarchyCatalog.MayBeChild"/> — the same single matrix the
    /// direct parent pass and the propagation pass consult, so the three cannot
    /// disagree about who is a child.
    /// </para>
    /// <para>
    /// <b>Why this had to widen.</b> Keying on <c>Type == CriticalInterval</c> meant
    /// a self-reference on a child ACTIVITY was never walked as a cycle. It then
    /// reached the depth rule, whose parent is the row itself and whose
    /// <c>ParentId</c> is non-null, so the user was told the row "is itself a child"
    /// — a depth complaint about a row that is not nested at all. The accurate finding
    /// is that the row names itself, which is what <c>ParentCycle</c> says. A
    /// two-activity cycle behaved the same way: both rows got
    /// <c>HierarchyTooDeep</c> and neither was told the truth.
    /// </para>
    /// </remarks>
    private static bool IsUsableChildCapableRow(int rowIndex, ValidatedRow?[] perRow, Dictionary<string, int> canonicalById)
    {
        ValidatedRow? row = perRow[rowIndex];
        return row is { Event: not null, HasBlockingError: false, Type: { } rowType }
            && EntityHierarchyCatalog.MayBeChild(rowType)
            && row.Id is not null
            && canonicalById.TryGetValue(row.Id.Value, out var canonicalIndex)
            && canonicalIndex == rowIndex;
    }
}
