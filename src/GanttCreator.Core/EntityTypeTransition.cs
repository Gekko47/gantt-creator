namespace GanttCreator.Core;

/// <summary>
/// One row's type-dependent cell text, exactly as the sheet stores it (R5.23).
/// </summary>
/// <param name="StyleKey">The raw <c>StyleKey</c> text, or null/blank.</param>
/// <param name="LabelPosition">The raw <c>LabelPosition</c> text, or null/blank.</param>
/// <param name="DateLabelPosition">The raw <c>DateLabelPosition</c> text, or null/blank.</param>
/// <param name="FillColour">The raw <c>FillColour</c> text, or null/blank.</param>
/// <param name="StrokeColour">The raw <c>StrokeColour</c> text, or null/blank.</param>
/// <param name="ParentId">The raw <c>ParentId</c> text, or null/blank.</param>
/// <param name="LaneId">The raw <c>LaneId</c> text, or null/blank.</param>
/// <param name="StackIndex">The raw <c>StackIndex</c> text, or null/blank.</param>
/// <param name="SiblingOrder">The raw <c>SiblingOrder</c> text, or null/blank.</param>
public sealed record TypeTransitionRow(
    string? StyleKey,
    string? LabelPosition,
    string? DateLabelPosition,
    string? FillColour,
    string? StrokeColour,
    string? ParentId,
    string? LaneId,
    string? StackIndex,
    string? SiblingOrder);

/// <summary>One cell the type transition rewrites (R5.23).</summary>
/// <param name="Column">The schema column name.</param>
/// <param name="NewValue">The value to write; empty clears the cell.</param>
/// <param name="Reason">The stable short reason, for the technical record.</param>
public sealed record TypeTransitionChange(string Column, string NewValue, string Reason);

/// <summary>
/// The single Core authority for the type-dependent cells a <c>Type</c> change must
/// repair (R5.23). Pure and Office-free; the adapter writes what this returns.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why it exists.</b> The old adapter re-implemented the catalogue's rules column by
/// column and drifted from the validator and the ribbon: a style key was never
/// re-evaluated, a delineator's corner vocabulary was only half understood, and the
/// outcome depended on which projection the reader happened to use. Every rule here is
/// expressed through <see cref="EntityTypeCatalog"/>, <see cref="EntityHierarchyCatalog"/>
/// and the style registry, which are the same authorities the validator and the ribbon
/// consume.
/// </para>
/// <para>
/// <b>Old type is knowledge, not data.</b> The sheet no longer holds the previous
/// <c>Type</c>, so a caller that knows it (the change-path heal, from the cached
/// pre-edit rows) passes <c>oldType</c>. A null old type runs only the idempotent
/// same-type rules, which is what the every-pass sweep can safely do: a valid stored
/// corner on a delineator is never wiped by an ordinary Refresh.
/// </para>
/// <para>
/// <b>Nothing unparseable is deleted.</b> A token that names no member of either label
/// vocabulary is left for the validator, which reports it. Only a value that is
/// type-dependent is rewritten.
/// </para>
/// </remarks>
public static class EntityTypeTransition
{
    /// <summary>The <c>StyleKey</c> column name.</summary>
    public const string StyleKeyColumn = "StyleKey";

    /// <summary>The <c>LabelPosition</c> column name.</summary>
    public const string LabelPositionColumn = "LabelPosition";

    /// <summary>The <c>DateLabelPosition</c> column name.</summary>
    public const string DateLabelPositionColumn = "DateLabelPosition";

    /// <summary>The <c>FillColour</c> column name.</summary>
    public const string FillColourColumn = "FillColour";

    /// <summary>The <c>StrokeColour</c> column name.</summary>
    public const string StrokeColourColumn = "StrokeColour";

    /// <summary>The <c>ParentId</c> column name.</summary>
    public const string ParentIdColumn = "ParentId";

    /// <summary>The <c>LaneId</c> column name.</summary>
    public const string LaneIdColumn = "LaneId";

    /// <summary>The <c>StackIndex</c> column name.</summary>
    public const string StackIndexColumn = "StackIndex";

    /// <summary>The <c>SiblingOrder</c> column name.</summary>
    public const string SiblingOrderColumn = "SiblingOrder";

    /// <summary>
    /// Computes the cells a row must change to be valid for <paramref name="newType"/>.
    /// </summary>
    /// <param name="row">The row's raw type-dependent cell text.</param>
    /// <param name="newType">The type the row now carries.</param>
    /// <param name="oldType">
    /// The type the row carried before the edit, when the caller knows it; null runs
    /// the idempotent same-type rules only.
    /// </param>
    /// <param name="styles">
    /// The workbook's named styles, or null when the caller has no catalogue.
    /// </param>
    /// <returns>
    /// The cells to rewrite, in a deterministic column order. An empty list means the
    /// row already matches its type.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="row"/> is null.</exception>
    public static IReadOnlyList<TypeTransitionChange> Apply(
        TypeTransitionRow row,
        GanttEntityType newType,
        GanttEntityType? oldType,
        GanttStyleRegistry? styles)
    {
        ArgumentNullException.ThrowIfNull(row);

        EntityTypeDefinition? definition = EntityTypeCatalog.GetDefinition(newType);
        if (definition is null)
        {
            return [];
        }

        List<TypeTransitionChange> changes = [];

        string? styleKey = NormaliseStyleKey(row, definition, newType, styles, changes);
        NormaliseLabelPosition(row, newType, styleKey, styles, changes);
        NormaliseDateLabelPosition(row, newType, oldType, changes);
        NormaliseColour(row.FillColour, EntityColourCapability.Fill, definition,
            FillColourColumn, changes);
        NormaliseColour(row.StrokeColour, EntityColourCapability.Stroke, definition,
            StrokeColourColumn, changes);
        NormaliseHierarchy(row, newType, changes);
        NormaliseLaneFields(row, definition, newType, changes);

        return changes;
    }

    private static string? NormaliseStyleKey(
        TypeTransitionRow row,
        EntityTypeDefinition definition,
        GanttEntityType newType,
        GanttStyleRegistry? styles,
        List<TypeTransitionChange> changes)
    {
        // A type that does not require a key is StyleKeyWritePlanner's business: it
        // already rewrites the cell to the type default on every pass. Doing it here as
        // well would be a second authority and would clear a choice the user may want
        // back if they convert the row again.
        if (!definition.RequiresStyleKey)
        {
            return Normalised(row.StyleKey);
        }

        string? current = Normalised(row.StyleKey);
        bool usable = current is not null
            && EntityTypeCatalog.IsStyleKeyUsable(newType, current, styles);
        if (usable)
        {
            return current;
        }

        string fallback = definition.DefaultStyleKey;
        if (string.Equals(current, fallback, StringComparison.Ordinal))
        {
            return fallback;
        }

        changes.Add(new TypeTransitionChange(
            StyleKeyColumn,
            fallback,
            current is null ? "style-default" : "style-not-usable"));
        return fallback;
    }

    private static void NormaliseLabelPosition(
        TypeTransitionRow row,
        GanttEntityType newType,
        string? styleKey,
        GanttStyleRegistry? styles,
        List<TypeTransitionChange> changes)
    {
        string? text = Normalised(row.LabelPosition);
        if (text is null
            || !Enum.TryParse(text, ignoreCase: false, out GanttLabelPosition position)
            || !Enum.IsDefined(position))
        {
            // Unparseable text is the validator's to report, not ours to delete.
            return;
        }

        if (!EntityTypeCatalog.IsLabelPositionAllowed(newType, position, styleKey, styles))
        {
            changes.Add(new TypeTransitionChange(
                LabelPositionColumn,
                string.Empty,
                "label-not-allowed"));
        }
    }

    private static void NormaliseDateLabelPosition(
        TypeTransitionRow row,
        GanttEntityType newType,
        GanttEntityType? oldType,
        List<TypeTransitionChange> changes)
    {
        string? text = Normalised(row.DateLabelPosition);
        if (text is null)
        {
            return;
        }

        bool newIsDelineator = newType == GanttEntityType.Delineator;
        bool hasWhichDateVocabulary =
            EntityTypeCatalog.GetAllowedDateLabelPositions(newType).Count > 0;
        bool hasPlacementVocabulary =
            EntityTypeCatalog.GetAllowedDateLabelPlacements(newType).Count > 0;

        bool whichParsed = Enum.TryParse(text, ignoreCase: false,
                out GanttDateLabelPosition whichDate)
            && Enum.IsDefined(whichDate);
        bool placementParsed = Enum.TryParse(text, ignoreCase: false,
                out GanttLabelPosition placement)
            && Enum.IsDefined(placement);

        // R6: a token that names no member of either vocabulary is the validator's to
        // report, never ours to delete. Only a recognised value is type-dependent.
        bool recognised = whichParsed || placementParsed;

        // A dateless type names neither vocabulary, so any recognised value is stale.
        bool dateless = !hasWhichDateVocabulary && !hasPlacementVocabulary && recognised;

        // R5.23 R3: the two vocabularies share Auto and None, so their meaning changes
        // silently across this boundary. The only deterministic answer is to blank the
        // cell.
        bool oldWasDelineator = oldType == GanttEntityType.Delineator;
        bool vocabularyChanged = oldType is not null
            && oldWasDelineator != newIsDelineator
            && recognised;

        bool staleForNewType = newIsDelineator
            ? (placementParsed
                    && !EntityTypeCatalog.IsDateLabelPlacementAllowed(newType, placement))
                || (whichParsed && !placementParsed)
            : (whichParsed
                    && !EntityTypeCatalog.IsDateLabelPositionAllowed(newType, whichDate))
                || (placementParsed && !whichParsed);

        if (dateless || vocabularyChanged || staleForNewType)
        {
            changes.Add(new TypeTransitionChange(
                DateLabelPositionColumn,
                string.Empty,
                "date-label-vocabulary"));
        }
    }

    private static void NormaliseColour(
        string? raw,
        EntityColourCapability required,
        EntityTypeDefinition definition,
        string column,
        List<TypeTransitionChange> changes)
    {
        if (Normalised(raw) is null || (definition.ColourCapability & required) != 0)
        {
            return;
        }

        changes.Add(new TypeTransitionChange(column, string.Empty, "colour-not-supported"));
    }

    private static void NormaliseHierarchy(
        TypeTransitionRow row,
        GanttEntityType newType,
        List<TypeTransitionChange> changes)
    {
        if (Normalised(row.ParentId) is not null
            && !EntityHierarchyCatalog.ParentIdIsRelevant(newType))
        {
            changes.Add(new TypeTransitionChange(
                ParentIdColumn,
                string.Empty,
                "parent-not-supported"));
        }

        if (Normalised(row.SiblingOrder) is not null
            && !EntityHierarchyCatalog.MayBeChild(newType))
        {
            changes.Add(new TypeTransitionChange(
                SiblingOrderColumn,
                string.Empty,
                "sibling-order-not-supported"));
        }
    }

    private static void NormaliseLaneFields(
        TypeTransitionRow row,
        EntityTypeDefinition definition,
        GanttEntityType newType,
        List<TypeTransitionChange> changes)
    {
        bool usesLane = definition.Kind is EntityKind.Span or EntityKind.Milestone;

        // Guide §24: a delineator keeps a populated Finish, LaneId and StackIndex with a
        // non-blocking warning, so data is never silently deleted. Splitter and Spacer
        // have no such rule, so the existing clearing stays for them.
        bool retainsLane = newType == GanttEntityType.Delineator;
        if (usesLane || retainsLane)
        {
            return;
        }

        if (Normalised(row.LaneId) is not null)
        {
            changes.Add(new TypeTransitionChange(
                LaneIdColumn,
                string.Empty,
                "lane-not-supported"));
        }

        if (Normalised(row.StackIndex) is not null)
        {
            changes.Add(new TypeTransitionChange(
                StackIndexColumn,
                string.Empty,
                "stack-not-supported"));
        }
    }

    private static string? Normalised(string? text) =>
        string.IsNullOrWhiteSpace(text) ? null : text.Trim();
}
