namespace GanttCreator.Core;

/// <summary>
/// The single code-owned catalogue of user-selectable <c>Type</c> values for
/// the <c>Type</c> column of the visible <c>tblGanttData</c> table.
/// </summary>
/// <remarks>
/// <para>
/// This catalogue is the authoritative definition of the 16 selectable Types;
/// the entity guide's broader visual/entity sections (chart frame through
/// validation indicators) are not additional workbook Type values. No other
/// component maintains its own type list.
/// </para>
/// <para>
/// The exact workbook display names form part of the workbook schema:
/// renaming or localising a display name is a schema migration that bumps
/// <see cref="GanttSchemaVersion.CurrentSchemaVersion"/>. Such a migration
/// does not rename or renumber the durable <see cref="GanttEntityType"/>
/// identity. The catalogue is materialised deterministically to
/// <c>tblGanttTypes</c> on
/// <c>_GanttCreatorConfig</c> (R2.9); the domain parser, worksheet validation,
/// Ribbon controls, and tests all use these same definitions.
/// </para>
/// </remarks>
public static class EntityTypeCatalog
{
    /// <summary>The number of types in the catalogue.</summary>
    public const int TypeCount = 16;

    private static readonly IReadOnlySet<GanttLabelPosition> _spanLabelPositions =
        new HashSet<GanttLabelPosition>(
        [
            GanttLabelPosition.Auto,
            GanttLabelPosition.Left,
            GanttLabelPosition.Right,
            GanttLabelPosition.Inside,
            GanttLabelPosition.None,
        ]);

    private static readonly IReadOnlySet<GanttLabelPosition> _milestoneLabelPositions =
        new HashSet<GanttLabelPosition>(
        [
            GanttLabelPosition.Auto,
            GanttLabelPosition.Left,
            GanttLabelPosition.Right,
            GanttLabelPosition.None,
        ]);

    private static readonly IReadOnlySet<GanttLabelPosition> _delineatorLabelPositions =
        new HashSet<GanttLabelPosition>(
        [
            GanttLabelPosition.Auto,
            GanttLabelPosition.TopLeft,
            GanttLabelPosition.TopRight,
            GanttLabelPosition.BottomLeft,
            GanttLabelPosition.BottomRight,
            GanttLabelPosition.None,
        ]);

    private static readonly IReadOnlySet<GanttLabelPosition> _splitterLabelPositions =
        new HashSet<GanttLabelPosition>(
        [
            GanttLabelPosition.DataPanelLeft,
            GanttLabelPosition.PlotCentre,
            GanttLabelPosition.Both,
            GanttLabelPosition.None,
        ]);

    private static readonly IReadOnlySet<GanttLabelPosition> _noLabelPositions = new HashSet<GanttLabelPosition>();

    /// <summary>Returns a single-element position set for types that permit exactly one position.</summary>
    private static HashSet<GanttLabelPosition> OnlyLabelPositions(GanttLabelPosition position) =>
        new([position]);

    private static readonly EntityTypeDefinition[] _entries =
    [
        new(
            GanttEntityType.Splitter,
            "Splitter",
            EntityKind.SectionHeader,
            EntityDateMode.None,
            "Splitter",
            EntityColourCapability.Fill,
            requiresStyleKey: false,
            _splitterLabelPositions),
        new(
            GanttEntityType.Spacer,
            "Spacer",
            EntityKind.Spacer,
            EntityDateMode.None,
            "Spacer",
            EntityColourCapability.None,
            requiresStyleKey: false,
            OnlyLabelPositions(GanttLabelPosition.None)),
        new(
            GanttEntityType.AsBuiltActivity,
            "As-Built Activity",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            "AsBuiltActivity",
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _spanLabelPositions),
        new(
            GanttEntityType.AsPlannedActivity,
            "As-Planned Activity",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            "AsPlannedActivity",
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _spanLabelPositions),
        new(
            GanttEntityType.BaselineActivity,
            "Baseline Activity",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            "BaselineActivity",
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _spanLabelPositions),
        new(
            GanttEntityType.CriticalInterval,
            "Critical Interval",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            "CriticalInterval",
            // ADR-0027 D5: the entity is a FILLED rectangle, so a user FillColour
            // override is now permitted. This flip, the validator's acceptance of
            // that override, and checklist §C land together (D6) — a style that
            // fills while the validator refuses the user's own fill is a defect
            // that looks like a rendering choice.
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            OnlyLabelPositions(GanttLabelPosition.None)),
        new(
            GanttEntityType.DelayEvent,
            "Delay Event",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            "DelayEvent",
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _spanLabelPositions),
        new(
            GanttEntityType.AsBuiltProcurement,
            "As-Built Procurement",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            "AsBuiltProcurement",
            EntityColourCapability.Hatch | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _spanLabelPositions),
        new(
            GanttEntityType.AsPlannedProcurement,
            "As-Planned Procurement",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            "AsPlannedProcurement",
            EntityColourCapability.Hatch | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _spanLabelPositions),
        new(
            GanttEntityType.BaselineProcurement,
            "Baseline Procurement",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            "BaselineProcurement",
            EntityColourCapability.Hatch | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _spanLabelPositions),
        new(
            GanttEntityType.CustomActivity,
            "Custom Activity",
            EntityKind.Span,
            EntityDateMode.StartFinish,
            string.Empty,
            EntityColourCapability.StyleDefined,
            requiresStyleKey: true,
            _noLabelPositions),
        new(
            GanttEntityType.AsBuiltMilestone,
            "As-Built Milestone",
            EntityKind.Milestone,
            EntityDateMode.StartOnly,
            "AsBuiltMilestone",
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _milestoneLabelPositions),
        new(
            GanttEntityType.AsPlannedMilestone,
            "As-Planned Milestone",
            EntityKind.Milestone,
            EntityDateMode.StartOnly,
            "AsPlannedMilestone",
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _milestoneLabelPositions),
        new(
            GanttEntityType.BaselineMilestone,
            "Baseline Milestone",
            EntityKind.Milestone,
            EntityDateMode.StartOnly,
            "BaselineMilestone",
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _milestoneLabelPositions),
        new(
            GanttEntityType.CriticalMilestone,
            "Critical Milestone",
            EntityKind.Milestone,
            EntityDateMode.StartOnly,
            "CriticalMilestone",
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _milestoneLabelPositions),
        new(
            GanttEntityType.Delineator,
            "Delineator",
            EntityKind.Delineator,
            EntityDateMode.StartOnly,
            "DefaultDelineator",
            EntityColourCapability.Stroke,
            requiresStyleKey: false,
            _delineatorLabelPositions),
    ];

    /// <summary>Gets the catalogue entries in the exact entity-guide order.</summary>
    public static IReadOnlyList<EntityTypeDefinition> Entries { get; } = Array.AsReadOnly(_entries);

    /// <summary>
    /// Returns the definition for a type, or <see langword="null"/> when the
    /// type is not in the catalogue.
    /// </summary>
    /// <param name="type">The entity type to look up.</param>
    public static EntityTypeDefinition? GetDefinition(GanttEntityType type) =>
        _entries.FirstOrDefault(e => e.Type == type);

    /// <summary>
    /// Returns the definition for an exact display name (Ordinal comparison
    /// on trimmed input), or <see langword="null"/> when no entry matches.
    /// </summary>
    /// <param name="displayName">The workbook display name to look up.</param>
    public static EntityTypeDefinition? GetDefinition(string? displayName)
    {
        if (displayName is null)
        {
            return null;
        }

        var trimmed = displayName.Trim();

        return trimmed.Length == 0
            ? null
            : _entries.FirstOrDefault(e => string.Equals(e.DisplayName, trimmed, StringComparison.Ordinal));
    }

    /// <summary>
    /// Formats a type into its exact workbook display name.
    /// </summary>
    /// <param name="type">The entity type to format.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the type is not in the catalogue.</exception>
    public static string Format(GanttEntityType type) =>
        GetDefinition(type)?.DisplayName
            ?? throw new ArgumentOutOfRangeException(nameof(type), type, "Type is not in the catalogue.");

    /// <summary>
    /// Parses a workbook display name into a type. Matching is Ordinal on
    /// trimmed input: no case-folding, no whitespace rewriting, no guessing.
    /// Never throws.
    /// </summary>
    /// <param name="displayName">The workbook text to parse.</param>
    /// <param name="type">The parsed type when the method returns <see langword="true"/>.</param>
    /// <returns><see langword="true"/> when the text matches exactly one catalogue display name.</returns>
    public static bool TryParse(string? displayName, out GanttEntityType type)
    {
        EntityTypeDefinition? definition = GetDefinition(displayName);
        if (definition is not null)
        {
            type = definition.Type;
            return true;
        }

        type = default;
        return false;
    }

    /// <summary>
    /// The ribbon item identifier that clears the row's <c>LabelPosition</c>
    /// cell (blank), resolving to the style default / <c>Auto</c> cascade per
    /// entity-guide §22. It is the absence of a value, not a catalogue
    /// position, so it never appears in <see cref="GetAllowedLabelPositions"/>.
    /// </summary>
    public const string DefaultLabelSelectionId = "Default";

    /// <summary>
    /// Display order for ribbon label options: the catalogue set in a stable,
    /// deterministic order (Ordinal by enum name would scatter the UX, so the
    /// order is fixed here once). Every element comes from
    /// <see cref="EntityTypeDefinition.AllowedLabelPositions"/> or the row's
    /// style registry set — never a second authority.
    /// </summary>
    private static readonly GanttLabelPosition[] _displayOrder =
    [
        GanttLabelPosition.Auto,
        GanttLabelPosition.Left,
        GanttLabelPosition.Right,
        GanttLabelPosition.Inside,
        GanttLabelPosition.TopLeft,
        GanttLabelPosition.TopRight,
        GanttLabelPosition.BottomLeft,
        GanttLabelPosition.BottomRight,
        GanttLabelPosition.DataPanelLeft,
        GanttLabelPosition.PlotCentre,
        GanttLabelPosition.Both,
        GanttLabelPosition.None,
    ];

    /// <summary>
    /// Returns the allowed label positions for a type in stable ribbon display
    /// order (R5.5 D1 projection). <see cref="GanttEntityType.CustomActivity"/>
    /// resolves from <paramref name="styleRegistry"/> via its
    /// <c>StyleKey</c>; every other type reads its catalogue definition.
    /// </summary>
    /// <param name="type">The entity type.</param>
    /// <param name="styleKey">The row's style key (Custom Activity only).</param>
    /// <param name="styleRegistry">The named-style capabilities, or null.</param>
    /// <returns>The allowed positions in display order; empty when unknown.</returns>
    public static IReadOnlyList<GanttLabelPosition> GetAllowedLabelPositions(
        GanttEntityType type,
        string? styleKey = null,
        GanttStyleRegistry? styleRegistry = null)
    {
        IReadOnlySet<GanttLabelPosition> allowed;
        if (type == GanttEntityType.CustomActivity)
        {
            if (styleRegistry is null
                || !styleRegistry.TryGet(styleKey, out GanttStyleDefinition? style)
                || style is null
                || style.AllowedLabelPositions is null)
            {
                return [];
            }

            allowed = style.AllowedLabelPositions;
        }
        else if (GetDefinition(type) is not { } definition)
        {
            return [];
        }
        else
        {
            allowed = definition.AllowedLabelPositions;
        }

        return [.. _displayOrder.Where(allowed.Contains)];
    }

    /// <summary>
    /// Returns whether a label position is allowed for a type (R5.5 D1 guard).
    /// <see cref="GanttEntityType.CustomActivity"/> delegates to its style.
    /// </summary>
    /// <param name="type">The entity type.</param>
    /// <param name="position">The candidate position.</param>
    /// <param name="styleKey">The row's style key (Custom Activity only).</param>
    /// <param name="styleRegistry">The named-style capabilities, or null.</param>
    public static bool IsLabelPositionAllowed(
        GanttEntityType type,
        GanttLabelPosition position,
        string? styleKey = null,
        GanttStyleRegistry? styleRegistry = null) =>
        GetAllowedLabelPositions(type, styleKey, styleRegistry).Contains(position);

    /// <summary>
    /// Parses a ribbon label-selection identifier. <c>Default</c> parses to a
    /// null position (clear the cell to blank); otherwise the identifier must
    /// name a defined <see cref="GanttLabelPosition"/> Ordinally. Unknown text
    /// is refused, never defaulted (the TimeScale-selection precedent).
    /// </summary>
    /// <param name="selectionId">The ribbon item identifier.</param>
    /// <param name="position">The parsed position, or null for Default.</param>
    /// <returns>True when the identifier is Default or a defined position name.</returns>
    public static bool TryParseLabelSelection(string? selectionId, out GanttLabelPosition? position)
    {
        if (string.Equals(selectionId, DefaultLabelSelectionId, StringComparison.Ordinal))
        {
            position = null;
            return true;
        }

        if (Enum.TryParse(selectionId, out GanttLabelPosition parsed)
            && Enum.IsDefined(parsed))
        {
            position = parsed;
            return true;
        }

        position = null;
        return false;
    }

    /// <summary>
    /// The ribbon item identifier that clears the row's
    /// <c>DateLabelPosition</c> cell (blank), resolving to
    /// <c>Auto</c> per R5.14 D3. It is the absence of a value,
    /// not a catalogue choice, so it never appears in
    /// <see cref="GetAllowedDateLabelPositions"/>.
    /// </summary>
    public const string DefaultDateLabelSelectionId = "Default";

    /// <summary>
    /// Display order for ribbon date-label options: <c>Auto</c>
    /// first, then the two single-date choices in date order. Every
    /// element is derived from the row type's
    /// <see cref="EntityDateMode"/> — never a second authority.
    /// </summary>
    private static readonly GanttDateLabelPosition[] _dateLabelDisplayOrder =
    [
        GanttDateLabelPosition.Auto,
        GanttDateLabelPosition.Start,
        GanttDateLabelPosition.Finish,
    ];

    /// <summary>
    /// Returns the allowed date-label choices for a type in stable
    /// ribbon display order (R5.14 D1 projection), derived from the
    /// type's <see cref="EntityDateMode"/>: a dateless type names
    /// none (the control is disabled), a <c>StartOnly</c> type names
    /// <c>Auto</c> and <c>Start</c>, and a span names all three.
    /// </summary>
    /// <param name="type">The entity type.</param>
    /// <returns>The allowed choices in display order; empty for a dateless or unknown type.</returns>
    public static IReadOnlyList<GanttDateLabelPosition> GetAllowedDateLabelPositions(
        GanttEntityType type)
    {
        EntityTypeDefinition? definition = GetDefinition(type);
        return definition is null
            ? []
            : definition.DateMode switch
            {
                EntityDateMode.None => [],
                EntityDateMode.StartOnly
                    => [.. _dateLabelDisplayOrder.Where(p => p != GanttDateLabelPosition.Finish)],
                EntityDateMode.StartFinish => [.. _dateLabelDisplayOrder],
                // An out-of-range date mode is a catalogue corruption; the
                // safe default is to treat the type as dateless (no options).
                _ => [],
            };
    }

    /// <summary>
    /// Returns whether a date-label choice is allowed for a type
    /// (R5.14 D1/D4 guard). The capability is type-only: a
    /// <c>Custom Activity</c> style carries no date mode, so it
    /// cannot narrow the set the way it narrows
    /// <see cref="GanttLabelPosition"/>.
    /// </summary>
    /// <param name="type">The entity type.</param>
    /// <param name="position">The candidate choice.</param>
    public static bool IsDateLabelPositionAllowed(
        GanttEntityType type, GanttDateLabelPosition position) =>
        GetAllowedDateLabelPositions(type).Contains(position);

    /// <summary>
    /// Parses a ribbon date-label-selection identifier. <c>Default</c>
    /// parses to a null choice (clear the cell to blank = <c>Auto</c>);
    /// otherwise the identifier must name a defined
    /// <see cref="GanttDateLabelPosition"/> Ordinally. Unknown text is
    /// refused, never defaulted (the label-selection precedent).
    /// </summary>
    /// <param name="selectionId">The ribbon item identifier.</param>
    /// <param name="position">The parsed choice, or null for Default.</param>
    /// <returns>True when the identifier is Default or a defined choice name.</returns>
    public static bool TryParseDateLabelSelection(
        string? selectionId, out GanttDateLabelPosition? position)
    {
        if (string.Equals(selectionId, DefaultDateLabelSelectionId, StringComparison.Ordinal))
        {
            position = null;
            return true;
        }

        if (Enum.TryParse(selectionId, out GanttDateLabelPosition parsed)
            && Enum.IsDefined(parsed))
        {
            position = parsed;
            return true;
        }

        position = null;
        return false;
    }
}
