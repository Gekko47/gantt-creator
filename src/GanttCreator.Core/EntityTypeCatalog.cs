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
            GanttLabelPosition.Above,
            GanttLabelPosition.Below,
            GanttLabelPosition.None,
        ]);

    private static readonly IReadOnlySet<GanttLabelPosition> _milestoneLabelPositions =
        new HashSet<GanttLabelPosition>(
        [
            GanttLabelPosition.Auto,
            GanttLabelPosition.Left,
            GanttLabelPosition.Right,
            GanttLabelPosition.Above,
            GanttLabelPosition.Below,
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
            EntityColourCapability.Stroke,
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
}
