using System.Collections.Frozen;

namespace GanttCreator.Core;

/// <summary>
/// The immutable contract metadata for one user-selectable row type, exactly
/// as defined by the entity guide. Constructed only by
/// <see cref="EntityTypeCatalog"/>'s code-owned data.
/// </summary>
public sealed class EntityTypeDefinition
{
    /// <summary>
    /// Initialises a definition.
    /// </summary>
    /// <param name="type">The code identity of the type.</param>
    /// <param name="displayName">The exact workbook display name; a schema value.</param>
    /// <param name="kind">The visual-entity class the type renders as.</param>
    /// <param name="dateMode">Which date columns the type reads for geometry.</param>
    /// <param name="defaultStyleKey">
    /// The default named style key, or empty when the row must supply a
    /// <c>StyleKey</c> (<see cref="RequiresStyleKey"/>).
    /// </param>
    /// <param name="colourCapability">Which per-row colour overrides are valid for the type.</param>
    /// <param name="requiresStyleKey">Whether the row must supply a <c>StyleKey</c>.</param>
    /// <param name="allowedLabelPositions">
    /// The permitted label positions, or empty when the permitted subset is
    /// resolved from the row's required named style.
    /// </param>
    public EntityTypeDefinition(
        GanttEntityType type,
        string displayName,
        EntityKind kind,
        EntityDateMode dateMode,
        string defaultStyleKey,
        EntityColourCapability colourCapability,
        bool requiresStyleKey,
        IReadOnlySet<GanttLabelPosition> allowedLabelPositions)
    {
        ArgumentNullException.ThrowIfNull(displayName);
        ArgumentNullException.ThrowIfNull(defaultStyleKey);
        ArgumentNullException.ThrowIfNull(allowedLabelPositions);

        Type = type;
        DisplayName = displayName;
        Kind = kind;
        DateMode = dateMode;
        DefaultStyleKey = defaultStyleKey;
        ColourCapability = colourCapability;
        RequiresStyleKey = requiresStyleKey;
        AllowedLabelPositions = allowedLabelPositions.ToFrozenSet();
    }

    /// <summary>The code identity of the type.</summary>
    public GanttEntityType Type { get; }

    /// <summary>
    /// The exact workbook display name. This value is part of the workbook
    /// schema; renaming or localising it is a schema migration.
    /// </summary>
    public string DisplayName { get; }

    /// <summary>The visual-entity class the type renders as.</summary>
    public EntityKind Kind { get; }

    /// <summary>Which date columns the type reads for geometry.</summary>
    public EntityDateMode DateMode { get; }

    /// <summary>
    /// The default named style key, or empty when the row must supply a
    /// <c>StyleKey</c> (<see cref="RequiresStyleKey"/>).
    /// </summary>
    public string DefaultStyleKey { get; }

    /// <summary>Which per-row colour overrides are valid for the type.</summary>
    public EntityColourCapability ColourCapability { get; }

    /// <summary>Whether the row must supply a <c>StyleKey</c>.</summary>
    public bool RequiresStyleKey { get; }

    /// <summary>
    /// The permitted label positions, or empty when the permitted subset is
    /// resolved from the row's required named style.
    /// </summary>
    public IReadOnlySet<GanttLabelPosition> AllowedLabelPositions { get; }
}
