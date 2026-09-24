namespace GanttCreator.Core;

/// <summary>
/// Builds the canonical scaffold row for a new visible <c>tblGanttData</c>
/// entity. The builder is pure and Office-free; row insertion belongs to the
/// Office adapter.
/// </summary>
public static class GanttRowDefaults
{
    /// <summary>
    /// Builds the 14 cells in <see cref="GanttTableSchema.Default"/> order.
    /// Only Id, Type, and the Type's default StyleKey are populated; all other
    /// fields are blank so validation can report the required authoring inputs.
    /// </summary>
    /// <param name="type">The catalogue type used by the command.</param>
    /// <param name="nextId">The one-shot stable row-ID generator.</param>
    /// <returns>The exact 14-cell scaffold values.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="nextId"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="type"/> is not in the catalogue.</exception>
    public static IReadOnlyList<object?> Build(GanttEntityType type, Func<GanttRowId> nextId)
    {
        ArgumentNullException.ThrowIfNull(nextId);
        EntityTypeDefinition definition = EntityTypeCatalog.GetDefinition(type)
            ?? throw new ArgumentOutOfRangeException(nameof(type), type, "Type is not in the catalogue.");

        return
        [
            nextId().Value,
            null,
            null,
            definition.DisplayName,
            null,
            null,
            null,
            null,
            definition.DefaultStyleKey,
            null,
            null,
            null,
            null,
            null,
        ];
    }
}
