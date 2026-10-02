namespace GanttCreator.Core;

/// <summary>
/// Builds the canonical scaffold row for a new visible <c>tblGanttData</c>
/// entity. The builder is pure and Office-free; row insertion belongs to the
/// Office adapter.
/// </summary>
public static class GanttRowDefaults
{
    /// <summary>
    /// Builds one blank cell per column of <see cref="GanttTableSchema.Default"/>,
    /// in schema order, with only Id, Type, and the Type's default StyleKey
    /// populated. Every other field is blank so validation can report the required
    /// authoring inputs.
    /// </summary>
    /// <param name="type">The catalogue type used by the command.</param>
    /// <param name="nextId">The one-shot stable row-ID generator.</param>
    /// <returns>The scaffold values, one per schema column and in schema order.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="nextId"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="type"/> is not in the catalogue.</exception>
    /// <remarks>
    /// <para>
    /// The value is placed by column NAME, never at a hard-coded position. An
    /// earlier version returned a 14-element literal, so when R4.7A inserted
    /// <c>SiblingOrder</c> the schema grew to 15 columns and the literal did not:
    /// the default StyleKey at literal index 8 still sat where <c>ParentId</c> had
    /// been, which after the insertion is <c>SiblingOrder</c>. Every newly
    /// inserted row would have carried its style key in the sibling-ordering
    /// column. The header-mapping contract test is what surfaced it.
    /// </para>
    /// <para>
    /// Building by name means a future column can be added without this literal
    /// drifting again, and a value for a column the schema no longer has fails
    /// visibly rather than being written into the wrong cell.
    /// </para>
    /// </remarks>
    public static IReadOnlyList<object?> Build(GanttEntityType type, Func<GanttRowId> nextId)
    {
        ArgumentNullException.ThrowIfNull(nextId);
        EntityTypeDefinition definition = EntityTypeCatalog.GetDefinition(type)
            ?? throw new ArgumentOutOfRangeException(nameof(type), type, "Type is not in the catalogue.");

        Dictionary<string, object?> values = new(StringComparer.Ordinal)
        {
            ["Id"] = nextId().Value,
            ["Type"] = definition.DisplayName,
            ["StyleKey"] = definition.DefaultStyleKey,
        };

        return [.. GanttTableSchema.Default.Columns.Select(c => values.GetValueOrDefault(c.Name))];
    }
}
