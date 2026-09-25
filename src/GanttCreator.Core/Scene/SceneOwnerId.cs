namespace GanttCreator.Core.Scene;

/// <summary>The closed ownership kind for a scene entity.</summary>
public enum SceneOwnerKind
{
    /// <summary>The entity belongs to one visible worksheet row.</summary>
    Row = 0,

    /// <summary>The entity belongs to the chart/workbook scene.</summary>
    Chart = 1,
}

/// <summary>An immutable scene owner that is either one worksheet row or the reserved chart owner.</summary>
public sealed record SceneOwnerId
{
    /// <summary>The exact reserved value used by the chart owner.</summary>
    public const string ChartValue = "chart";

    private SceneOwnerId(SceneOwnerKind kind, string value)
    {
        Kind = kind;
        Value = value;
    }

    /// <summary>Gets the ownership kind.</summary>
    public SceneOwnerKind Kind { get; }

    /// <summary>Gets the exact owner value: the row ID text or <see cref="ChartValue"/>.</summary>
    public string Value { get; }

    /// <summary>Gets the single reserved chart-level owner.</summary>
    public static SceneOwnerId Chart { get; } = new(SceneOwnerKind.Chart, ChartValue);

    /// <summary>Creates a row owner from a stable worksheet row ID.</summary>
    /// <param name="rowId">The stable row identifier.</param>
    /// <returns>The row scene owner.</returns>
    public static SceneOwnerId ForRow(GanttRowId rowId)
    {
        ArgumentNullException.ThrowIfNull(rowId);
        return new(SceneOwnerKind.Row, rowId.Value);
    }

    /// <summary>Attempts to parse a serialized owner kind and value.</summary>
    /// <param name="kind">The exact owner kind text.</param>
    /// <param name="value">The exact owner value.</param>
    /// <param name="owner">The parsed owner when successful.</param>
    /// <returns><see langword="true"/> when the owner is valid.</returns>
    public static bool TryParse(string? kind, string? value, out SceneOwnerId? owner)
    {
        owner = null;
        if (string.Equals(kind, "row", StringComparison.Ordinal) && GanttRowId.TryParse(value, out GanttRowId? rowId) && rowId is not null)
        {
            owner = ForRow(rowId);
            return true;
        }

        if (string.Equals(kind, "chart", StringComparison.Ordinal) && string.Equals(value, ChartValue, StringComparison.Ordinal))
        {
            owner = Chart;
            return true;
        }

        return false;
    }
}
