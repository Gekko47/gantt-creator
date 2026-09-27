namespace GanttCreator.Core.Scene;

/// <summary>The closed ownership kind for a scene entity.</summary>
public enum SceneOwnerKind
{
    /// <summary>The entity belongs to one visible worksheet row.</summary>
    Row = 0,

    /// <summary>The entity belongs to the chart/workbook scene.</summary>
    Chart = 1,

    /// <summary>
    /// The entity is shared by two or more visible worksheet rows because
    /// they deduplicated to one primitive (ADR-0017 D1).
    /// </summary>
    Rows = 2,
}

/// <summary>
/// An immutable scene owner: one worksheet row, the reserved chart owner, or a
/// canonical shared set of rows.
/// </summary>
public sealed record SceneOwnerId
{
    /// <summary>The exact reserved value used by the chart owner.</summary>
    public const string ChartValue = "chart";

    /// <summary>
    /// The separator between row IDs in a shared owner's canonical value
    /// (ADR-0017 D2). A <see cref="GanttRowId"/> is <c>G-</c> plus 32 lowercase
    /// hexadecimal digits, so this character can never occur inside an ID, and
    /// the separator matches the one ADR-0007 D6 uses for the catalogue hash.
    /// </summary>
    public const string RowsSeparator = "|";

    private SceneOwnerId(SceneOwnerKind kind, string value)
    {
        Kind = kind;
        Value = value;
    }

    /// <summary>Gets the ownership kind.</summary>
    public SceneOwnerKind Kind { get; }

    /// <summary>
    /// Gets the exact owner value: the row ID text, the reserved
    /// <see cref="ChartValue"/>, or the canonical <see cref="RowsSeparator"/>-joined
    /// set for a shared owner.
    /// </summary>
    /// <remarks>
    /// This is the owner's entire state. <see cref="OwnedRows"/> is derived from
    /// it rather than cached beside it, so the encoded and decoded forms cannot
    /// drift and the record's value equality stays correct (ADR-0017 D2).
    /// </remarks>
    public string Value { get; }

    /// <summary>
    /// Gets the rows that own this entity: a single element for
    /// <see cref="SceneOwnerKind.Row"/>, every row for
    /// <see cref="SceneOwnerKind.Rows"/>, and an empty list for
    /// <see cref="SceneOwnerKind.Chart"/>, which belongs to the workbook.
    /// </summary>
    /// <remarks>
    /// This is the single accessor renderers and ownership repair use, so they
    /// never re-derive "which rows own this" from the encoded value themselves
    /// (ADR-0017 D1). It is computed rather than cached because a cached list
    /// would join the record's synthesized equality by reference and two
    /// independently built equal owners would compare unequal.
    /// </remarks>
    public IReadOnlyList<GanttRowId> OwnedRows => Kind switch
    {
        SceneOwnerKind.Row => GanttRowId.TryParse(Value, out GanttRowId? single) && single is not null
            ? [single]
            : [],
        SceneOwnerKind.Rows => DecodeRows(Value),
        SceneOwnerKind.Chart => [],
        // Chart is the only remaining kind and has no rows. The default arm is
        // kept deliberately: it stops a kind added later from silently
        // inheriting Chart's empty-row answer.
        _ => [],
    };

    private static List<GanttRowId> DecodeRows(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return [];
        }

        var parts = value.Split(RowsSeparator, StringSplitOptions.None);
        List<GanttRowId> rows = new(parts.Length);
        foreach (var part in parts)
        {
            if (GanttRowId.TryParse(part, out GanttRowId? rowId) && rowId is not null)
            {
                rows.Add(rowId);
            }
        }

        return rows;
    }

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

    /// <summary>Creates a shared owner from the rows that deduplicated to one primitive.</summary>
    /// <param name="rowIds">The contributing row identifiers; sorted and de-duplicated here.</param>
    /// <returns>The shared scene owner.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rowIds"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when fewer than two distinct row IDs are supplied. A single
    /// contributing row is a <see cref="ForRow"/> owner; allowing both spellings
    /// would put the same ownership in the snapshot two ways (ADR-0017 D2).
    /// </exception>
    public static SceneOwnerId ForRows(IEnumerable<GanttRowId> rowIds)
    {
        ArgumentNullException.ThrowIfNull(rowIds);

        // Sorted Ordinal with no case-folding: the ordering must not depend on
        // the host's culture, or the same row set would hash differently per
        // machine. De-duplication is by identifier equality, not by text, so it
        // agrees with the sort.
        List<GanttRowId> ordered = [.. rowIds.OrderBy(rowId => rowId.Value, StringComparer.Ordinal)];
        List<GanttRowId> distinct = [.. ordered.Distinct()];

        ArgumentOutOfRangeException.ThrowIfLessThan(
            distinct.Count,
            2,
            distinct.Count == 0
                ? nameof(rowIds)
                : $"{nameof(rowIds)} (a single row is not shared; use ForRow)");

        return new(
            SceneOwnerKind.Rows,
            string.Join(RowsSeparator, distinct.Select(rowId => rowId.Value)));
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

        if (string.Equals(kind, "rows", StringComparison.Ordinal) && TryParseRows(value, out SceneOwnerId? shared))
        {
            owner = shared;
            return true;
        }

        return false;
    }

    private static bool TryParseRows(string? value, out SceneOwnerId? owner)
    {
        owner = null;
        if (string.IsNullOrEmpty(value))
        {
            return false;
        }

        var parts = value.Split(RowsSeparator, StringSplitOptions.None);
        if (parts.Length < 2)
        {
            return false;
        }

        // Strictly ascending also rejects a repeated ID, so the parsed set is
        // genuinely distinct without a second comparison pass. The rows are not
        // stored: Value is the whole state, and OwnedRows decodes on demand.
        List<GanttRowId> rows = new(parts.Length);
        foreach (var part in parts)
        {
            if (!GanttRowId.TryParse(part, out GanttRowId? rowId) || rowId is null)
            {
                return false;
            }

            if (rows.Count > 0 && string.CompareOrdinal(rows[^1].Value, rowId.Value) >= 0)
            {
                return false;
            }

            rows.Add(rowId);
        }

        owner = new(SceneOwnerKind.Rows, value);
        return true;
    }
}
