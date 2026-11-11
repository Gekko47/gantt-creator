namespace GanttCreator.Core;

/// <summary>
/// The result of comparing a live header row against the code-owned schema.
/// </summary>
public enum GanttTableHeaderAlignmentOutcome
{
    /// <summary>Every header cell already carries the schema name at its position.</summary>
    Aligned = 0,

    /// <summary>At least one header cell is blank and every non-blank cell matches its position.</summary>
    Writable = 1,

    /// <summary>A non-blank header cell differs from the schema name at its position.</summary>
    Misaligned = 2,
}

/// <summary>
/// A pure plan for restoring the visible Gantt table's header row.
/// </summary>
/// <param name="Outcome">Whether the row is aligned, can be completed, or is ambiguous.</param>
/// <param name="CellsToWrite">The zero-based schema positions whose live cell is blank.</param>
/// <param name="MismatchedColumn">The schema column name at the first mismatching position, when <see cref="Outcome"/> is <see cref="GanttTableHeaderAlignmentOutcome.Misaligned"/>.</param>
public sealed record GanttTableHeaderAlignmentPlan(
    GanttTableHeaderAlignmentOutcome Outcome,
    IReadOnlyList<int> CellsToWrite,
    string? MismatchedColumn = null)
{
    /// <summary>Gets whether the plan can be applied without guessing.</summary>
    public bool Succeeded => Outcome != GanttTableHeaderAlignmentOutcome.Misaligned;
}

/// <summary>
/// Compares the live header row with <see cref="GanttTableSchema.Default"/>.
/// </summary>
/// <remarks>
/// <para>
/// The rule is deliberately positional. The reader already maps by name when the
/// table is healthy, so a reordered table never reaches a header repair. A
/// name-at-position check is the smallest proof that filling a blank cell cannot
/// mislabel data. A renamed header, a deleted column, or a table shifted off the
/// contract origin fails the check and is refused.
/// </para>
/// <para>
/// Pure and Office-free so the repair boundary is testable without a host.
/// </para>
/// </remarks>
public static class GanttTableHeaderAlignment
{
    /// <summary>
    /// Plans the header-row repair.
    /// </summary>
    /// <param name="liveHeaderCells">The live header values in schema order; a missing value is treated as blank.</param>
    /// <returns>The plan.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="liveHeaderCells"/> is null.</exception>
    public static GanttTableHeaderAlignmentPlan Plan(IReadOnlyList<string?> liveHeaderCells)
    {
        ArgumentNullException.ThrowIfNull(liveHeaderCells);

        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        List<int> write = [];

        for (int index = 0; index < schema.Count; index++)
        {
            string? live = index < liveHeaderCells.Count ? liveHeaderCells[index] : null;
            if (string.IsNullOrWhiteSpace(live))
            {
                write.Add(index);
                continue;
            }

            if (!string.Equals(live.Trim(), schema[index].Name, StringComparison.Ordinal))
            {
                return new GanttTableHeaderAlignmentPlan(
                    GanttTableHeaderAlignmentOutcome.Misaligned,
                    [],
                    schema[index].Name);
            }
        }

        return write.Count == 0
            ? new GanttTableHeaderAlignmentPlan(GanttTableHeaderAlignmentOutcome.Aligned, [])
            : new GanttTableHeaderAlignmentPlan(GanttTableHeaderAlignmentOutcome.Writable, write);
    }
}

/// <summary>How a damaged table column is restored.</summary>
public enum GanttTableColumnRepairKind
{
    /// <summary>Rename an existing column in place.</summary>
    Rename = 0,

    /// <summary>Insert a new blank column at a schema position.</summary>
    Insert = 1,
}

/// <summary>One column change in a table-structure repair.</summary>
/// <param name="Kind">Rename or insert.</param>
/// <param name="SchemaIndex">The zero-based schema position.</param>
/// <param name="Name">The schema header name to write.</param>
public sealed record GanttTableColumnRepair(
    GanttTableColumnRepairKind Kind,
    int SchemaIndex,
    string Name);

/// <summary>Why a table-structure repair was refused.</summary>
public enum GanttTableStructureRefusal
{
    /// <summary>No refusal.</summary>
    None = 0,

    /// <summary>The live table is readable, so no structure repair is needed.</summary>
    Inconsistent = 1,

    /// <summary>A known column is not at its schema position and a deletion cannot be distinguished from a rename.</summary>
    ColumnOrderShifted = 2,

    /// <summary>Unknown and missing names cannot be paired safely.</summary>
    HeaderAmbiguous = 3,

    /// <summary>Extra columns make the fixed plot-anchor width unsafe.</summary>
    ExtraColumnsPresent = 4,
}

/// <summary>
/// A pure plan for restoring a partially damaged <c>tblGanttData</c>.
/// </summary>
/// <param name="Repairs">The ordered column changes to apply.</param>
/// <param name="Refusal">The typed refusal, or <see cref="GanttTableStructureRefusal.None"/>.</param>
public sealed record GanttTableStructureRepairPlan(
    IReadOnlyList<GanttTableColumnRepair> Repairs,
    GanttTableStructureRefusal Refusal)
{
    /// <summary>Gets whether the plan can be applied.</summary>
    public bool Succeeded => Refusal == GanttTableStructureRefusal.None;
}

/// <summary>
/// Plans the safe repair of a contract table whose required headers no longer resolve.
/// </summary>
public static class GanttTableStructureRepair
{
    /// <summary>
    /// Plans the repair from the live column names in table order.
    /// </summary>
    /// <param name="liveColumnNames">The live table column names in table order.</param>
    /// <returns>The plan.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="liveColumnNames"/> is null.</exception>
    public static GanttTableStructureRepairPlan Plan(IReadOnlyList<string> liveColumnNames)
    {
        ArgumentNullException.ThrowIfNull(liveColumnNames);

        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        if (liveColumnNames.Count > schema.Count)
        {
            return Refused(GanttTableStructureRefusal.ExtraColumnsPresent);
        }

        Dictionary<string, int> schemaIndexByName = new(StringComparer.Ordinal);
        for (int index = 0; index < schema.Count; index++)
        {
            schemaIndexByName[schema[index].Name] = index;
        }

        HashSet<string> present = new(StringComparer.Ordinal);
        List<int> unknownPositions = [];
        for (int index = 0; index < liveColumnNames.Count; index++)
        {
            string name = liveColumnNames[index]?.Trim() ?? string.Empty;
            if (schemaIndexByName.ContainsKey(name))
            {
                _ = present.Add(name);
            }
            else if (index < schema.Count)
            {
                unknownPositions.Add(index);
            }
        }

        List<int> missing = [];
        for (int index = 0; index < schema.Count; index++)
        {
            if (schema[index].IsRequired && !present.Contains(schema[index].Name))
            {
                missing.Add(index);
            }
        }

        return missing.Count == 0 && unknownPositions.Count == 0
            ? Refused(GanttTableStructureRefusal.Inconsistent)
            : unknownPositions.Count is > 0
                ? PlanRenames(liveColumnNames, schema, schemaIndexByName, missing)
                : PlanInserts(liveColumnNames, schema, schemaIndexByName, missing);
    }

    private static GanttTableStructureRepairPlan PlanRenames(
        IReadOnlyList<string> liveColumnNames,
        IReadOnlyList<GanttTableColumn> schema,
        Dictionary<string, int> schemaIndexByName,
        List<int> missing)
    {
        // A rename is safe only when every known column still sits at its schema
        // position and each unknown position is exactly a missing schema name.
        Dictionary<int, string> unknownAtPosition = [];
        for (int index = 0; index < liveColumnNames.Count && index < schema.Count; index++)
        {
            string name = liveColumnNames[index]?.Trim() ?? string.Empty;
            if (schemaIndexByName.ContainsKey(name))
            {
                if (!string.Equals(name, schema[index].Name, StringComparison.Ordinal))
                {
                    return Refused(GanttTableStructureRefusal.ColumnOrderShifted);
                }
            }
            else
            {
                unknownAtPosition[index] = name;
            }
        }

        HashSet<int> missingSet = [.. missing];
        foreach (int position in unknownAtPosition.Keys)
        {
            if (!missingSet.Contains(position))
            {
                return Refused(GanttTableStructureRefusal.HeaderAmbiguous);
            }
        }

        if (missing.Count != unknownAtPosition.Count)
        {
            // A missing column with no unknown in its place is a deletion, which
            // cannot be paired with a rename in the same pass.
            return Refused(GanttTableStructureRefusal.HeaderAmbiguous);
        }

        List<GanttTableColumnRepair> repairs = [];
        foreach (int position in unknownAtPosition.Keys.OrderBy(value => value))
        {
            repairs.Add(
                new GanttTableColumnRepair(
                    GanttTableColumnRepairKind.Rename,
                    position,
                    schema[position].Name));
        }

        return new GanttTableStructureRepairPlan(repairs, GanttTableStructureRefusal.None);
    }

    private static GanttTableStructureRepairPlan PlanInserts(
        IReadOnlyList<string> liveColumnNames,
        IReadOnlyList<GanttTableColumn> schema,
        Dictionary<string, int> schemaIndexByName,
        List<int> missing)
    {
        // The known names must be an in-order subsequence of the schema. Any
        // reorder would make the insert position ambiguous.
        int schemaCursor = 0;
        foreach (string raw in liveColumnNames)
        {
            string name = raw?.Trim() ?? string.Empty;
            if (!schemaIndexByName.TryGetValue(name, out int schemaIndex))
            {
                return Refused(GanttTableStructureRefusal.HeaderAmbiguous);
            }

            while (schemaCursor < schemaIndex)
            {
                if (!missing.Contains(schemaCursor))
                {
                    return Refused(GanttTableStructureRefusal.ColumnOrderShifted);
                }

                schemaCursor++;
            }

            if (schemaCursor != schemaIndex)
            {
                return Refused(GanttTableStructureRefusal.ColumnOrderShifted);
            }

            schemaCursor++;
        }

        List<GanttTableColumnRepair> repairs = [];
        foreach (int position in missing.OrderBy(value => value))
        {
            repairs.Add(
                new GanttTableColumnRepair(
                    GanttTableColumnRepairKind.Insert,
                    position,
                    schema[position].Name));
        }

        return new GanttTableStructureRepairPlan(repairs, GanttTableStructureRefusal.None);
    }

    private static GanttTableStructureRepairPlan Refused(GanttTableStructureRefusal refusal) =>
        new([], refusal);
}
