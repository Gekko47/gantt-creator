namespace GanttCreator.Core;

/// <summary>
/// The single authority for "is this the Gantt data table?" in the product.
/// </summary>
/// <remarks>
/// <para>
/// Every Office adapter that walks a workbook looking for the Gantt table
/// compares a table name against <see cref="GanttTableSchema.TableName"/>. That
/// comparison was written out inline in twenty-odd adapters, and the copies had
/// drifted: most were case-insensitive, four were ordinal. A single predicate
/// makes the rule reviewable in one place.
/// </para>
/// <para>
/// <b>Two strengths, deliberately.</b> Excel preserves whatever case a table was
/// created with, so the lookup rule documented by
/// <c>ExcelRowHeightNormaliser</c> and <c>ExcelOutlineGroupWriter</c> is
/// case-insensitive, and that is <see cref="IsDataTableName(string?)"/>.
/// The protection-arming and shape-ownership paths
/// (<c>ExcelWorksheetProtectionPort</c>, <c>ExcelShapeWriter</c>,
/// <c>ExcelPanelGridMeasurement</c>, <c>ExcelColumnPresentationRestorer</c>) use
/// the ordinal rule today; <see cref="IsDataTableNameExact(string?)"/> preserves
/// that behaviour exactly so this change is a refactor, not a silent weakening
/// or tightening of a safety path. Unifying the two strengths is a product
/// decision, not a refactor: it decides whether a case-renamed table is treated
/// as ours (chart rendered onto it, protection armed on it) or as foreign. Until
/// that decision is recorded, both strengths stay named here rather than being
/// re-inlined at the call sites.
/// </para>
/// <para>
/// This type is the seam the single-chart locator (R5.17e) grows from: a future
/// per-chart contract supplies the name this predicate compares against.
/// </para>
/// </remarks>
public static class GanttChartContract
{
    /// <summary>
    /// Whether <paramref name="tableName"/> names the Gantt data table, using the
    /// case-insensitive rule Excel itself applies to table names.
    /// </summary>
    /// <param name="tableName">The table name read from a workbook, or null.</param>
    /// <returns>
    /// <see langword="true"/> when the name matches
    /// <see cref="GanttTableSchema.TableName"/> ignoring case; otherwise
    /// <see langword="false"/>. A null name is never the Gantt data table.
    /// </returns>
    public static bool IsDataTableName(string? tableName) =>
        string.Equals(tableName, GanttTableSchema.TableName, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Whether <paramref name="tableName"/> names the Gantt data table under the
    /// ordinal rule used by the protection and shape-ownership paths.
    /// </summary>
    /// <param name="tableName">The table name read from a workbook, or null.</param>
    /// <returns>
    /// <see langword="true"/> when the name matches
    /// <see cref="GanttTableSchema.TableName"/> exactly, including case;
    /// otherwise <see langword="false"/>.
    /// </returns>
    /// <remarks>
    /// Kept as its own member so the existing ordinal sites keep their exact
    /// behaviour. See the type remarks for why the divergence is named rather
    /// than erased by this change.
    /// </remarks>
    public static bool IsDataTableNameExact(string? tableName) =>
        string.Equals(tableName, GanttTableSchema.TableName, StringComparison.Ordinal);
}
