namespace GanttCreator.Core;

/// <summary>Why a column-visibility request was refused (R5.18h).</summary>
public enum ColumnVisibilityRefusal
{
    /// <summary>The stored list named a column the ribbon does not offer.</summary>
    UnknownColumn = 0,

    /// <summary>The stored list could not be read as a list of column names.</summary>
    MalformedList = 1,
}

/// <summary>The typed result of a column-visibility operation.</summary>
/// <param name="Hidden">The resulting hidden set, or <see langword="null"/> on refusal.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record ColumnVisibilityOutcome(
    IReadOnlySet<string>? Hidden,
    ColumnVisibilityRefusal? Refusal)
{
    /// <summary>Gets whether the operation produced a state.</summary>
    public bool Succeeded => Hidden is not null && Refusal is null;
}

/// <summary>
/// The set of table columns the user has hidden, and the rules for changing it (R5.18h).
/// </summary>
/// <remarks>
/// <para>
/// <b>Six columns, and only these.</b> The dropdown offers the six the user reads and
/// edits. The twelve engine columns are hidden by the schema's own classification
/// (<see cref="GanttColumnAccess.EngineHidden"/>) and are not the user's to show: an
/// engine column made visible is a column whose cells the add-in owns, and the whole
/// point of ADR-0029 D8 is that those stay out of sight.
/// </para>
/// <para>
/// <b>Zero visible columns is a supported state (ADR-0065 D1).</b> The analyst may hide
/// every user-visible column and work with the chart alone; the table, its rows and its
/// cells stay on the worksheet and the add-in keeps reading and writing them. The toggle
/// no longer refuses the last authoring column, and <see cref="ToggleAll"/> moves the whole
/// offered set between visible and hidden.
/// </para>
/// <para>
/// <b>One strictly-parsed settings key.</b> The choice is stored as a comma-separated
/// list of schema display names in <see cref="SettingKey"/>, and a value that does not
/// parse is REFUSED rather than treated as "nothing hidden": a silent default would
/// un-hide a column the user hid and call it a fresh workbook. <see cref="Format"/> is
/// the only writer, and it emits the names in schema order so the stored text is
/// deterministic and a diff of two workbooks shows a real change rather than an ordering
/// artefact.
/// </para>
/// </remarks>
public static class GanttColumnVisibility
{
    /// <summary>The settings key that stores the hidden column names (schema version 24).</summary>
    public const string SettingKey = "HiddenColumns";

    /// <summary>
    /// The columns the dropdown offers, in schema order.
    /// </summary>
    /// <remarks>
    /// Declared once, here, and asserted against the schema by a test: a name that is not
    /// a visible column would offer the user a toggle that hides nothing they can see.
    /// </remarks>
    public static IReadOnlyList<string> ToggleableColumnNames { get; } =
        ["Type", "Activity ID", "Description", "Start", "Finish", "Duration"];

    /// <summary>
    /// Returns whether the schema, rather than the user, owns a column's hidden flag.
    /// </summary>
    /// <param name="column">The schema column, or null.</param>
    /// <returns><see langword="true"/> for an engine-hidden column.</returns>
    /// <remarks>
    /// <para>
    /// Only the engine columns' visibility is the add-in's: they are hidden because the
    /// add-in owns their cells, and a table with them on show invites a hand edit that the
    /// next Refresh overwrites. Every user-visible column's hidden flag is the dropdown's
    /// choice, recorded in <see cref="SettingKey"/>.
    /// </para>
    /// <para>
    /// <b>This exists so a writer cannot undo the user's choice by accident.</b>
    /// `ExcelColumnPresentationRestorer` and `ExcelWorkbookRecovery` both force a column's
    /// visibility from <see cref="GanttTableSchema.Default"/> on every Refresh and Repair;
    /// written unconditionally, that silently un-hides a column the user hid, once per
    /// pass. Both ask this question first instead, and the answer is asserted here rather
    /// than restated at each call site.
    /// </para>
    /// </remarks>
    public static bool SchemaOwnsVisibility(GanttTableColumn? column) =>
        column is not null && column.IsHidden;

    /// <summary>
    /// The stable ribbon row id for a column: its display name with spaces removed.
    /// </summary>
    /// <param name="columnName">The exact schema display name.</param>
    /// <returns>The id suffix to append to the control-id prefix.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="columnName"/> is null.</exception>
    /// <remarks>
    /// <para>
    /// <b>RibbonX ids cannot contain a space.</b> The <c>id</c> attribute is typed
    /// <c>ST_UniqueID</c>, and the live host rejected <c>colActivity ID</c> on load:
    /// "A name contained an invalid character ... Error parsing 'colActivity ID' as ID
    /// datatype" (2026-11-17). The display name keeps its space because it is the header
    /// text; only the control id is slugged, so <c>Activity ID</c> becomes
    /// <c>ActivityID</c> and the column it controls is still the one the header names.
    /// </para>
    /// <para>
    /// Space removal is the whole rule rather than a per-column table: it is
    /// deterministic, it leaves every single-word column's id unchanged
    /// (<c>colStart</c> stays <c>colStart</c>), and a future display name with a space
    /// gets a legal id without an edit here. <see cref="TryResolveRibbonItemId"/> is the
    /// inverse lookup the ribbon callbacks use, so the mapping has one home.
    /// </para>
    /// </remarks>
    public static string RibbonItemId(string columnName)
    {
        ArgumentNullException.ThrowIfNull(columnName);
        return columnName.Replace(" ", string.Empty, StringComparison.Ordinal);
    }

    /// <summary>
    /// Resolves a ribbon row id back to the column's exact display name.
    /// </summary>
    /// <param name="itemId">The id suffix, without the control-id prefix.</param>
    /// <param name="columnName">The display name when the id resolves.</param>
    /// <returns><see langword="true"/> when the id names an offered column.</returns>
    public static bool TryResolveRibbonItemId(
        string? itemId,
        [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? columnName)
    {
        columnName = null;
        if (string.IsNullOrEmpty(itemId))
        {
            return false;
        }

        foreach (string name in ToggleableColumnNames)
        {
            if (string.Equals(RibbonItemId(name), itemId, StringComparison.Ordinal))
            {
                columnName = name;
                return true;
            }
        }

        return false;
    }

    /// <summary>Returns whether the dropdown offers this column.</summary>
    /// <param name="columnName">The exact schema display name, or null.</param>
    /// <returns><see langword="true"/> for a name in <see cref="ToggleableColumnNames"/>.</returns>
    public static bool IsToggleable(string? columnName)
    {
        if (string.IsNullOrEmpty(columnName))
        {
            return false;
        }

        foreach (string name in ToggleableColumnNames)
        {
            if (string.Equals(name, columnName, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Reads the stored hidden-column list.
    /// </summary>
    /// <param name="stored">The stored value, or null/blank for "nothing hidden".</param>
    /// <returns>The hidden set, or a refusal naming the first unreadable entry.</returns>
    /// <remarks>
    /// A blank value is the default and means nothing is hidden; anything that is not a
    /// comma-separated list of known column names is refused, because the alternative is
    /// silently showing a column the user hid.
    /// </remarks>
    public static ColumnVisibilityOutcome Parse(string? stored)
    {
        HashSet<string> hidden = new(StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(stored))
        {
            return new ColumnVisibilityOutcome(hidden, null);
        }

        string[] parts = stored.Split(',', StringSplitOptions.None);
        foreach (string part in parts)
        {
            string name = part.Trim();
            if (name.Length == 0)
            {
                // An empty entry is a malformed list, not "no columns": `Type,,Start`
                // came from something other than Format, and guessing its meaning is
                // how a hidden column reappears.
                return new ColumnVisibilityOutcome(null, ColumnVisibilityRefusal.MalformedList);
            }

            if (!IsToggleable(name))
            {
                return new ColumnVisibilityOutcome(null, ColumnVisibilityRefusal.UnknownColumn);
            }

            _ = hidden.Add(name);
        }

        return new ColumnVisibilityOutcome(hidden, null);
    }

    /// <summary>
    /// Writes the hidden set as the stored value.
    /// </summary>
    /// <param name="hidden">The hidden columns.</param>
    /// <returns>A comma-separated list in schema order, or an empty string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="hidden"/> is null.</exception>
    public static string Format(IReadOnlySet<string> hidden)
    {
        ArgumentNullException.ThrowIfNull(hidden);

        List<string> ordered = [];
        foreach (string name in ToggleableColumnNames)
        {
            if (hidden.Contains(name))
            {
                ordered.Add(name);
            }
        }

        return string.Join(',', ordered);
    }

    /// <summary>
    /// Flips one column between shown and hidden.
    /// </summary>
    /// <param name="hidden">The current hidden set.</param>
    /// <param name="columnName">The column to flip.</param>
    /// <returns>The new set, or a refusal.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="hidden"/> is null.</exception>
    public static ColumnVisibilityOutcome Toggle(IReadOnlySet<string> hidden, string? columnName)
    {
        ArgumentNullException.ThrowIfNull(hidden);

        if (!IsToggleable(columnName))
        {
            return new ColumnVisibilityOutcome(null, ColumnVisibilityRefusal.UnknownColumn);
        }

        HashSet<string> next = new(hidden, StringComparer.Ordinal);
        if (!next.Remove(columnName!))
        {
            // Remove reported the name absent, so this add cannot be a duplicate.
            _ = next.Add(columnName!);
        }

        return new ColumnVisibilityOutcome(next, null);
    }

    /// <summary>
    /// Flips the whole offered set: any hidden column means show all, none hidden means hide
    /// all (ADR-0065 D3).
    /// </summary>
    /// <param name="hidden">The current hidden set.</param>
    /// <returns>The new hidden set.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="hidden"/> is null.</exception>
    /// <remarks>
    /// The primary split-button action is this one rule, so a partial state — some columns
    /// hidden and some shown — shows all rather than completing the hide. That keeps the
    /// big button a forgiving reset and leaves precise hiding to the menu rows.
    /// </remarks>
    public static ColumnVisibilityOutcome ToggleAll(IReadOnlySet<string> hidden)
    {
        ArgumentNullException.ThrowIfNull(hidden);

        HashSet<string> next = hidden.Count == 0
            ? new HashSet<string>(ToggleableColumnNames, StringComparer.Ordinal)
            : new HashSet<string>(StringComparer.Ordinal);
        return new ColumnVisibilityOutcome(next, null);
    }

    /// <summary>
    /// Returns the schema's visible columns that the hidden set does not name, in schema order.
    /// </summary>
    /// <param name="hidden">The user's hidden set.</param>
    /// <returns>The columns the live panel measures and the export reproduces.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="hidden"/> is null.</exception>
    /// <remarks>
    /// <para>
    /// <b>One authority for "what is visible" (ADR-0065 D6).</b> The live measurement reads
    /// this list, and the future export composition must read it too, so a hidden column
    /// cannot be measured by one half of the product and exported by the other.
    /// </para>
    /// <para>
    /// An empty result is the chart-only state, not a refusal. The engine columns are
    /// excluded by the schema's own classification rather than by a second list here.
    /// </para>
    /// </remarks>
    public static IReadOnlyList<string> VisibleColumns(IReadOnlySet<string> hidden)
    {
        ArgumentNullException.ThrowIfNull(hidden);

        return
        [
            .. GanttTableSchema.Default.Columns
                .Where(column => !column.IsHidden && !hidden.Contains(column.Name))
                .Select(column => column.Name),
        ];
    }
}
