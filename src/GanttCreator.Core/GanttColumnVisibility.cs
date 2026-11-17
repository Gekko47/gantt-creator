namespace GanttCreator.Core;

/// <summary>Why a column-visibility request was refused (R5.18h).</summary>
public enum ColumnVisibilityRefusal
{
    /// <summary>The stored list named a column the ribbon does not offer.</summary>
    UnknownColumn = 0,

    /// <summary>Showing or hiding would leave no authoring column visible.</summary>
    LastVisibleColumn = 1,

    /// <summary>The stored list could not be read as a list of column names.</summary>
    MalformedList = 2,
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
/// <b>The last visible authoring column is refused.</b> A table with no authoring column
/// visible cannot be authored by hand at all — the user could not type the next activity
/// without going to the ribbon first — so the toggle that would produce it is refused
/// with its own reason rather than allowed and left to be discovered.
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
            if (!AnyAuthoringColumnVisible(next))
            {
                return new ColumnVisibilityOutcome(null, ColumnVisibilityRefusal.LastVisibleColumn);
            }
        }

        return new ColumnVisibilityOutcome(next, null);
    }

    /// <summary>
    /// Returns whether at least one column the schema classifies as authoring is visible.
    /// </summary>
    /// <param name="hidden">The candidate hidden set.</param>
    /// <returns><see langword="true"/> when the table can still be authored by hand.</returns>
    /// <remarks>
    /// The authoring set comes from the schema's own classification rather than from a
    /// second list here, so a future column classified as authoring is covered without
    /// anyone remembering to add it.
    /// </remarks>
    internal static bool AnyAuthoringColumnVisible(IReadOnlySet<string> hidden)
    {
        foreach (GanttTableColumn column in GanttTableSchema.Default.Columns)
        {
            if (column.Access == GanttColumnAccess.Authoring && !hidden.Contains(column.Name))
            {
                return true;
            }
        }

        return false;
    }
}
