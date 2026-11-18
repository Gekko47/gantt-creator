namespace GanttCreator.Core.Scene;

/// <summary>The reason an export-bounds resolution was refused.</summary>
public enum ExportBoundsRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The selection value is not one of the two defined export modes.</summary>
    InvalidSelection = 1,

    /// <summary>
    /// A table-plus-chart composition supplied no resolved panel styles, so the table
    /// could not be composed rather than silently degrading to the chart.
    /// </summary>
    PanelThemeMissing = 2,

    /// <summary>
    /// A table-plus-chart composition has no visible worksheet row to export, and an
    /// empty data table is not a valid export (entity guide section 1).
    /// </summary>
    EmptyVisibleTable = 3,

    /// <summary>
    /// The caller-supplied plot bounds have no area, so no export rectangle could be
    /// derived. Raised before any composition runs.
    /// </summary>
    InvalidBounds = 4,

    /// <summary>
    /// The scene builder refused the export composition. The scene builder's own typed
    /// reason is carried on the resolution's <see cref="ExportBoundsOutcome"/>.
    /// </summary>
    SceneBuildRefused = 5,
}

/// <summary>The actionable message the export contract attaches to a refusal.</summary>
/// <remarks>
/// Only the empty-table refusal carries one, because it is the one refusal the analyst
/// can act on: it names both ways forward. Every other refusal is either a construction
/// error or is already classified by the scene builder, so flattening those into
/// sentences here would be a second copy of a reason Core does not own.
/// </remarks>
public static class ExportBoundsRefusalMessages
{
    /// <summary>The empty-visible-table message for a table-plus-chart export.</summary>
    public const string EmptyVisibleTable =
        "A table-plus-chart copy needs at least one visible row. Make a row visible, or copy the chart only.";
}
