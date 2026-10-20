namespace GanttCreator.Core;

/// <summary>
/// The family a built-in style belongs to, used to group the Entity styles list
/// so every style appears under exactly one heading.
/// </summary>
/// <remarks>
/// The family is a <b>presentation grouping</b> of the code-owned presets,
/// transcribed from the entity guide's own per-entity sections. It is not a
/// workbook column: it never reaches <c>tblGanttStyles</c> or
/// <c>tblGanttTypes</c>, so it carries no schema version and no catalogue
/// hash.
/// </remarks>
public enum GanttEntityFamily
{
    /// <summary>Activity spans, including the named custom activity.</summary>
    Activities = 0,

    /// <summary>Milestone diamonds, including the named custom milestone.</summary>
    Milestones = 1,

    /// <summary>Procurement spans, which carry the diagonal hatch.</summary>
    Procurement = 2,

    /// <summary>Chart elements: the vertical line, the section header and the spacer.</summary>
    ChartElements = 3,
}

/// <summary>
/// The single authority for the family a built-in style key belongs to.
/// </summary>
/// <remarks>
/// <para>
/// The map is deliberately code-owned and exhaustive over
/// <see cref="GanttCatalogues.StylePresets"/>: a style with no entry here is a
/// defect, not a style to file under a fallback heading. The previous
/// arrangement — a presentation switch in the dialog with a catch-all bucket —
/// produced two duplicated headings and filed the two custom types away from
/// the families they belong to, because a style's family is a property of the
/// catalogue rather than of the order the dialog happens to render it in.
/// </para>
/// <para>
/// A missing key throws rather than defaulting: a silent fallback is a second
/// authority that disagrees with this one at some margin.
/// </para>
/// </remarks>
public static class GanttStyleFamilies
{
    private static readonly Dictionary<string, GanttEntityFamily> _map =
        new(StringComparer.Ordinal)
        {
            // Activities (entity guide sections 12 to 17 and 19).
            ["AsBuiltActivity"] = GanttEntityFamily.Activities,
            ["AsPlannedActivity"] = GanttEntityFamily.Activities,
            ["BaselineActivity"] = GanttEntityFamily.Activities,
            ["CriticalInterval"] = GanttEntityFamily.Activities,
            ["DelayEvent"] = GanttEntityFamily.Activities,
            ["CustomActivity"] = GanttEntityFamily.Activities,

            // Milestones (entity guide sections 20 and 21).
            ["AsBuiltMilestone"] = GanttEntityFamily.Milestones,
            ["AsPlannedMilestone"] = GanttEntityFamily.Milestones,
            ["BaselineMilestone"] = GanttEntityFamily.Milestones,
            ["CriticalMilestone"] = GanttEntityFamily.Milestones,
            ["CustomMilestone"] = GanttEntityFamily.Milestones,

            // Procurement (entity guide section 18).
            ["AsBuiltProcurement"] = GanttEntityFamily.Procurement,
            ["AsPlannedProcurement"] = GanttEntityFamily.Procurement,
            ["BaselineProcurement"] = GanttEntityFamily.Procurement,

            // Chart elements (entity guide sections 10, 11 and 24).
            ["DefaultDelineator"] = GanttEntityFamily.ChartElements,
            ["Splitter"] = GanttEntityFamily.ChartElements,
            ["Spacer"] = GanttEntityFamily.ChartElements,
        };

    /// <summary>
    /// Gets the families in the order the Entity styles list presents them.
    /// </summary>
    public static IReadOnlyList<GanttEntityFamily> OrderedFamilies { get; } =
    [
        GanttEntityFamily.Activities,
        GanttEntityFamily.Milestones,
        GanttEntityFamily.Procurement,
        GanttEntityFamily.ChartElements,
    ];

    /// <summary>
    /// Returns the family a built-in style key belongs to.
    /// </summary>
    /// <param name="styleKey">The built-in style key.</param>
    /// <returns>The style's family.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when the key is not a built-in preset, or names a preset the map
    /// does not cover.
    /// </exception>
    public static GanttEntityFamily For(string styleKey)
    {
        ArgumentNullException.ThrowIfNull(styleKey);

        return _map.TryGetValue(styleKey, out GanttEntityFamily family)
            ? family
            : throw new ArgumentOutOfRangeException(
                nameof(styleKey),
                styleKey,
                "The style key names no built-in preset family.");
    }

    /// <summary>
    /// Returns the human-readable heading for a family.
    /// </summary>
    /// <param name="family">The family.</param>
    /// <returns>The heading text.</returns>
    public static string DisplayNameFor(GanttEntityFamily family) => family switch
    {
        GanttEntityFamily.Activities => "Activities",
        GanttEntityFamily.Milestones => "Milestones",
        GanttEntityFamily.Procurement => "Procurement",
        GanttEntityFamily.ChartElements => "Chart elements",
        _ => throw new ArgumentOutOfRangeException(nameof(family), family, "The family is not defined."),
    };
}

/// <summary>
/// Which built-in styles' stroke dash the user may change.
/// </summary>
/// <remarks>
/// <para>
/// The delineator is the one built-in entity drawn as a <i>line</i> rather
/// than a filled shape, so its line type is a property the user is expected
/// to choose. Every other built-in style's dash is code-owned: the reader
/// exact-matches that column against the preset, so an edited value could
/// not be kept, and offering a control for it would promise a change the
/// workbook would silently discard.
/// </para>
/// <para>
/// This is the single authority for that rule. The reader, the writer, and
/// the request validator all consult it, so the three cannot disagree about
/// which dashes are editable.
/// </para>
/// </remarks>
public static class GanttStyleDashPolicy
{
    /// <summary>
    /// Returns whether the stroke dash of a built-in style is user-editable.
    /// </summary>
    /// <param name="styleKey">The built-in style key.</param>
    /// <returns>True when the dash may be edited.</returns>
    public static bool IsUserEditable(string styleKey)
    {
        ArgumentNullException.ThrowIfNull(styleKey);
        return string.Equals(styleKey, "DefaultDelineator", StringComparison.Ordinal);
    }
}
