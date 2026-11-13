namespace GanttCreator.Core;

/// <summary>How a visible table cell's text is aligned horizontally.</summary>
/// <remarks>
/// A closed set of the two alignments the contract actually uses. Word wrap, rotation, and
/// indent are not exposed: none of them is part of the data panel's presentation, and an enum
/// that could express them would invite a caller to set one.
/// </remarks>
public enum GanttCellAlignment
{
    /// <summary>Text starts at the cell's left edge. Authoring text reads this way.</summary>
    Left = 0,

    /// <summary>Text is centred in the cell. Numeric and date values read this way.</summary>
    Centre = 1,
}

/// <summary>
/// The chart theme's typography and alignment, as the visible data table consumes them.
/// </summary>
/// <remarks>
/// <para>
/// The table is presented from the same code-owned tokens the chart and its export use, so the
/// live panel and the exported panel cannot drift apart: the font family and body size come from
/// <see cref="GanttCatalogues.Typography"/>, which is the single authority the entity guide's
/// typography table is materialised from.
/// </para>
/// <para>
/// <b>Vertical alignment is not a setting.</b> Every visible cell is middle-aligned, because a
/// table row is one worksheet row and a top-aligned cell would sit against the row above it as
/// soon as the row grew. It is therefore a constant rather than a member: there is nothing to
/// choose.
/// </para>
/// <para>
/// <b>Horizontal alignment follows the value, not the column number.</b> Dates and durations are
/// short, comparable values and read centred; free text reads from the left. The set is derived
/// from the schema's own date-column list plus the one computed value, so renaming or moving a
/// column in the schema moves its alignment with it.
/// </para>
/// </remarks>
public static class GanttTableAppearance
{
    /// <summary>
    /// The columns whose values are centred because they are short comparable values rather
    /// than prose.
    /// </summary>
    public static IReadOnlyList<string> CentredColumnNames { get; } = ["Start", "Finish", "Duration"];

    /// <summary>The theme's font family for table text.</summary>
    public static string FontFamily => GanttCatalogues.TypographyDefaultText("FontFamily");

    /// <summary>The theme's body text size, in points, for table text.</summary>
    public static double BodyFontSizePt => GanttCatalogues.TypographyDefault("BodyFontSizePt");

    /// <summary>
    /// Returns the horizontal alignment for a column.
    /// </summary>
    /// <param name="columnName">The exact header text.</param>
    /// <returns>
    /// <see cref="GanttCellAlignment.Centre"/> for a listed value column, otherwise
    /// <see cref="GanttCellAlignment.Left"/> — including for an absent or unknown name, where
    /// left-aligned prose is the safe reading rather than a centred guess.
    /// </returns>
    /// <remarks>
    /// The comparison is ordinal, like every other header lookup in this schema.
    /// </remarks>
    public static GanttCellAlignment AlignmentFor(string? columnName)
    {
        if (!string.IsNullOrEmpty(columnName))
        {
            foreach (string name in CentredColumnNames)
            {
                if (string.Equals(name, columnName, StringComparison.Ordinal))
                {
                    return GanttCellAlignment.Centre;
                }
            }
        }

        return GanttCellAlignment.Left;
    }
}
