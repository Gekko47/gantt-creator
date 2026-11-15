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

    /// <summary>The built-in Excel table style the add-in assigns to the data table.</summary>
    /// <remarks>
    /// <para>
    /// <b>Why a style as well as direct formatting (owner ruling 2026-11-14).</b> Hiding and
    /// re-showing a table column makes Excel re-apply the table's own style; the add-in's
    /// <em>direct</em> cell formatting — header fill and type, border colour and the
    /// outside weight — can be dropped by that re-application, and a Refresh deliberately
    /// never rewrites cell formatting. Assigning a neutral built-in style during Initialise
    /// and Repair means the host's re-application produces the same thin grey lines every
    /// time, with the theme's colour and weight written on top by the adapters.
    /// </para>
    /// <para>
    /// "Table Style Light 1" is the neutral one: no banding, no fill, thin grey grid lines —
    /// the closest built-in to the presentation this guide describes.
    /// </para>
    /// </remarks>
    public static string TableStyleName => "TableStyleLight1";

    /// <summary>The theme's font family for table text.</summary>
    public static string FontFamily => GanttCatalogues.TypographyDefaultText("FontFamily");

    /// <summary>The theme's body text size, in points, for table text.</summary>
    public static double BodyFontSizePt => GanttCatalogues.TypographyDefault("BodyFontSizePt");

    /// <summary>The header band fill: the chart's period-band fill.</summary>
    /// <remarks>
    /// <para>
    /// <b>The header row is the period band's row, so it takes the period band's fill</b>
    /// (ADR-0030 D5, owner ruling 2026-11-14). It previously took the <c>HeaderFill</c>
    /// colour token, which is white, while the band beside it rendered the neutral
    /// <c>AlternateBandFill</c> — one heading in two fills, reported from the live host.
    /// The role is read through <see cref="GanttChartChrome"/> rather than restating the
    /// token, so the band, its setting key, and this fill cannot disagree.
    /// </para>
    /// <para>
    /// The fill is the chart-chrome <em>default</em>. A workbook whose user changed
    /// <c>PeriodHeaderColour</c> is not reflected here, exactly as the header's type is not
    /// reflected from the chart chrome: cell presentation is written by Initialise and
    /// Repair from the code-owned theme, and a Refresh never rewrites it.
    /// </para>
    /// </remarks>
    public static ColourHex HeaderFill =>
        ColourHex.Parse(GanttChartChrome.DefaultHexFor(GanttChartChromeRole.PeriodHeader));

    /// <summary>
    /// The colour the visible header band's text takes: the theme's body text colour.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>A fill without its text colour is an invisible header.</b> Live evidence
    /// 2026-11-14: the initialiser wrote <see cref="HeaderFill"/> (white) over the header
    /// row but never the font colour, so the table style's own header font — white on
    /// <c>TableStyleMedium2</c>'s dark band — stayed white. The header cells held every
    /// schema name and the user could see none of them.
    /// </para>
    /// <para>
    /// The body text colour, not the export panel's header text colour: the panel's header
    /// is white <em>on</em> a dark band, whereas the live band is filled with the period
    /// band's neutral <see cref="HeaderFill"/> — a light fill, which needs dark text. The
    /// two must be read as a pair, and this is the pair the theme actually supports.
    /// </para>
    /// </remarks>
    public static ColourHex HeaderTextColour => Colour("DefaultText");

    /// <summary>The theme's header text size, in points.</summary>
    /// <remarks>
    /// <b>9pt, because that is the type the chart's period band draws.</b> The owner's
    /// ruling (2026-11-14) is that the live header row and the period band above the plot
    /// are one heading, so the table reads as the band's own column labels rather than as
    /// a second, smaller row. The size comes from the catalogue token the entity guide
    /// already names, so the guide and the host cannot disagree.
    /// </remarks>
    public static double HeaderFontSizePt => GanttCatalogues.TypographyDefault("HeaderFontSizePt");

    /// <summary>Whether the theme's header text is bold; it is regular.</summary>
    /// <remarks>
    /// The period band is regular, so the header is too. This is written explicitly rather
    /// than left alone: Excel's table style bolds the header row, and a value the adapter
    /// never writes is a value Excel keeps.
    /// </remarks>
    public static bool HeaderIsBold => GanttCatalogues.TypographyDefaultIsBold("HeaderFontSizePt");

    /// <summary>
    /// The width, in points, the table's outer border shares with the chart's frame.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The frame's closing boundary is the structural <c>MajorBoundaryPt</c> token, so the
    /// table's own outside edge reads as the frame's edge where the two meet (owner ruling
    /// 2026-11-14). Reading the token rather than restating <c>1</c> keeps the two
    /// boundaries one decision.
    /// </para>
    /// <para>
    /// Only the frame token's default is mirrored on a cell border today: Excel's
    /// <c>Border.Weight</c> is a closed set of weights, not an arbitrary point width, so a
    /// user-raised frame width (R5.6d) cannot be reproduced exactly on a cell. The adapter
    /// maps 1pt to Excel's <c>xlThin</c>, which is the 1pt weight.
    /// </para>
    /// </remarks>
    public static double OuterBorderWeightPt => GanttCatalogues.MetricDefault("MajorBoundaryPt");

    /// <summary>
    /// The theme colour a live cell border takes.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The owner's rulings take the <em>colour</em> of every cell border from the chart theme
    /// and the <em>weight</em> of every border the add-in draws from one number (2026-11-13,
    /// 2026-11-14 and ADR-0061 D5). The token is the minor grid stroke, the light boundary the
    /// chart already draws between time columns, so a table border and a grid line are one
    /// boundary colour rather than two near-misses.
    /// </para>
    /// <para>
    /// <b>The interior grid is drawn, not inherited (owner report 2026-11-15).</b> It used to be
    /// left to the table style on the theory that the style owned the cell grid. That theory
    /// cost a live defect: a table-style border is one device pixel painted through a style
    /// consult, and the host drops individual rules on rows whose height or position rounds to a
    /// sub-pixel, reported as "some rows are missing the vertical line between Start and
    /// Finish". A live measurement cleared this code base - the cells held a continuous rule in
    /// both the direct and the displayed border APIs, before and after <c>CalculateFull</c> - so
    /// the fix is to draw the grid the add-in presents: Initialise and Repair write the two
    /// interior indices as a continuous thin rule in this colour, and a Refresh still writes no
    /// cell formatting.
    /// </para>
    /// </remarks>
    public static ColourHex BorderColour => Colour("MinorGridStroke");

    /// <summary>Resolves a colour token by exact name.</summary>
    /// <param name="tokenName">The token name from the colour catalogue.</param>
    /// <returns>The parsed colour.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown when the catalogue has no token with that name — a missing token is a broken
    /// catalogue, not a colour to guess.
    /// </exception>
    private static ColourHex Colour(string tokenName)
    {
        foreach (GanttColourToken token in GanttCatalogues.Colours)
        {
            if (string.Equals(token.Name, tokenName, StringComparison.Ordinal))
            {
                return ColourHex.Parse(token.HexValue);
            }
        }

        throw new InvalidOperationException(
            $"The '{tokenName}' colour token is missing from the catalogue.");
    }

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
