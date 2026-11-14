namespace GanttCreator.Core.Tests;

/// <summary>
/// R5.18d: the visible table is presented from the same code-owned typography and alignment
/// authorities the chart and its export use, so the live panel cannot drift from the export.
/// </summary>
public sealed class GanttTableAppearanceTests
{
    [Fact]
    public void The_value_columns_are_centred_and_the_prose_is_left_aligned()
    {
        Assert.Equal(GanttCellAlignment.Centre, GanttTableAppearance.AlignmentFor("Start"));
        Assert.Equal(GanttCellAlignment.Centre, GanttTableAppearance.AlignmentFor("Finish"));
        Assert.Equal(GanttCellAlignment.Centre, GanttTableAppearance.AlignmentFor("Duration"));
        Assert.Equal(GanttCellAlignment.Left, GanttTableAppearance.AlignmentFor("Type"));
        Assert.Equal(GanttCellAlignment.Left, GanttTableAppearance.AlignmentFor("Description"));
    }

    [Fact]
    public void An_unknown_or_absent_column_reads_as_left_aligned_prose()
    {
        // The safe direction, and the guard's positive cases: a differently-cased name, an
        // engine column, a blank name and no name at all are never centred by accident.
        Assert.Equal(GanttCellAlignment.Left, GanttTableAppearance.AlignmentFor("start"));
        Assert.Equal(GanttCellAlignment.Left, GanttTableAppearance.AlignmentFor("Id"));
        Assert.Equal(GanttCellAlignment.Left, GanttTableAppearance.AlignmentFor(string.Empty));
        Assert.Equal(GanttCellAlignment.Left, GanttTableAppearance.AlignmentFor(null));
    }

    [Fact]
    public void The_centred_columns_all_exist_in_the_schema_and_are_visible_to_the_user()
    {
        // A typo in the list would leave a value column left-aligned with nothing failing, and
        // an engine column in it would centre a cell the user never sees.
        foreach (string name in GanttTableAppearance.CentredColumnNames)
        {
            Assert.True(GanttTableSchema.Default.TryGetColumn(name, out GanttTableColumn? column));
            Assert.NotEqual(GanttColumnAccess.EngineHidden, column!.Access);
        }
    }

    [Fact]
    public void The_font_family_and_sizes_come_from_the_typography_tokens()
    {
        // The catalogues are the authority, not a literal here: changing the token must move
        // the live table with the chart and the exported panel.
        Assert.Equal(GanttCatalogues.TypographyDefaultText("FontFamily"), GanttTableAppearance.FontFamily);
        Assert.Equal(GanttCatalogues.TypographyDefault("BodyFontSizePt"), GanttTableAppearance.BodyFontSizePt);
        Assert.Equal(
            GanttCatalogues.TypographyDefault("HeaderFontSizePt"),
            GanttTableAppearance.HeaderFontSizePt);
    }

    [Fact]
    public void The_header_type_matches_the_period_band_and_is_not_bold()
    {
        // Owner ruling 2026-11-14: the live header row and the chart's period band are one
        // heading, so the header takes the 9pt regular type the band already renders. The
        // bold flag is asserted separately because Excel's table style bolds a header row,
        // so a regular header only exists if the adapter writes the flag.
        Assert.Equal(9, GanttTableAppearance.HeaderFontSizePt);
        Assert.False(GanttTableAppearance.HeaderIsBold);
    }

    [Fact]
    public void The_outside_border_weight_is_the_frame_token()
    {
        // The table's outside edge and the chart frame are one boundary (owner ruling
        // 2026-11-14). The Office adapter writes Excel's xlThin, which IS the 1pt weight,
        // so the token must stay at 1pt for the two boundaries to meet: this fact is the
        // guard that fails if the frame token moves away from the mapping.
        Assert.Equal(
            GanttCatalogues.MetricDefault("MajorBoundaryPt"),
            GanttTableAppearance.OuterBorderWeightPt);
        Assert.Equal(1, GanttTableAppearance.OuterBorderWeightPt);
    }

    [Fact]
    public void The_header_fill_and_the_border_colour_are_catalogue_colours()
    {
        // The owner's ruling takes the fill and the border COLOUR from the chart theme while
        // leaving the border style to Excel's table style, so both must resolve from the
        // colour catalogue rather than from a literal in the writer.
        Assert.Equal(ColourHex.Parse("#FFFFFF"), GanttTableAppearance.HeaderFill);
        Assert.Equal(ColourHex.Parse("#D9D9D9"), GanttTableAppearance.BorderColour);
    }

    [Fact]
    public void The_header_text_colour_never_equals_the_header_fill()
    {
        // Live evidence 2026-11-14: the initialiser wrote the fill alone, so the table
        // style's white header font stayed white ON the white band and every schema name
        // was invisible. The pair is what must be asserted, not either colour alone — a
        // future token change that made them equal would be the same defect again.
        Assert.Equal(ColourHex.Parse("#000000"), GanttTableAppearance.HeaderTextColour);
        Assert.NotEqual(GanttTableAppearance.HeaderFill, GanttTableAppearance.HeaderTextColour);
    }
}
