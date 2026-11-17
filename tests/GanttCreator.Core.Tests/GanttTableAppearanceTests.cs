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
        // 2026-11-14), so both read the SAME token rather than restating a width. That
        // equality is the guard; the value itself is the token's, and since 2026-11-15 the
        // token is 0.25pt (a hairline) because 1pt and 0.5pt both render as one pixel at
        // 100 %. Excel's cell-border weight is a CLOSED enum, so the Office adapter maps
        // the token onto xlThin instead of writing 0.25 as points -- a cell border cannot
        // be finer than that -- and the frame is what carries the token's own width.
        Assert.Equal(
            GanttCatalogues.MetricDefault("MajorBoundaryPt"),
            GanttTableAppearance.OuterBorderWeightPt);
        Assert.Equal(0.25, GanttTableAppearance.OuterBorderWeightPt, 6);
    }

    [Fact]
    public void The_header_fill_is_the_period_bands_and_the_border_colour_is_the_minor_grid()
    {
        // The header row IS the period band's row, so its fill is the period band's
        // (owner ruling 2026-11-14). It previously took the white HeaderFill token while
        // the band beside it rendered the neutral AlternateBandFill, which is the mismatch
        // reported from the host. Reading the CHROME ROLE rather than restating the token
        // is what keeps the band, its setting key and this fill one decision.
        Assert.Equal(
            ColourHex.Parse(GanttChartChrome.DefaultHexFor(GanttChartChromeRole.PeriodHeader)),
            GanttTableAppearance.HeaderFill);

        // Not white any more: this is the reported defect stated as an assertion.
        Assert.NotEqual(ColourHex.Parse("#FFFFFF"), GanttTableAppearance.HeaderFill);

        // The border colour is still the theme's minor grid stroke, so a cell boundary and
        // a grid line stay one colour while the table style owns the border's weight.
        Assert.Equal(ColourHex.Parse("#D9D9D9"), GanttTableAppearance.BorderColour);
    }

    [Fact]
    public void The_table_style_is_the_neutral_builtin_one()
    {
        // Owner ruling 2026-11-14: hiding and re-showing a table column makes Excel re-apply
        // the table's own style, and the default style is not the neutral one the guide
        // describes. The assigned style must therefore be the light, unbanded built-in one,
        // so the host's own re-application keeps the same thin grey lines.
        Assert.Equal("TableStyleLight1", GanttTableAppearance.TableStyleName);
    }

    [Fact]
    public void The_visible_column_default_width_is_sixty_pixels_which_is_forty_five_points()
    {
        // Owner ruling 2026-11-17: a freshly initialised table gives every visible column
        // 60 pixels, the unit Excel's width dialog reports. The first cut targeted 60
        // points and the host showed 80 pixels (60 × 96/72), so the conversion is pinned
        // rather than left to an adapter. A presentation constant, not a schema metric, so
        // it must not move the catalogue hash.
        Assert.Equal(60.0, GanttTableAppearance.DefaultVisibleColumnWidthPx);
        Assert.Equal(45.0, GanttTableAppearance.DefaultVisibleColumnWidthPt, 6);
        Assert.True(double.IsFinite(GanttTableAppearance.DefaultVisibleColumnWidthPt));
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
