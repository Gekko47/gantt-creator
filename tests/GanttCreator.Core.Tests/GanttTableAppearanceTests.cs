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
    public void The_font_family_and_size_come_from_the_typography_tokens()
    {
        // The catalogues are the authority, not a literal here: changing the token must move
        // the live table with the chart and the exported panel.
        Assert.Equal(GanttCatalogues.TypographyDefaultText("FontFamily"), GanttTableAppearance.FontFamily);
        Assert.Equal(GanttCatalogues.TypographyDefault("BodyFontSizePt"), GanttTableAppearance.BodyFontSizePt);
    }
}
