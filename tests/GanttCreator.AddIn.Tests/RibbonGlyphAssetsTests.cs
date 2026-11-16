using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The asset half of the glyph contract: every catalogued control has a picture in both
/// palettes at its declared size, and nothing on the path throws.
/// </summary>
public sealed class RibbonGlyphAssetsTests
{
    private const string _workbookRibbonId = "Microsoft.Excel.Workbook";

    private static readonly XNamespace _customUiNamespace =
        "http://schemas.microsoft.com/office/2009/07/customui";

    private static readonly RibbonGlyphPalette[] _palettes =
    [
        RibbonGlyphPalette.Light,
        RibbonGlyphPalette.Dark,
    ];

    [Fact]
    public void The_suffix_names_the_slug_the_palette_and_the_size()
    {
        Assert.Equal(
            "refresh-chart-light-32.png",
            RibbonGlyphAssets.Suffix("refresh-chart", RibbonGlyphPalette.Light, 32));

        Assert.Equal(
            "refresh-chart-dark-16.png",
            RibbonGlyphAssets.Suffix("refresh-chart", RibbonGlyphPalette.Dark, 16));
    }

    [Fact]
    public void A_null_slug_is_refused() =>
        Assert.Throws<ArgumentNullException>(
            () => RibbonGlyphAssets.Suffix(null!, RibbonGlyphPalette.Light, 16));

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void A_blank_slug_is_refused(string slug) =>
        Assert.Throws<ArgumentException>(
            () => RibbonGlyphAssets.Suffix(slug, RibbonGlyphPalette.Light, 16));

    [Fact]
    public void Every_catalogued_control_has_both_palettes_at_its_declared_size()
    {
        List<string> missing = [];

        foreach (KeyValuePair<string, RibbonGlyphCatalog.Entry> pair in RibbonGlyphCatalog.Controls)
        {
            foreach (RibbonGlyphPalette palette in _palettes)
            {
                if (RibbonGlyphAssets.ReadPng(pair.Value.Glyph, palette, pair.Value.Size) is null)
                {
                    missing.Add($"{pair.Key} -> {pair.Value.Glyph} {palette} {pair.Value.Size}");
                }
            }
        }

        Assert.Empty(missing);
    }

    [Theory]
    [InlineData(RibbonGlyphCatalog.ColumnVisible)]
    [InlineData(RibbonGlyphCatalog.ColumnHidden)]
    [InlineData(RibbonGlyphCatalog.CellLockClosed)]
    [InlineData(RibbonGlyphCatalog.CellLockOpen)]
    [InlineData(RibbonGlyphCatalog.ChartLockClosed)]
    [InlineData(RibbonGlyphCatalog.ChartLockOpen)]
    public void Every_state_glyph_has_both_palettes_at_sixteen_pixels(string slug)
    {
        foreach (RibbonGlyphPalette palette in _palettes)
        {
            Assert.NotNull(RibbonGlyphAssets.ReadPng(slug, palette, 16));
        }
    }

    [Fact]
    public void A_catalogued_control_resolves_to_a_picture()
    {
        Assert.NotNull(RibbonGlyphAssets.ForControl("btnRefreshSheet"));
        Assert.NotNull(RibbonGlyphAssets.ForControl("btnWarnings"));
        Assert.NotNull(RibbonGlyphAssets.ForControl("mnuColumnVisibility"));
    }

    [Fact]
    public void An_unmapped_control_has_no_picture() =>
        Assert.Null(RibbonGlyphAssets.ForControl("btnDoesNotExist"));

    [Fact]
    public void A_null_control_has_no_picture() =>
        Assert.Null(RibbonGlyphAssets.ForControl(null));

    [Fact]
    public void The_callback_passes_the_control_id_to_the_resolver()
    {
        GlyphStubControl control = new("btnRefreshSheet");
        string? observed = null;

        object? picture = GanttRibbon.GetGlyph(
            control,
            id =>
            {
                observed = id;
                return null;
            });

        Assert.Null(picture);
        Assert.Equal("btnRefreshSheet", observed);
    }

    [Fact]
    public void The_callback_tolerates_a_null_control()
    {
        string? observed = "not-called";

        object? picture = GanttRibbon.GetGlyph(
            null,
            id =>
            {
                observed = id;
                return null;
            });

        Assert.Null(picture);
        Assert.Null(observed);
    }

    [Fact]
    public void The_ribbon_and_the_catalogue_agree_on_every_control()
    {
        // The drift gate, in process as well as in scripts/check-ribbon-glyphs.ps1: a
        // control that declares getImage without a catalogue entry renders blank, and a
        // catalogue entry the ribbon never uses is a stale asset.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(_workbookRibbonId));

        string[] declared =
        [
            .. doc.Descendants()
                .Where(element => (string?)element.Attribute("getImage") == "GetGlyph")
                .Select(element => (string?)element.Attribute("id"))
                .Where(id => id is not null)
                .Select(id => id!),
        ];

        Assert.Equal(
            declared.OrderBy(id => id, StringComparer.Ordinal),
            RibbonGlyphCatalog.Controls.Keys.OrderBy(id => id, StringComparer.Ordinal));
    }

    [Fact]
    public void The_ribbon_declares_no_gallery_ids()
    {
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(_workbookRibbonId));

        Assert.DoesNotContain(
            doc.Descendants(),
            element => element.Attribute("imageMso") is not null);
    }

    [Fact]
    public void The_show_columns_menu_carries_a_glyph()
    {
        // mnuColumnVisibility was the one large control that shipped with no glyph at all
        // before R5.19, so it is pinned rather than left to the completeness test alone.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(_workbookRibbonId));

        XElement menu = doc.Descendants(_customUiNamespace + "menu")
            .Single(element => (string?)element.Attribute("id") == "mnuColumnVisibility");

        Assert.Equal("GetGlyph", (string?)menu.Attribute("getImage"));
    }

    [Fact]
    public void The_size_dropdown_carries_no_glyph()
    {
        // The owner ruled that Size should not show an icon: it is a text dropdown beside
        // an inline label, and a picture there only crowds the row. CT_Item has no
        // getImage either, so the four preset rows carry no picture at all.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(_workbookRibbonId));

        XElement dropDown = doc.Descendants(_customUiNamespace + "dropDown")
            .Single(element => (string?)element.Attribute("id") == "ddnSize");

        Assert.Null(dropDown.Attribute("getImage"));
        Assert.Null(dropDown.Attribute("imageMso"));

        Assert.All(
            dropDown.Elements(_customUiNamespace + "item"),
            item => Assert.Null(item.Attribute("imageMso")));
    }

    [Fact]
    public void The_add_primaries_follow_their_label_with_their_glyph()
    {
        // R5.18g made the three primaries follow the last menu pick in their LABEL. A
        // static glyph beside a changing label is the original defect one channel over, so
        // the glyph is served by a callback that reads the same remembered choice.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(_workbookRibbonId));

        foreach (string id in new[] { "btnAddActivity", "btnAddMilestone", "btnAddDelineator" })
        {
            XElement button = doc.Descendants(_customUiNamespace + "button")
                .Single(element => (string?)element.Attribute("id") == id);

            Assert.Equal("GetAddEntityPrimaryImage", (string?)button.Attribute("getImage"));
            Assert.Equal("GetAddEntityPrimaryLabel", (string?)button.Attribute("getLabel"));
        }
    }

    [Fact]
    public void The_element_primary_glyph_names_the_element_its_label_names()
    {
        RibbonStateService state = new();

        (GanttEntityType Type, string Slug)[] expected =
        [
            (GanttEntityType.Delineator, "delineator"),
            (GanttEntityType.Spacer, "spacer"),
            (GanttEntityType.Splitter, "splitter"),
        ];

        foreach ((GanttEntityType type, string slug) in expected)
        {
            state.RememberAddedType(type);

            Assert.Equal(
                slug,
                RibbonStateService.AddEntityPrimaryGlyph(
                    state.PrimaryTypeFor(RibbonStateService.AddEntityFamily.Element)));

            Assert.Equal(
                EntityTypeCatalog.Format(type),
                RibbonStateService.AddEntityPrimaryLabel("btnAddDelineator", state));

            Assert.NotNull(RibbonStateService.AddEntityPrimaryImage("btnAddDelineator", state));
        }
    }

    [Fact]
    public void Every_entity_type_maps_into_the_five_primary_glyphs()
    {
        // The checker verifies these five assets at 32 px directly, because they are
        // reached through the primary callback rather than through the catalogue. This
        // test is what keeps that list honest.
        string[] slugs =
        [
            .. Enum.GetValues<GanttEntityType>()
                .Select(RibbonStateService.AddEntityPrimaryGlyph)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(slug => slug, StringComparer.Ordinal),
        ];

        Assert.Equal(
            ["add-activity", "add-milestone", "delineator", "spacer", "splitter"],
            slugs);
    }

    [Fact]
    public void Activity_and_milestone_variants_keep_one_family_glyph()
    {
        // Their variants differ by timing, not by shape, so only the element family splits
        // its glyph. The label still names the exact type.
        Assert.Equal("add-activity", RibbonStateService.AddEntityPrimaryGlyph(GanttEntityType.AsBuiltActivity));
        Assert.Equal("add-activity", RibbonStateService.AddEntityPrimaryGlyph(GanttEntityType.CriticalInterval));
        Assert.Equal("add-milestone", RibbonStateService.AddEntityPrimaryGlyph(GanttEntityType.CriticalMilestone));
        Assert.Equal("add-milestone", RibbonStateService.AddEntityPrimaryGlyph(GanttEntityType.CustomMilestone));
    }

    // Ported from ColumnVisibilityImagesTests: the column menu offers the five authoring
    // columns, never an engine column, so an id outside that set must not resolve to a
    // picture.
    [Theory]
    [InlineData("colId")]
    [InlineData("col")]
    [InlineData("colNope")]
    [InlineData("")]
    public void An_engine_column_never_has_a_glyph(string id) =>
        Assert.Null(RibbonGlyphAssets.ForControl(id));

    [Theory]
    [InlineData("btnToggleCellLock", "GetToggleCellLockImage")]
    [InlineData("btnToggleChartLock", "GetToggleChartLockImage")]
    public void The_lock_toggles_serve_a_state_glyph_rather_than_a_gallery_id(string id, string callback)
    {
        // Ported from LockGlyphImagesTests when that wrapper was deleted in R5.19: the
        // contract it guarded — a toggle names its own image callback and carries no
        // imageMso — is still worth pinning, and it now applies to the themed provider.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(_workbookRibbonId));

        XElement toggle = doc.Descendants(_customUiNamespace + "toggleButton")
            .Single(element => element.Attribute("id")?.Value == id);

        Assert.Equal(callback, (string?)toggle.Attribute("getImage"));
        Assert.Null(toggle.Attribute("imageMso"));
    }

    [Theory]
    [InlineData("colType")]
    [InlineData("colDescription")]
    [InlineData("colStart")]
    [InlineData("colFinish")]
    [InlineData("colDuration")]
    public void Column_rows_keep_their_dynamic_tick_and_dash(string id)
    {
        // Ported from ColumnVisibilityImagesTests when that wrapper was deleted in R5.19.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(_workbookRibbonId));

        XElement row = doc.Descendants(_customUiNamespace + "button")
            .Single(element => element.Attribute("id")?.Value == id);

        Assert.Equal("GetColumnVisibilityImage", (string?)row.Attribute("getImage"));
        Assert.Null(row.Attribute("imageMso"));
    }

    private sealed class GlyphStubControl(string id) : IRibbonControl
    {
        public string Id => id;

        public string Tag => string.Empty;

        public object? Context => null;
    }
}
