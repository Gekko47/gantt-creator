using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Logging;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>R2.8 Ribbon and state contracts for the Data-group buttons.</summary>
public class AddRowRibbonTests
{
    private const string WorkbookRibbonId = "Microsoft.Excel.Workbook";
    private const string CustomUiNamespace = "http://schemas.microsoft.com/office/2009/07/customui";

    [Fact]
    public void Ribbon_declares_the_data_group_controls()
    {
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;
        XElement group = Assert.Single(
            doc.Descendants(ns + "group"),
            g => g.Attribute("id")?.Value == RibbonControlIds.DataGroup);

        // R5.16: four large split buttons plus the large Remove. Initialise
        // moved to the first-run Get started group.
        XElement[] splits = [.. group.Elements(ns + "splitButton")];
        Assert.Equal(
            [
                RibbonControlIds.FixSplit,
                RibbonControlIds.AddActivitySplit,
                RibbonControlIds.AddMilestoneSplit,
                RibbonControlIds.AddElementsSplit,
            ],
            splits.Select(element => element.Attribute("id")?.Value));
        Assert.All(splits, element => Assert.Equal("large", element.Attribute("size")?.Value));
        Assert.All(splits, element => Assert.Single(element.Elements(ns + "button")));
        Assert.All(splits, element => Assert.Single(element.Elements(ns + "menu")));

        // The primary button carries the default action; the menu holds the rest.
        Assert.Equal(
            [
                "OnRepairConfigClick",
                "OnAddActivityClick",
                "OnAddMilestoneClick",
                "OnAddDelineatorClick",
            ],
            splits.Select(element => element.Elements(ns + "button").First().Attribute("onAction")?.Value));

        XElement remove = Assert.Single(
            group.Elements(ns + "button"),
            button => button.Attribute("id")?.Value == RibbonControlIds.Remove);
        Assert.Equal("large", remove.Attribute("size")?.Value);

        Assert.DoesNotContain(
            group.Descendants(ns + "button"),
            button => button.Attribute("id")?.Value == RibbonControlIds.InitialiseSheet);
    }

    [Fact]
    public void Add_row_buttons_carry_their_own_themed_glyphs()
    {
        // R5.19: the three add buttons are served by the themed glyph provider, so each
        // names the shared callback instead of an Office gallery id. What the callback
        // then resolves — a Lucide slug at a declared size, in both palettes — is pinned
        // by RibbonGlyphAssetsTests.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        foreach (string id in new[]
                 {
                     RibbonControlIds.AddActivity,
                     RibbonControlIds.AddMilestone,
                     RibbonControlIds.AddDelineator,
                 })
        {
            XElement button = doc.Descendants(ns + "button")
                .Single(element => element.Attribute("id")?.Value == id);

            // The picture follows the same remembered menu choice the label does, so a
            // button reading "Spacer" cannot draw a delineator.
            Assert.Equal("GetAddEntityPrimaryImage", button.Attribute("getImage")?.Value);
            Assert.Null(button.Attribute("imageMso"));
        }
    }

    [Theory]
    [InlineData("btnAddActivity")]
    [InlineData("btnAddMilestone")]
    [InlineData("btnAddDelineator")]
    public void Add_row_callback_runs_through_boundary_and_refreshes(string controlId)
    {
        var shown = new List<string>();
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        _ = log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>((format, args) =>
                written.Add(string.Format(System.Globalization.CultureInfo.InvariantCulture, format, args)));
        var boundary = new CommandBoundary(presenter: shown.Add);
        boundary.SetLog(log.Object);
        var control = new Mock<IRibbonControl>();
        _ = control.SetupGet(c => c.Id).Returns(controlId);
        var refreshes = 0;
        var executed = false;

        Action command = () => executed = true;
        switch (controlId)
        {
            case RibbonControlIds.AddActivity:
                GanttRibbon.OnAddActivityClick(control.Object, boundary, command, () => refreshes++);
                break;
            case RibbonControlIds.AddMilestone:
                GanttRibbon.OnAddMilestoneClick(control.Object, boundary, command, () => refreshes++);
                break;
            default:
                GanttRibbon.OnAddDelineatorClick(control.Object, boundary, command, () => refreshes++);
                break;
        }

        Assert.True(executed);
        Assert.Empty(written);
        Assert.Empty(shown);
        Assert.Equal(1, refreshes);
        Assert.Equal(controlId, GanttRibbon.ResolveCommandName(control.Object));
    }

    [Fact]
    public void Add_row_callback_failure_uses_control_id_and_still_refreshes()
    {
        var written = new List<string>();
        var log = new Mock<IRollingLog>();
        _ = log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>((format, args) =>
                written.Add(string.Format(System.Globalization.CultureInfo.InvariantCulture, format, args)));
        var boundary = new CommandBoundary(presenter: _ => { });
        boundary.SetLog(log.Object);
        var control = new Mock<IRibbonControl>();
        _ = control.SetupGet(c => c.Id).Returns(RibbonControlIds.AddMilestone);
        var refreshes = 0;

        GanttRibbon.OnAddMilestoneClick(
            control.Object,
            boundary,
            () => throw new InvalidOperationException("simulated"),
            () => refreshes++);

        Assert.Contains("command=btnAddMilestone", Assert.Single(written), StringComparison.Ordinal);
        Assert.Equal(1, refreshes);
    }

    [Fact]
    public void Add_row_controls_follow_the_cached_workbook_state()
    {
        var adapter = new Mock<IExcelApplicationAdapter>();
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(true);
        var tableReader = new Mock<IGanttTableReader>();
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        var service = new RibbonStateService();
        service.SetApplicationAdapter(adapter.Object);
        service.SetLogAvailabilitySource(() => false);
        service.SetTableReader(tableReader.Object);
        service.Refresh();

        Assert.True(service.GetEnabled(RibbonControlIds.AddActivity));
        Assert.True(service.GetEnabled(RibbonControlIds.AddMilestone));
        Assert.True(service.GetEnabled(RibbonControlIds.AddDelineator));

        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(false);
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.AddActivity));
        Assert.False(service.GetEnabled(RibbonControlIds.AddMilestone));
        Assert.False(service.GetEnabled(RibbonControlIds.AddDelineator));
    }

    [Fact]
    public void Add_row_primaries_take_their_label_from_the_catalogue()
    {
        // R5.18g: a split button whose primary keeps a static label lies as soon as the
        // menu changes what that primary adds. The caption is therefore projected from
        // EntityTypeCatalog through getLabel and never authored in the Ribbon XML.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        string[] primaries =
        [
            RibbonControlIds.AddActivity,
            RibbonControlIds.AddMilestone,
            RibbonControlIds.AddDelineator,
        ];

        foreach (string id in primaries)
        {
            XElement button = doc.Descendants(ns + "button")
                .Single(element => element.Attribute("id")?.Value == id);
            Assert.Null(button.Attribute("label"));
            Assert.Equal("GetAddEntityPrimaryLabel", button.Attribute("getLabel")?.Value);
        }
    }

    [Fact]
    public void Every_menu_item_belongs_to_the_family_that_declares_it()
    {
        // The family map is the state's view of the menu structure: a remembered pick must
        // reach the primary of the menu the user actually picked it from. If the XML gains
        // an item the map does not claim (or the catalogue gains a type no menu offers),
        // remembering the pick would silently stop following it back. Both directions are
        // pinned here so neither list can drift alone.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        (string MenuId, RibbonStateService.AddEntityFamily Family)[] menus =
        [
            (RibbonControlIds.AddActivityMenu, RibbonStateService.AddEntityFamily.Activity),
            (RibbonControlIds.AddMilestoneMenu, RibbonStateService.AddEntityFamily.Milestone),
            (RibbonControlIds.AddElementsMenu, RibbonStateService.AddEntityFamily.Element),
        ];

        List<GanttEntityType> offered = [];
        foreach ((string menuId, RibbonStateService.AddEntityFamily family) in menus)
        {
            XElement menu = doc.Descendants(ns + "menu")
                .Single(element => element.Attribute("id")?.Value == menuId);
            foreach (XElement item in menu.Elements(ns + "button"))
            {
                Assert.True(
                    GanttRibbon.TryParseEntityMenuId(item.Attribute("id")?.Value, out GanttEntityType type),
                    $"{menuId} declares an item that is not a catalogue entity type.");
                Assert.Equal(family, RibbonStateService.FamilyOf(type));
                offered.Add(type);
            }
        }

        List<GanttEntityType> catalogue =
            [.. EntityTypeCatalog.Entries.Select(entry => entry.Type)];
        Assert.Equal(
            [.. catalogue.Order()],
            [.. offered.Order()]);
    }
}
