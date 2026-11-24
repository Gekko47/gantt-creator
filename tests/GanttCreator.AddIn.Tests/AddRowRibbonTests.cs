using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Logging;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>R2.8 Ribbon and state contracts for the Data-group buttons.</summary>
// Shares the diagnostics-service collection: this class drives product callbacks that
// resolve the process-wide CommandBoundary/DiagnosticsService singletons at call time, so
// running it beside the host-lifecycle tests injects technical records into their logs.
[Collection("diagnostics-service")]
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

    [Fact]
    public void Every_entity_menu_is_a_list_of_titled_sections()
    {
        // Owner ruling 2026-11-17: a split-button primary renames itself to the last menu
        // pick, so the button cannot say what the menu's first block is. Every menu
        // therefore OPENS with a titled divider, every later section is titled too, and a
        // separator is a heading rather than a control: no keytip, no action.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        (string MenuId, string[] Expected)[] menus =
        [
            (
                RibbonControlIds.AddActivityMenu,
                [
                    "sep:Activity",
                    "ent:entAsPlannedActivity",
                    "ent:entAsBuiltActivity",
                    "ent:entBaselineActivity",
                    "ent:entCustomActivity",
                    "sep:Events",
                    "ent:entDelayEvent",
                    "ent:entCriticalInterval",
                    "sep:Procurement",
                    "ent:entAsPlannedProcurement",
                    "ent:entAsBuiltProcurement",
                    "ent:entBaselineProcurement",
                ]),
            (
                RibbonControlIds.AddMilestoneMenu,
                [
                    "sep:Milestone",
                    "ent:entAsPlannedMilestone",
                    "ent:entAsBuiltMilestone",
                    "ent:entBaselineMilestone",
                    "ent:entCriticalMilestone",
                    "sep:Custom",
                    "ent:entCustomMilestone",
                ]),
            (
                RibbonControlIds.AddElementsMenu,
                [
                    "sep:Chart marker",
                    "ent:entDelineator",
                    "sep:Table structure",
                    "ent:entSpacer",
                    "ent:entSplitter",
                ]),
        ];

        string[] approvedTitles =
        [
            "Activity",
            "Events",
            "Procurement",
            "Milestone",
            "Custom",
            "Chart marker",
            "Table structure",
        ];

        foreach ((string menuId, string[] expected) in menus)
        {
            XElement menu = doc.Descendants(ns + "menu")
                .Single(element => element.Attribute("id")?.Value == menuId);

            Assert.Equal("menuSeparator", menu.Elements().First().Name.LocalName);
            Assert.Equal(
                expected,
                menu.Elements().Select(child => child.Name.LocalName == "menuSeparator"
                    ? "sep:" + child.Attribute("title")?.Value
                    : "ent:" + child.Attribute("id")?.Value));

            foreach (XElement separator in menu.Elements(ns + "menuSeparator"))
            {
                Assert.Contains(separator.Attribute("title")?.Value, approvedTitles);
                Assert.Null(separator.Attribute("keytip"));
                Assert.Null(separator.Attribute("onAction"));
            }
        }
    }

    [Fact]
    public void Entity_menu_keytips_read_as_basis_plus_kind()
    {
        // The scheme is basis + kind: AP/AB/BL/CU/CR followed by A activity, P procurement,
        // M milestone or I interval. Three characters is the keytip attribute's maximum, so
        // the scheme replaced the old two-letter mixture rather than extending it. A codepoint
        // that breaks the rule (CIM for Critical Milestone, TD for Delineator) reads as a new
        // concept and is pinned out by this fact.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        (string Id, string Keytip)[] expected =
        [
            ("entAsPlannedActivity", "APA"),
            ("entAsBuiltActivity", "ABA"),
            ("entBaselineActivity", "BLA"),
            ("entCustomActivity", "CUA"),
            ("entDelayEvent", "DLY"),
            ("entCriticalInterval", "CRI"),
            ("entAsPlannedProcurement", "APP"),
            ("entAsBuiltProcurement", "ABP"),
            ("entBaselineProcurement", "BLP"),
            ("entAsPlannedMilestone", "APM"),
            ("entAsBuiltMilestone", "ABM"),
            ("entBaselineMilestone", "BLM"),
            ("entCriticalMilestone", "CRM"),
            ("entCustomMilestone", "CUM"),
            ("entDelineator", "DEL"),
            ("entSpacer", "SPA"),
            ("entSplitter", "SPL"),
        ];

        foreach ((string id, string keytip) in expected)
        {
            XElement button = doc.Descendants(ns + "button")
                .Single(element => element.Attribute("id")?.Value == id);

            Assert.Equal(keytip, button.Attribute("keytip")?.Value);
        }
    }

    [Fact]
    public void Entity_menu_tooltips_state_the_differentiator_not_the_label()
    {
        // A screentip that restates the label and a supertip that restates the screentip
        // spend the only space there is on nothing. Each row states what it draws and the
        // one fact that changes the choice, in at most 120 characters, and none begins
        // "Adds " because the menu already says Add.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        (string Id, string Screentip, string Supertip)[] expected =
        [
            (
                "entAsPlannedActivity",
                "Planned span",
                "The planned programme span. Reads Start and Finish; may share its lane with actual and baseline rows."),
            (
                "entAsBuiltActivity",
                "Actual span",
                "The actual execution span. Reads Start and Finish; critical intervals point at it by parent."),
            (
                "entBaselineActivity",
                "Baseline span",
                "The approved baseline span. Reads Start and Finish; the stack index decides which bar sits where."),
            (
                "entCustomActivity",
                "Style-defined span",
                "Needs a style key from the Style group. Reads Start and Finish; an unknown key is an error, never a fallback."),
            (
                "entDelayEvent",
                "Delay span",
                "An explicit delay period. Reads Start and Finish; drawn in its own red body with the label inside by default."),
            (
                "entCriticalInterval",
                "Critical portion",
                "Reads its own Start and Finish. Attaches to the nearest activity that can own it; disabled when none can."),
            (
                "entAsPlannedProcurement",
                "Planned procurement",
                "A planned procurement period. Reads Start and Finish; drawn hatched instead of solid."),
            (
                "entAsBuiltProcurement",
                "Actual procurement",
                "An actual procurement period. Reads Start and Finish; drawn hatched instead of solid."),
            (
                "entBaselineProcurement",
                "Baseline procurement",
                "A baseline procurement period. Reads Start and Finish; drawn hatched instead of solid."),
            (
                "entAsPlannedMilestone",
                "Planned milestone",
                "A planned milestone. A four-point diamond at its Start date; Finish is never read."),
            (
                "entAsBuiltMilestone",
                "Actual milestone",
                "An actual milestone. A four-point diamond at its Start date; Finish is never read."),
            (
                "entBaselineMilestone",
                "Baseline milestone",
                "A baseline milestone. A four-point diamond at its Start date; Finish is never read."),
            (
                "entCriticalMilestone",
                "Critical milestone",
                "A critical milestone. A four-point diamond at its Start date; it takes no styling from a nearby interval."),
            (
                "entCustomMilestone",
                "Style-defined milestone",
                "Needs a style key from the Style group. A diamond at its Start date; an unknown key is an error, never a fallback."),
            (
                "entDelineator",
                "Full-height date line",
                "A full-height line at its Start date. Finish, lane and stack are kept but flagged as not used."),
            (
                "entSpacer",
                "Blank spacer row",
                "A blank row that only adds vertical space. No chart shape; any dates on it are ignored with a warning."),
            (
                "entSplitter",
                "Section band",
                "A named band across the table and plot that separates sections. No date geometry; the description is the band label."),
        ];

        foreach ((string id, string screentip, string supertip) in expected)
        {
            XElement button = doc.Descendants(ns + "button")
                .Single(element => element.Attribute("id")?.Value == id);

            Assert.Equal(screentip, button.Attribute("screentip")?.Value);
            Assert.Equal(supertip, button.Attribute("supertip")?.Value);
            Assert.False(supertip.StartsWith("Adds ", StringComparison.Ordinal));
            Assert.InRange(supertip.Length, 1, 120);
        }
    }
}
