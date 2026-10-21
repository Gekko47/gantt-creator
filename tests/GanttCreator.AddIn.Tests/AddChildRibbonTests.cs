using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the R5.3 WI-E Add child dropdown: the item list is
/// PROJECTED from the hierarchy catalogue, so the dropdown and the Core
/// planner's acceptance cannot disagree, and the control is enabled only for
/// a selection that may own children.
/// </summary>
/// <remarks>
/// The class shares the <c>diagnostics-service</c> collection because
/// <see cref="RibbonStateService"/> is a session singleton armed by
/// <see cref="AddInHost.AutoOpen"/>; xUnit runs a collection's classes
/// sequentially, so the host lifecycle tests cannot race these.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class AddChildRibbonTests : IDisposable
{
    public AddChildRibbonTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Hand-written fake for <see cref="ISelectionIdentifier"/>: returns a
    /// fixed selection and counts subscription attach/detach.
    /// </summary>
    private sealed class FakeSelectionIdentifier : ISelectionIdentifier
    {
        public SelectionFacts Facts { get; set; } = SelectionFacts.None;

        public SelectionFacts Identify() => Facts;

        public IDisposable SubscribeSelectionChanged(Action handler) => new Subscription();

        private sealed class Subscription : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    /// <summary>
    /// Minimal <see cref="IExcelApplicationAdapter"/> that reports a live
    /// workbook, so the shared gate's workbook half is satisfied.
    /// </summary>
    private sealed class FakeApplicationAdapter : IExcelApplicationAdapter
    {
        public bool? HasActiveWorkbook() => true;

        public void ShowStatusBar(string message)
        {
            // Intentionally empty: the status bar is not under test here.
        }

        public IDisposable SubscribeWorkbookStateChanged(Action handler) => new Subscription();

        private sealed class Subscription : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    /// <summary>
    /// Builds one table body row through the normalised DTO shape.
    /// </summary>
    private static GanttRowDto Row(string type) => new(
        1,
        "G-" + new string('0', 31) + 'a',
        null,
        null,
        type,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null);

    /// <summary>
    /// Arms the service with a single-selected body row of the given type.
    /// </summary>
    private static RibbonStateService ArmWithRow(
        FakeSelectionIdentifier selection,
        GanttRowDto row)
    {
        var service = RibbonStateService.Instance;
        service.SetApplicationAdapter(new FakeApplicationAdapter());
        service.SetSelectionIdentifier(selection);

        var reader = new Mock<IGanttTableReader>();
        reader.Setup(item => item.Read()).Returns(GanttTableReadOutcome.Ok([row]));
        service.SetTableReader(reader.Object);

        service.Activate();
        return service;
    }

    private static SelectionFacts SingleRow() => new(true, true, 1, false, 1);

    [Fact]
    public void The_dropdown_is_enabled_for_a_parent_that_may_own_children()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity"));

        Assert.True(service.GetEnabled(RibbonControlIds.AddChild));
    }

    [Fact]
    public void The_dropdown_is_disabled_for_a_type_that_may_not_own_children()
    {
        // A milestone is a leaf: it may be a child, but it never owns one, so
        // the dropdown offers nothing and is greyed.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Milestone"));

        Assert.False(service.GetEnabled(RibbonControlIds.AddChild));
        Assert.Empty(service.GetPermittedChildTypes());
    }

    [Fact]
    public void The_dropdown_is_disabled_when_the_selection_is_not_one_body_row()
    {
        var selection = new FakeSelectionIdentifier
        {
            Facts = new SelectionFacts(true, false, null, true, 0),
        };
        var service = ArmWithRow(selection, Row("As-Built Activity"));

        Assert.False(service.GetEnabled(RibbonControlIds.AddChild));
    }

    [Fact]
    public void The_dropdown_offers_the_catalogue_s_child_capable_types()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity"));

        // PROJECTED, never authored beside the catalogue: the dropdown and the
        // planner cannot disagree.
        Assert.Equal(
            EntityHierarchyCatalog.TypesThatMayBeChild().Count,
            service.GetPermittedChildTypes().Count);
    }

    [Fact]
    public void The_item_id_names_the_child_type_and_the_label_names_it_in_words()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity"));
        IRibbonControl control = new Mock<IRibbonControl>().Object;

        int count = GanttRibbon.GetAddChildCount(control, service);
        Assert.True(count > 0);

        for (int index = 0; index < count; index++)
        {
            string id = GanttRibbon.GetAddChildItemID(control, index, service);
            string label = GanttRibbon.GetAddChildItemLabel(control, index, service);

            Assert.True(
                Enum.TryParse(id, ignoreCase: false, out GanttEntityType parsed),
                $"The item id '{id}' is not a type name.");
            Assert.True(
                EntityHierarchyCatalog.MayBeChild(parsed),
                $"The dropdown offered a type that may not be a child: {parsed}.");
            Assert.False(string.IsNullOrWhiteSpace(label));
        }
    }

    [Fact]
    public void An_out_of_range_index_yields_an_empty_item()
    {
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity"));
        IRibbonControl control = new Mock<IRibbonControl>().Object;

        int count = GanttRibbon.GetAddChildCount(control, service);

        Assert.Equal(string.Empty, GanttRibbon.GetAddChildItemID(control, count, service));
        Assert.Equal(string.Empty, GanttRibbon.GetAddChildItemLabel(control, count, service));
        Assert.Equal(string.Empty, GanttRibbon.GetAddChildItemID(control, -1, service));
        Assert.Equal(string.Empty, GanttRibbon.GetAddChildItemLabel(control, -1, service));
    }

    [Fact]
    public void Every_type_the_dropdown_offers_is_accepted_by_the_planner()
    {
        // The projection is the SAME one the planner uses, so this is the
        // test that keeps the dropdown and the planner agreeing.
        var selection = new FakeSelectionIdentifier { Facts = SingleRow() };
        var service = ArmWithRow(selection, Row("As-Built Activity"));

        GanttRowId parentId = GanttRowId.New();
        List<GanttEvent> rows =
        [
            new(
                1,
                parentId,
                GanttRowId.New(),
                0,
                GanttEntityType.AsPlannedActivity,
                "Build frame",
                new DateOnly(2026, 1, 1),
                new DateOnly(2026, 12, 31),
                null,
                "AsPlannedActivity",
                null,
                null,
                null,
                true,
                null),
        ];

        foreach (GanttEntityType childType in service.GetPermittedChildTypes())
        {
            AddChildPlan plan = AddChildPlanner.Plan(rows, parentId, childType, GanttRowId.New());

            Assert.True(plan.Succeeded, $"The planner refused a type the dropdown offered: {childType}.");
        }
    }

    [Fact]
    public void The_ribbon_declares_the_add_child_dropdown_with_its_callbacks()
    {
        string xml = new GanttRibbon().GetCustomUI("Microsoft.Excel.Workbook")!;
        System.Xml.Linq.XDocument doc = System.Xml.Linq.XDocument.Parse(xml);
        System.Xml.Linq.XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";

        var dropdown = System.Linq.Enumerable.Single(
            System.Linq.Enumerable.Where(
                doc.Descendants(ns + "dropDown"),
                element => element.Attribute("id")?.Value == RibbonControlIds.AddChild));

        Assert.Equal("GetAddChildCount", dropdown.Attribute("getItemCount")?.Value);
        Assert.Equal("GetAddChildItemID", dropdown.Attribute("getItemID")?.Value);
        Assert.Equal("GetAddChildItemLabel", dropdown.Attribute("getItemLabel")?.Value);
        Assert.Equal("OnAddChildChange", dropdown.Attribute("onAction")?.Value);
    }
}
