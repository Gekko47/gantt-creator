using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core.Logging;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R5.4 ribbon contracts for the collapse-all and expand-all buttons: the
/// Ribbon surface over the native Excel outline.
/// </summary>
public class CollapseGroupsRibbonTests
{
    private const string WorkbookRibbonId = "Microsoft.Excel.Workbook";
    private const string CustomUiNamespace = "http://schemas.microsoft.com/office/2009/07/customui";

    [Fact]
    public void Ribbon_declares_both_outline_buttons_in_a_dedicated_outline_group()
    {
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        XElement group = Assert.Single(
            doc.Descendants(ns + "group"),
            g => g.Attribute("id")?.Value == "grpOutline");

        // The two buttons sit side by side in one horizontal box, so the pair
        // reads as the two halves of a single action rather than as two
        // unrelated commands.
        XElement box = Assert.Single(group.Elements(ns + "box"));
        XElement[] buttons = box.Elements(ns + "button").ToArray();

        Assert.Equal(
            [
                RibbonControlIds.ExpandAllGroups,
                RibbonControlIds.CollapseAllGroups,
            ],
            buttons.Select(button => button.Attribute("id")?.Value));
        Assert.Equal(
            [
                "OnExpandAllGroupsClick",
                "OnCollapseAllGroupsClick",
            ],
            buttons.Select(button => button.Attribute("onAction")?.Value));
        Assert.All(buttons, button => Assert.Equal("GetEnabled", button.Attribute("getEnabled")?.Value));
    }

    [Fact]
    public void The_outline_group_sits_beside_the_ordering_group()
    {
        // The Outline group is a hierarchy peer of the Ordering group, so it
        // follows it rather than being buried at the far end of the tab.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        string[] groupIds = doc.Descendants(ns + "group")
            .Select(group => group.Attribute("id")?.Value ?? string.Empty)
            .ToArray();

        int ordering = Array.IndexOf(groupIds, "grpOrdering");
        int outline = Array.IndexOf(groupIds, "grpOutline");

        Assert.True(ordering >= 0, "The Ordering group must exist.");
        Assert.True(outline >= 0, "The Outline group must exist.");
        Assert.True(
            outline > ordering,
            $"The Outline group must follow the Ordering group, but the order was {string.Join(", ", groupIds)}.");
    }

    [Fact]
    public void The_outline_buttons_carry_their_verified_icons()
    {
        // Both IDs satisfy the two bases the R2.2 D8 protocol requires: the
        // Office 2010 Icons Gallery (customUI14.xml, 7344 unique IDs) and the
        // Excel idMso table of [MS-CUSTOMUI]-250218, where the rows are
        // idMso=OutlineExpand, control=button, label="Expand" and
        // idMso=OutlineCollapse, control=button, label="Collapse".
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        string? ImageMso(string id) => doc.Descendants(ns + "button")
            .Single(button => button.Attribute("id")?.Value == id)
            .Attribute("imageMso")?.Value;

        Assert.Equal("OutlineExpand", ImageMso(RibbonControlIds.ExpandAllGroups));
        Assert.Equal("OutlineCollapse", ImageMso(RibbonControlIds.CollapseAllGroups));
    }

    [Theory]
    [InlineData(RibbonControlIds.ExpandAllGroups)]
    [InlineData(RibbonControlIds.CollapseAllGroups)]
    public void The_outline_callback_runs_through_the_boundary_and_refreshes(string controlId)
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
        if (controlId == RibbonControlIds.ExpandAllGroups)
        {
            GanttRibbon.OnExpandAllGroupsClick(control.Object, boundary, command, () => refreshes++);
        }
        else
        {
            GanttRibbon.OnCollapseAllGroupsClick(control.Object, boundary, command, () => refreshes++);
        }

        Assert.True(executed);
        Assert.Empty(written);
        Assert.Empty(shown);
        Assert.Equal(1, refreshes);
    }

    [Theory]
    [InlineData(RibbonControlIds.ExpandAllGroups)]
    [InlineData(RibbonControlIds.CollapseAllGroups)]
    public void The_outline_callback_failure_uses_the_control_id_and_still_refreshes(string controlId)
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
        _ = control.SetupGet(c => c.Id).Returns(controlId);
        var refreshes = 0;

        Action failing = () => throw new InvalidOperationException("simulated");
        if (controlId == RibbonControlIds.ExpandAllGroups)
        {
            GanttRibbon.OnExpandAllGroupsClick(control.Object, boundary, failing, () => refreshes++);
        }
        else
        {
            GanttRibbon.OnCollapseAllGroupsClick(control.Object, boundary, failing, () => refreshes++);
        }

        Assert.Single(written);
        Assert.Contains(controlId, written[0], StringComparison.Ordinal);
        Assert.Equal(1, refreshes);
    }
}
