using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core.Logging;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>R5.3 ribbon and state contracts for the parent/child commands (D5, D6).</summary>
public class ParentChildRibbonTests
{
    private const string WorkbookRibbonId = "Microsoft.Excel.Workbook";
    private const string CustomUiNamespace = "http://schemas.microsoft.com/office/2009/07/customui";

    [Fact]
    public void Ribbon_declares_both_parent_child_buttons_in_a_dedicated_ordering_group()
    {
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;
        XElement group = Assert.Single(
            doc.Descendants(ns + "group"),
            g => g.Attribute("id")?.Value == "grpOrdering");
        XElement[] buttons = group.Elements(ns + "button").ToArray();

        Assert.Equal(
            [
                RibbonControlIds.PromoteToParent,
                RibbonControlIds.DemoteToChild,
            ],
            buttons.Select(button => button.Attribute("id")?.Value));
        Assert.Equal(
            [
                "OnPromoteToParentClick",
                "OnDemoteToChildClick",
            ],
            buttons.Select(button => button.Attribute("onAction")?.Value));
        Assert.All(buttons, button => Assert.Equal("GetEnabled", button.Attribute("getEnabled")?.Value));
    }

    [Fact]
    public void Ordering_buttons_use_the_verified_table_insert_image()
    {
        // R5.9 owns final icon selection; until then the buttons reuse the
        // imageMso the add-row buttons already verified against the Office
        // 2010 Icons Gallery and the Excel idMso table (R2.2 D8 protocol) —
        // never an unverified guess (R5.3 stop condition).
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        foreach (string id in new[]
        {
            RibbonControlIds.PromoteToParent,
            RibbonControlIds.DemoteToChild,
        })
        {
            XElement button = doc.Descendants(ns + "button")
                .Single(element => element.Attribute("id")?.Value == id);
            Assert.Equal("TableInsertExcel", button.Attribute("imageMso")?.Value);
        }
    }

    [Theory]
    [InlineData("btnPromoteToParent")]
    [InlineData("btnDemoteToChild")]
    public void Parent_child_callback_runs_through_boundary_and_refreshes(string controlId)
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
            case RibbonControlIds.PromoteToParent:
                GanttRibbon.OnPromoteToParentClick(control.Object, boundary, command, () => refreshes++);
                break;
            default:
                GanttRibbon.OnDemoteToChildClick(control.Object, boundary, command, () => refreshes++);
                break;
        }

        Assert.True(executed);
        Assert.Empty(written);
        Assert.Empty(shown);
        Assert.Equal(1, refreshes);
    }

    [Theory]
    [InlineData("btnPromoteToParent")]
    [InlineData("btnDemoteToChild")]
    public void Parent_child_callback_failure_uses_control_id_and_still_refreshes(string controlId)
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
        switch (controlId)
        {
            case RibbonControlIds.PromoteToParent:
                GanttRibbon.OnPromoteToParentClick(control.Object, boundary, failing, () => refreshes++);
                break;
            default:
                GanttRibbon.OnDemoteToChildClick(control.Object, boundary, failing, () => refreshes++);
                break;
        }

        Assert.Contains(
            $"command={controlId}",
            Assert.Single(written),
            StringComparison.Ordinal);
        Assert.Equal(1, refreshes);
    }

    [Fact]
    public void Parent_child_controls_follow_the_cached_workbook_and_initialised_state()
    {
        // D5: the buttons gate on HasActiveWorkbook && SheetInitialised —
        // exactly like the add-row buttons. A wrong-kind selection never
        // greys; it refuses at execution time with a message.
        var adapter = new Mock<IExcelApplicationAdapter>();
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(true);
        var tableReader = new Mock<IGanttTableReader>();
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        var service = new RibbonStateService();
        service.SetApplicationAdapter(adapter.Object);
        service.SetLogAvailabilitySource(() => false);
        service.SetTableReader(tableReader.Object);
        service.Refresh();

        Assert.True(service.GetEnabled(RibbonControlIds.PromoteToParent));
        Assert.True(service.GetEnabled(RibbonControlIds.DemoteToChild));

        // Workbook present but the table missing: not initialised, so both grey.
        _ = tableReader.Setup(t => t.Read())
            .Returns(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing));
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.PromoteToParent));
        Assert.False(service.GetEnabled(RibbonControlIds.DemoteToChild));

        // No workbook: both grey.
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(false);
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.PromoteToParent));
        Assert.False(service.GetEnabled(RibbonControlIds.DemoteToChild));
    }
}
