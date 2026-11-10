using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;
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
    public void Add_row_buttons_use_the_verified_table_insert_image()
    {
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
            Assert.Equal("TableInsertExcel", button.Attribute("imageMso")?.Value);
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
}
