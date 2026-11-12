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
    public void Ribbon_declares_both_parent_child_buttons_in_the_outline_group()
    {
        // R5.16a: the former Ordering group is merged into Outline. Promote
        // and Demote are that group's first row, exactly as they were the
        // Ordering group's whole content.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;
        XElement group = Assert.Single(
            doc.Descendants(ns + "group"),
            g => g.Attribute("id")?.Value == RibbonControlIds.OutlineGroup);

        XElement row = Assert.Single(
            group.Elements(ns + "box"),
            box => box.Attribute("id")?.Value == "boxOutlineRow1");
        XElement[] buttons = row.Elements(ns + "button").ToArray();

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
    public void Outline_buttons_use_excel_outline_and_sort_images()
    {
        // R5.16a: the pair no longer borrows the table glyph. Each id is in
        // the Office 2010 Icons Gallery and in the Excel block of the
        // [MS-CUSTOMUI] idMso table: OutlineGroup = "Group...",
        // OutlineUngroup = "Ungroup...", SortAscendingExcel =
        // "Sort Ascending", SortDescendingExcel = "Sort Descending".
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        string? ImageMso(string id) => doc.Descendants(ns + "button")
            .Single(element => element.Attribute("id")?.Value == id)
            .Attribute("imageMso")?.Value;

        Assert.Equal("OutlineUngroup", ImageMso(RibbonControlIds.PromoteToParent));
        Assert.Equal("OutlineGroup", ImageMso(RibbonControlIds.DemoteToChild));
        Assert.Equal("SortAscendingExcel", ImageMso(RibbonControlIds.MoveUp));
        Assert.Equal("SortDescendingExcel", ImageMso(RibbonControlIds.MoveDown));
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
    public void Parent_child_controls_stay_disabled_without_a_resolved_selection()
    {
        // R5.9b: the buttons now also require a row under the active cell, so an
        // unresolved selection greys them instead of raising a refusal dialog.
        var adapter = new Mock<IExcelApplicationAdapter>();
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(true);
        var tableReader = new Mock<IGanttTableReader>();
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        var service = new RibbonStateService();
        service.SetApplicationAdapter(adapter.Object);
        service.SetLogAvailabilitySource(() => false);
        service.SetTableReader(tableReader.Object);
        var identifier = new Mock<ISelectionIdentifier>();
        _ = identifier.Setup(i => i.Identify()).Returns(SelectionFacts.None);
        service.SetSelectionIdentifier(identifier.Object);
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.PromoteToParent));
        Assert.False(service.GetEnabled(RibbonControlIds.DemoteToChild));

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

    [Fact]
    public void Move_buttons_sit_in_a_two_column_row_of_the_outline_group()
    {
        // R5.16a: the Outline group is a two-wide, three-storey grid. Row 2
        // is a horizontal box of exactly two columns — Up then Down.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        XElement box = Assert.Single(
            doc.Descendants(ns + "box"),
            b => b.Attribute("id")?.Value == "boxOutlineRow2");
        Assert.Equal("horizontal", box.Attribute("boxStyle")?.Value);
        Assert.Equal(
            [
                RibbonControlIds.MoveUp,
                RibbonControlIds.MoveDown,
            ],
            box.Elements(ns + "button").Select(button => button.Attribute("id")?.Value));

        XElement outline = Assert.Single(
            doc.Descendants(ns + "group"),
            g => g.Attribute("id")?.Value == RibbonControlIds.OutlineGroup);
        Assert.Equal(
            ["boxOutlineRow1", "boxOutlineRow2", "boxOutlineRow3"],
            outline.Elements(ns + "box").Select(element => element.Attribute("id")?.Value));
    }

    /// <summary>
    /// The Outline group's Expand/Collapse pair shares the grid's third row.
    /// </summary>
    /// <remarks>
    /// R5.16a replaced the previous one-column vertical stack with the 2 x 3
    /// grid the whole merged group uses. No callback, keytip or enablement
    /// changes: only the layout.
    /// </remarks>
    [Fact]
    public void Outline_expand_and_collapse_share_the_third_row()
    {
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        XElement box = Assert.Single(
            doc.Descendants(ns + "box"),
            b => b.Attribute("id")?.Value == "boxOutlineRow3");
        Assert.Equal("horizontal", box.Attribute("boxStyle")?.Value);
        Assert.Equal(
            [RibbonControlIds.ExpandAllGroups, RibbonControlIds.CollapseAllGroups],
            box.Elements(ns + "button").Select(button => button.Attribute("id")?.Value));
    }

    [Fact]
    public void Move_buttons_label_Up_and_Down_with_correct_actions()
    {
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        Assert.Equal("Up", doc.Descendants(ns + "button")
            .Single(b => b.Attribute("id")?.Value == RibbonControlIds.MoveUp)
            .Attribute("label")?.Value);
        Assert.Equal("OnMoveUpClick", doc.Descendants(ns + "button")
            .Single(b => b.Attribute("id")?.Value == RibbonControlIds.MoveUp)
            .Attribute("onAction")?.Value);

        Assert.Equal("Down", doc.Descendants(ns + "button")
            .Single(b => b.Attribute("id")?.Value == RibbonControlIds.MoveDown)
            .Attribute("label")?.Value);
        Assert.Equal("OnMoveDownClick", doc.Descendants(ns + "button")
            .Single(b => b.Attribute("id")?.Value == RibbonControlIds.MoveDown)
            .Attribute("onAction")?.Value);
    }

    [Fact]
    public void Move_buttons_carry_a_cell_tracking_todo_for_future_selection_gating()
    {
        // R5.4 owner direction: a note to later add cell tracking and disable
        // the ribbon buttons based on cell selection. Pinned so the deferral
        // is not silently lost.
        string xml = new GanttRibbon().GetCustomUI(WorkbookRibbonId)!;
        Assert.Contains("TODO(R5.10)", xml, StringComparison.Ordinal);
        Assert.Contains("cell tracking", xml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("btnMoveUp")]
    [InlineData("btnMoveDown")]
    public void Move_callback_runs_through_boundary_and_refreshes(string controlId)
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
        if (controlId == RibbonControlIds.MoveUp)
        {
            GanttRibbon.OnMoveUpClick(control.Object, boundary, command, () => refreshes++);
        }
        else
        {
            GanttRibbon.OnMoveDownClick(control.Object, boundary, command, () => refreshes++);
        }

        Assert.True(executed);
        Assert.Empty(written);
        Assert.Empty(shown);
        Assert.Equal(1, refreshes);
    }

    [Theory]
    [InlineData("btnMoveUp")]
    [InlineData("btnMoveDown")]
    public void Move_callback_failure_uses_control_id_and_still_refreshes(string controlId)
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
        if (controlId == RibbonControlIds.MoveUp)
        {
            GanttRibbon.OnMoveUpClick(control.Object, boundary, failing, () => refreshes++);
        }
        else
        {
            GanttRibbon.OnMoveDownClick(control.Object, boundary, failing, () => refreshes++);
        }

        Assert.Contains(
            $"command={controlId}",
            Assert.Single(written),
            StringComparison.Ordinal);
        Assert.Equal(1, refreshes);
    }

    [Fact]
    public void Move_controls_stay_disabled_without_a_resolved_selection()
    {
        // R5.9b: the move pair now also requires a row under the active cell, so
        // the boundary no-ops the commands tolerate are no longer offered.
        var adapter = new Mock<IExcelApplicationAdapter>();
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(true);
        var tableReader = new Mock<IGanttTableReader>();
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        var service = new RibbonStateService();
        service.SetApplicationAdapter(adapter.Object);
        service.SetLogAvailabilitySource(() => false);
        service.SetTableReader(tableReader.Object);
        var identifier = new Mock<ISelectionIdentifier>();
        _ = identifier.Setup(i => i.Identify()).Returns(SelectionFacts.None);
        service.SetSelectionIdentifier(identifier.Object);
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.MoveUp));
        Assert.False(service.GetEnabled(RibbonControlIds.MoveDown));

        // Workbook present but the table missing: not initialised, so both grey.
        _ = tableReader.Setup(t => t.Read())
            .Returns(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing));
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.MoveUp));
        Assert.False(service.GetEnabled(RibbonControlIds.MoveDown));

        // No workbook: both grey.
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(false);
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.MoveUp));
        Assert.False(service.GetEnabled(RibbonControlIds.MoveDown));
    }
}
