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

        // R5.4 restructured the group into a vertical box so the move pair
        // can sit as its third row; the promote/demote pair is still exactly
        // this group's content, now the first two rows of that box.
        XElement ordering = Assert.Single(group.Elements(ns + "box"));
        XElement[] buttons = ordering.Elements(ns + "button").ToArray();

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

    [Fact]
    public void Move_buttons_sit_in_a_two_column_box_inside_the_ordering_group()
    {
        // R5.4 ribbon restructure (owner direction): the Ordering group is a
        // vertical box whose third row is a horizontal box of exactly two
        // columns — column 1 Up, column 2 Down. The move buttons are NOT
        // direct children of the ordering box; they live in boxMoveRows, so
        // the vertical stack renders promote/demote above the up/down pair.
        XDocument doc = XDocument.Parse(new GanttRibbon().GetCustomUI(WorkbookRibbonId)!);
        XNamespace ns = CustomUiNamespace;

        XElement box = Assert.Single(
            doc.Descendants(ns + "box"),
            b => b.Attribute("id")?.Value == "boxMoveRows");
        Assert.Equal("horizontal", box.Attribute("boxStyle")?.Value);
        Assert.Equal(
            [
                RibbonControlIds.MoveUp,
                RibbonControlIds.MoveDown,
            ],
            box.Elements(ns + "button").Select(button => button.Attribute("id")?.Value));

        // The move box is the third element of the vertical ordering box,
        // after promote and demote.
        XElement ordering = Assert.Single(
            doc.Descendants(ns + "box"),
            b => b.Attribute("id")?.Value == "boxOrdering");
        Assert.Equal("vertical", ordering.Attribute("boxStyle")?.Value);
        Assert.Equal(
            [RibbonControlIds.PromoteToParent, RibbonControlIds.DemoteToChild, "boxMoveRows"],
            ordering.Elements()
                .Where(element => element.Name == ns + "button" || element.Name == ns + "box")
                .Select(element => element.Attribute("id")?.Value));
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
    public void Move_controls_follow_the_cached_workbook_and_initialised_state()
    {
        // The move buttons gate on HasActiveWorkbook && SheetInitialised, exactly
        // like promote and demote: the top/bottom of the selection's set is a
        // quiet no-op by design (R5.4 D4), not a greyed button.
        var adapter = new Mock<IExcelApplicationAdapter>();
        _ = adapter.Setup(a => a.HasActiveWorkbook()).Returns(true);
        var tableReader = new Mock<IGanttTableReader>();
        _ = tableReader.Setup(t => t.Read()).Returns(GanttTableReadOutcome.Ok([]));
        var service = new RibbonStateService();
        service.SetApplicationAdapter(adapter.Object);
        service.SetLogAvailabilitySource(() => false);
        service.SetTableReader(tableReader.Object);
        service.Refresh();

        Assert.True(service.GetEnabled(RibbonControlIds.MoveUp));
        Assert.True(service.GetEnabled(RibbonControlIds.MoveDown));

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
