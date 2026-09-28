using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the source-row to panel-row projection, the explicit owner of the §3
/// membership and ordering rules.
/// </summary>
public sealed class PanelRowProjectionTests
{
    private static readonly GanttRowId _lane = GanttRowId.New();

    private static PanelCellGrid Grid(int columns = 2)
    {
        PanelColumn[] all = [new("Type", 100), new("Description", 200)];
        string[] required = ["Type", "Description"];
        return PanelCellGrid.TryCreate(
            all[..columns],
            [10.0, 10.0, 10.0, 10.0, 10.0, 10.0],
            18.0,
            required[..columns]).Grid!;
    }

    private static GanttEvent Event(
        int rowNumber,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        GanttRowId? id = null,
        GanttRowId? laneId = null,
        bool visible = true,
        string? description = "work") =>
        new(
            rowNumber,
            id ?? GanttRowId.New(),
            laneId,
            null,
            type,
            description,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            visible,
            null);

    private static PanelRowProjectionResult Project(params GanttEvent[] events) =>
        PanelRowProjection.TryProject(events, Grid(), GanttDateDisplayFormat.DdMMyyyy);

    [Fact]
    public void Every_source_row_gets_a_panel_row_whatever_its_type_or_visibility()
    {
        // The case list the remediation required, all in one scene: a Splitter, a
        // Spacer, a Delineator, a Critical Interval, and a hidden row. None of these
        // is an ordinary lane-placed entity, which is exactly why deriving panel rows
        // from lane placements dropped them.
        GanttEvent[] events =
        [
            Event(1, GanttEntityType.Splitter),
            Event(2, GanttEntityType.Spacer),
            Event(3, GanttEntityType.Delineator),
            Event(4, GanttEntityType.CriticalInterval),
            Event(5, visible: false),
            Event(6, GanttEntityType.AsPlannedMilestone),
        ];

        PanelRowProjectionResult result = Project(events);

        Assert.Null(result.Refusal);
        Assert.Equal(6, result.Rows.Count);
        Assert.Equal(
            [events[0].Id, events[1].Id, events[2].Id, events[3].Id, events[4].Id, events[5].Id],
            result.Rows.Select(row => row.RowId));
    }

    [Fact]
    public void Rows_come_out_in_worksheet_order_not_lane_order()
    {
        // Two rows sharing a LaneId, and two stacked events on one lane: the panel
        // must still follow the table. Lane layout orders by SortOrder then row, so
        // the two orders are genuinely different inputs and the panel must not
        // inherit the plot's ordering by accident.
        GanttEvent[] events =
        [
            Event(3, laneId: _lane),
            Event(1, laneId: _lane),
            Event(2, laneId: _lane),
        ];

        PanelRowProjectionResult result = Project(events);

        Assert.Equal(
            [events[1].Id, events[2].Id, events[0].Id],
            result.Rows.Select(row => row.RowId));
    }

    [Fact]
    public void Several_rows_sharing_one_lane_each_get_their_own_panel_row()
    {
        GanttEvent first = Event(1, laneId: _lane);
        GanttEvent second = Event(2, laneId: _lane);
        GanttEvent third = Event(3, laneId: _lane);

        PanelRowProjectionResult result = Project(first, second, third);

        Assert.Equal(3, result.Rows.Count);
        Assert.Equal(3, result.Rows.Select(row => row.RowId).Distinct().Count());
    }

    [Fact]
    public void A_hidden_row_still_reproduces_its_type_and_description_cells()
    {
        // "Visible=false suppresses entity and its label" - it says nothing about the
        // data panel, which reproduces each included cell of the table.
        GanttEvent hidden = Event(1, GanttEntityType.AsBuiltActivity, visible: false, description: "still shown");

        PanelRowProjectionResult result = Project(hidden);

        Assert.Single(result.Rows);
        Assert.Equal("As-Built Activity", result.Rows[0].Cells[0]);
        Assert.Equal("still shown", result.Rows[0].Cells[1]);
    }

    [Fact]
    public void Every_refusal_has_a_positive_case()
    {
        GanttEvent atEvent = Event(1);
        GanttEvent sameId = Event(2, id: atEvent.Id);
        GanttEvent sameRowNumber = Event(1);

        Assert.Equal(
            PanelRowProjectionRefusal.NullEvents,
            PanelRowProjection.TryProject(null, Grid(), GanttDateDisplayFormat.DdMMyyyy).Refusal);
        Assert.Equal(
            PanelRowProjectionRefusal.NullGrid,
            PanelRowProjection.TryProject([atEvent], null, GanttDateDisplayFormat.DdMMyyyy).Refusal);
        Assert.Equal(
            PanelRowProjectionRefusal.UndefinedDateFormat,
            PanelRowProjection.TryProject([atEvent], Grid(), (GanttDateDisplayFormat)99).Refusal);
        // Two rows sharing an identity would emit the same primitive IDs.
        Assert.Equal(
            PanelRowProjectionRefusal.DuplicateRowId,
            Project(atEvent, sameId).Refusal);
        // Two rows claiming one body position would each claim one measured height,
        // so one would silently displace the other.
        Assert.Equal(
            PanelRowProjectionRefusal.DuplicateRowNumber,
            Project(atEvent, sameRowNumber).Refusal);
    }
}
