using System.Globalization;

namespace GanttCreator.Core.Scene;

/// <summary>The reason a panel row projection was refused.</summary>
public enum PanelRowProjectionRefusal
{
    /// <summary>The event collection was null.</summary>
    NullEvents = 0,

    /// <summary>The measured grid was null.</summary>
    NullGrid = 1,

    /// <summary>The approved event-date display format was not a defined value.</summary>
    UndefinedDateFormat = 2,

    /// <summary>Two source rows shared a row identity, which would collide primitive IDs.</summary>
    DuplicateRowId = 3,

    /// <summary>Two source rows claimed the same body-row number.</summary>
    DuplicateRowNumber = 4,
}

/// <summary>The typed result of projecting source rows onto panel rows.</summary>
/// <param name="Rows">The projected panel rows, in worksheet order.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record PanelRowProjectionResult(
    IReadOnlyList<PanelRow> Rows,
    PanelRowProjectionRefusal? Refusal)
{
    /// <summary>Gets whether the projection succeeded.</summary>
    public bool Succeeded => Refusal is null;
}

/// <summary>
/// Projects the validated source rows onto the section 3 data-panel rows, so the
/// mapping is one named, tested Core function rather than a property of whichever
/// collection the orchestrator happened to have to hand.
/// </summary>
/// <remarks>
/// <para>
/// <b>Every source row gets a panel row.</b> §3's source is
/// "<c>tblGanttData</c>, visible approved columns, current Excel column widths, and
/// row heights" — "visible" qualifies the <em>columns</em>, not the rows — and the
/// guide requires the panel to reproduce each included cell. The field table is
/// equally explicit that <c>Visible</c> "suppresses entity and its label", which says
/// nothing about the panel. So a <c>Splitter</c>, a <c>Spacer</c>, a
/// <c>Delineator</c>, a <c>Critical Interval</c>, and a hidden row all appear in the
/// panel, because all of them are rows in the table the panel reproduces.
/// </para>
/// <para>
/// <b>Order is worksheet order, not lane order.</b> Rows come out in
/// <see cref="GanttEvent.RowNumber"/> order because the panel restates the table.
/// <see cref="LaneOrdering"/> deliberately sorts by <c>SortOrder</c> and then row
/// number for the <em>plot</em>, and the two orders differ once a user supplies
/// <c>SortOrder</c>. Sorting the panel by lane order would make the data panel
/// disagree with the worksheet beside it.
/// </para>
/// <para>
/// This is a pure function: it measures nothing, mutates nothing, and reads no
/// Office state. Dates go through <see cref="GanttDateFormatting"/> with the
/// request's approved format so a panel cell and a §23 date label cannot disagree and
/// no host culture can re-derive the pattern.
/// </para>
/// </remarks>
public static class PanelRowProjection
{
    /// <summary>Projects the source rows onto panel rows in worksheet order.</summary>
    /// <param name="events">The validated source rows, in any order.</param>
    /// <param name="grid">The measured grid, whose columns fix the cell order.</param>
    /// <param name="dateFormat">The approved event-date display format.</param>
    /// <returns>The projected rows, or a typed refusal.</returns>
    public static PanelRowProjectionResult TryProject(
        IReadOnlyList<GanttEvent>? events,
        PanelCellGrid? grid,
        GanttDateDisplayFormat dateFormat)
    {
        if (events is null)
        {
            return new PanelRowProjectionResult([], PanelRowProjectionRefusal.NullEvents);
        }

        if (grid is null)
        {
            return new PanelRowProjectionResult([], PanelRowProjectionRefusal.NullGrid);
        }

        if (!Enum.IsDefined(dateFormat))
        {
            return new PanelRowProjectionResult([], PanelRowProjectionRefusal.UndefinedDateFormat);
        }

        HashSet<GanttRowId> ids = [];
        HashSet<int> rowNumbers = [];
        foreach (GanttEvent atEvent in events)
        {
            if (!ids.Add(atEvent.Id))
            {
                return new PanelRowProjectionResult([], PanelRowProjectionRefusal.DuplicateRowId);
            }

            // Two rows claiming one body position would share a measured height and
            // silently displace each other, so it is refused rather than resolved by
            // an arbitrary tie-break.
            if (!rowNumbers.Add(atEvent.RowNumber))
            {
                return new PanelRowProjectionResult([], PanelRowProjectionRefusal.DuplicateRowNumber);
            }
        }

        // RowNumber is the one-based body-row index, so ascending is table order.
        // The stable-ID tie-break is unreachable given the duplicate check above, but
        // keeps the ordering total if that check is ever relaxed.
        List<PanelRow> rows =
        [
            .. events
                .OrderBy(atEvent => atEvent.RowNumber)
                .ThenBy(atEvent => atEvent.Id.Value, StringComparer.Ordinal)
                .Select(atEvent => new PanelRow(atEvent.Id, Cells(atEvent, grid, dateFormat))),
        ];

        return new PanelRowProjectionResult(rows, null);
    }

    /// <summary>Projects one event's cell texts in grid-column order.</summary>
    /// <param name="atEvent">The validated event supplying the cell values.</param>
    /// <param name="grid">The measured grid, whose columns fix the cell order.</param>
    /// <param name="dateFormat">The approved event-date display format.</param>
    /// <returns>
    /// One entry per grid column. A column with no schema mapping, or a field the
    /// entity type does not use, is <see langword="null"/>: blank is legal cell
    /// data (R2.5 U2) and the builder already omits text for it.
    /// </returns>
    private static List<string?> Cells(GanttEvent atEvent, PanelCellGrid grid, GanttDateDisplayFormat dateFormat)
    {
        List<string?> cells = new(grid.Columns.Count);
        foreach (PanelColumn column in grid.Columns)
        {
            cells.Add(column.Name switch
            {
                // Every schema column is mapped explicitly rather than relying on a
                // default: a panel cell must reproduce the worksheet value it stands
                // for, and a silently null column renders as a blank cell that reads
                // as an empty worksheet cell.
                "Id" => atEvent.Id.Value,
                "Type" => EntityTypeCatalog.GetDefinition(atEvent.Type)?.DisplayName,
                "Description" => atEvent.Description,
                "Start" => atEvent.Start is { } start ? GanttDateFormatting.Format(start, dateFormat) : null,
                "Finish" => atEvent.Finish is { } finish ? GanttDateFormatting.Format(finish, dateFormat) : null,
                "LaneId" => atEvent.LaneId?.Value,
                "StackIndex" => atEvent.StackIndex?.ToString(CultureInfo.InvariantCulture),
                "ParentId" => atEvent.ParentId?.Value,
                "StyleKey" => atEvent.StyleKey,
                // The worksheet holds the enum member name, which is also what
                // GanttRowValidator parses back, so the cell round-trips.
                "LabelPosition" => atEvent.LabelPosition?.ToString(),
                "FillColour" => atEvent.FillColour,
                "StrokeColour" => atEvent.StrokeColour,
                // ExcelCellConverter.ToText renders a bool as TRUE/FALSE, so the cell
                // matches the value the table reader would parse back.
                "Visible" => atEvent.Visible ? "TRUE" : "FALSE",
                "SortOrder" => atEvent.SortOrder?.ToString(CultureInfo.InvariantCulture),
                _ => null,
            });
        }

        return cells;
    }
}

