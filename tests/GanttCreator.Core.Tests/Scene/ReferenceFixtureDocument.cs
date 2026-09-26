// These two records are only ever instantiated by System.Text.Json through
// reflection during fixture deserialization, so CA1812 (never instantiated)
// and CA1515 (make internal) cannot both be satisfied: public fails one,
// internal fails the other. CA1002 is included because System.Text.Json
// needs a settable List<T> to populate, and a read-only collection would
// stop the deserializer binding. The suppression is deliberate and
// file-local rather than a project-wide NoWarn, per tests/Directory.Build.props.
#pragma warning disable CA1812, CA1515, CA1002

namespace GanttCreator.Core.Tests.Scene;

/// <summary>The deserialized root of the committed reference Gantt fixture file.</summary>
public sealed record ReferenceFixtureDocument
{
    /// <summary>Gets the fixture schema version; only 1 is supported.</summary>
    public int SchemaVersion { get; init; }

    /// <summary>Gets the inclusive plot start date declared by the file.</summary>
    public DateOnly PlotStart { get; init; }

    /// <summary>Gets the inclusive plot finish date declared by the file.</summary>
    public DateOnly PlotFinish { get; init; }

    /// <summary>Gets the calendar scale name.</summary>
    public string? Scale { get; init; }

    /// <summary>Gets the period label format name.</summary>
    public string? PeriodLabelFormat { get; init; }

    /// <summary>Gets the event-date display format name.</summary>
    public string? DateFormat { get; init; }

    /// <summary>Gets the chart title.</summary>
    public string? Title { get; init; }

    /// <summary>Gets the human-readable note describing the fixture.</summary>
    public string? Description { get; init; }

    /// <summary>Gets whether alternating plot bands are emitted.</summary>
    public bool AlternateBanding { get; init; }

    /// <summary>Gets whether minor period grid lines are emitted.</summary>
    public bool ShowMinorGrid { get; init; }

    /// <summary>Gets whether major year/plot grid lines are emitted.</summary>
    public bool ShowMajorGrid { get; init; }

    /// <summary>Gets the body rows in file order.</summary>
    public List<ReferenceFixtureRow>? Rows { get; init; }
}

/// <summary>One <c>tblGanttData</c> body row as stored in the fixture file.</summary>
public sealed record ReferenceFixtureRow
{
    /// <summary>Gets the one-based body-row index.</summary>
    public int RowNumber { get; init; }

    /// <summary>Gets the stable row identifier.</summary>
    public string? Id { get; init; }

    /// <summary>Gets the stable visual-lane identifier.</summary>
    public string? LaneId { get; init; }

    /// <summary>Gets the compatibility stack value, or null when the cell is blank.</summary>
    /// <remarks>
    /// Recorded in the fixture because a lane-bound Type requires the cell, but it
    /// is never read for layout: ADR-0012 makes the effective stack Core-derived.
    /// </remarks>
    public int? StackIndex { get; init; }

    /// <summary>Gets the code-owned Type display name.</summary>
    public string? Type { get; init; }

    /// <summary>Gets the row description.</summary>
    public string? Description { get; init; }

    /// <summary>Gets the inclusive start date.</summary>
    public DateOnly? Start { get; init; }

    /// <summary>Gets the inclusive finish date, or null for a point event.</summary>
    public DateOnly? Finish { get; init; }

    /// <summary>Gets the owning activity for a critical interval.</summary>
    public string? ParentId { get; init; }

    /// <summary>Gets the named style key.</summary>
    public string? StyleKey { get; init; }

    /// <summary>Gets the explicit label position.</summary>
    public string? LabelPosition { get; init; }

    /// <summary>Gets the per-row fill override.</summary>
    public string? FillColour { get; init; }

    /// <summary>Gets the per-row stroke override.</summary>
    public string? StrokeColour { get; init; }
}
#pragma warning restore CA1812, CA1515, CA1002
