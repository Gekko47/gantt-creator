namespace GanttCreator.Core;

/// <summary>
/// The visual-entity class a row type renders as. Prevents the scene engine
/// and renderers from re-deriving the shape class from the date mode.
/// </summary>
public enum EntityKind
{
    /// <summary>Section header band (Splitter).</summary>
    SectionHeader = 0,

    /// <summary>Blank vertical space; no visible entity (Spacer).</summary>
    Spacer = 1,

    /// <summary>A span between two dates (activities, procurements, delay events, critical intervals).</summary>
    Span = 2,

    /// <summary>A point marker on a single date (milestones).</summary>
    Milestone = 3,

    /// <summary>A full-height vertical date line (Delineator).</summary>
    Delineator = 4,
}
