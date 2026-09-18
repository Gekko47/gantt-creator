namespace GanttCreator.Core;

/// <summary>
/// Which date columns a row type reads for geometry. Drives the invariant
/// that milestones and delineators read <c>Start</c> as their single date and
/// never use <c>Finish</c> for positioning.
/// </summary>
public enum EntityDateMode
{
    /// <summary>No date columns are read (Splitter, Spacer).</summary>
    None = 0,

    /// <summary>Reads <c>Start</c> and <c>Finish</c> (span events).</summary>
    StartFinish = 1,

    /// <summary>Reads only <c>Start</c>; <c>Finish</c> is never used for geometry (milestones, delineators).</summary>
    StartOnly = 2,
}
