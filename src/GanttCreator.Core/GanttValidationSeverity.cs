namespace GanttCreator.Core;

/// <summary>
/// The severity of one <see cref="GanttValidationIssue"/>.
/// Blocking <see cref="Error"/>s exclude the row from <see cref="GanttValidationOutcome.Events"/>;
/// non-blocking <see cref="Warning"/>s permit rendering.
/// </summary>
public enum GanttValidationSeverity
{
    /// <summary>A blocking error; the row produces no event.</summary>
    Error = 0,

    /// <summary>A non-blocking warning; the row may still produce an event.</summary>
    Warning = 1,
}
