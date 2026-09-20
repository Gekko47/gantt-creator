namespace GanttCreator.Core;

/// <summary>
/// The result of <see cref="GanttRowValidator.Validate"/>: the valid events
/// in input table order plus every blocking error and non-blocking warning in
/// deterministic <c>Severity → Row → Field → Code</c> order. Immutable.
/// </summary>
/// <param name="Events">The validated events; rows with blocking errors are excluded.</param>
/// <param name="Issues">Every finding, sorted by severity, row, field, then code.</param>
/// <param name="IsValid"><see langword="true"/> when no issue has <see cref="GanttValidationSeverity.Error"/> severity.</param>
public sealed record GanttValidationOutcome(
    IReadOnlyList<GanttEvent> Events,
    IReadOnlyList<GanttValidationIssue> Issues,
    bool IsValid);
