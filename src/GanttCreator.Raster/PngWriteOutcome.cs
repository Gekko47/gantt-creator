namespace GanttCreator.Raster;

/// <summary>The result of one atomic PNG write.</summary>
/// <param name="Succeeded">Whether the file was written and verified.</param>
/// <param name="Refusal">The failure kind when <paramref name="Succeeded"/> is false; null on success.</param>
/// <param name="ActionableMessage">The user-facing message; non-null and non-empty in both outcomes.</param>
/// <param name="TechnicalRecord">The redaction-ready technical record; non-null and non-empty in both outcomes.</param>
/// <remarks>
/// On success <see cref="Refusal"/> is null; on failure all three other members
/// are populated. The command boundary owns operation-ID stamping, logging, and
/// message presentation.
/// </remarks>
public sealed record PngWriteOutcome(
    bool Succeeded,
    PngWriteRefusal? Refusal,
    string ActionableMessage,
    string TechnicalRecord);
