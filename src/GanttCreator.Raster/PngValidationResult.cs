namespace GanttCreator.Raster;

/// <summary>The outcome of re-reading and validating a written PNG file.</summary>
/// <param name="IsValid">Whether every pinned fact matched.</param>
/// <param name="Failure">The first failing fact when <paramref name="IsValid"/> is false; empty otherwise.</param>
public sealed record PngValidationResult(bool IsValid, string Failure);
