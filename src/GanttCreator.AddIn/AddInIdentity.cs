namespace GanttCreator.AddIn;

/// <summary>
/// Identifies one add-in load session for the lifecycle log: version,
/// Excel version, process bitness, and XLL file name only. Never carries
/// workbook content — the logging content rule in
/// <c>docs/01-ENVIRONMENT.md</c> permits identifiers and counts only.
/// </summary>
/// <param name="AddInVersion">The product version string, from Core's <see cref="Core.VersionInfo"/> informational version.</param>
/// <param name="ExcelVersion">Excel major.minor formatted with the invariant culture, or <c>"unknown"</c>.</param>
/// <param name="ProcessBitness"><c>"x64"</c> or <c>"x86"</c>, or <c>"unknown"</c>.</param>
/// <param name="XllFileName">File name (never a full path) of the loaded XLL, or <c>"unknown"</c>.</param>
public sealed record AddInIdentity(
    string AddInVersion,
    string ExcelVersion,
    string ProcessBitness,
    string XllFileName);
