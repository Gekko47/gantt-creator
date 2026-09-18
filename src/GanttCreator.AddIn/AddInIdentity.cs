namespace GanttCreator.AddIn;

/// <summary>
/// Identifies one add-in load session for the lifecycle log: version,
/// Excel version, process bitness, XLL file name, and a per-load session
/// token that correlates the <c>open</c> and <c>close</c> records of the
/// same load. Never carries workbook content — the logging content rule in
/// <c>docs/01-ENVIRONMENT.md</c> permits identifiers and counts only.
/// </summary>
/// <param name="AddInVersion">The product version string. Both the <c>open</c> and <c>Diagnostics:</c> records carry Core's <see cref="Core.VersionInfo"/> informational version unchanged, including build metadata.</param>
/// <param name="ExcelVersion">Excel major.minor formatted with the invariant culture, or <c>"unknown"</c>.</param>
/// <param name="ProcessBitness"><c>"x64"</c> or <c>"x86"</c>, or <c>"unknown"</c>.</param>
/// <param name="XllFileName">File name (never a full path) of the loaded XLL, or <c>"unknown"</c>.</param>
/// <param name="SessionToken">
/// Per-load correlation token shared by the <c>open</c> and <c>close</c>
/// records of one load (work item R1.6 D5). Deliberately not a GUID or long
/// hex string: Core's <see cref="Core.Logging.Redactor"/> masks both to
/// <c>[guid]</c>/<c>[token]</c>, which would destroy correlation on disk.
/// </param>
public sealed record AddInIdentity(
    string AddInVersion,
    string ExcelVersion,
    string ProcessBitness,
    string XllFileName,
    string SessionToken);
