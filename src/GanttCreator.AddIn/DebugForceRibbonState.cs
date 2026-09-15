namespace GanttCreator.AddIn;

/// <summary>
/// DEBUG-only force-state hook. Lets the required Office gate
/// (docs/03-ROADMAP.md R1.5: "controls enable/disable correctly as workbook
/// state changes") demonstrate both state transitions in a live Excel session
/// even when Excel does not expose the custom tab with no workbook open.
/// </summary>
/// <remarks>
/// <para>
/// The environment-variable read lives inside <c>#if DEBUG</c>, so the hook is
/// compiled out of Release builds: <see cref="Apply"/> returns its input
/// unchanged and no environment probing exists in a Release assembly. The
/// DEBUG-only guarantee is pinned by
/// <c>tests/GanttCreator.Architecture.Tests/DebugHookGuardTests.cs</c> (the
/// repo-wide scan) with a positive control, matching the R1.4 hook.
/// </para>
/// <para>
/// Usage (Visual Studio F5 debug session): set the environment variable to
/// <c>no-workbook</c> or <c>no-log</c>; the next refresh reports that fact as
/// false, so the gated control disables without a real workbook/log change.
/// </para>
/// </remarks>
internal static class DebugForceRibbonState
{
    /// <summary>The environment variable selecting the forced fact.</summary>
    internal const string EnvironmentVariableName = "GANTTCREATOR_FORCE_RIBBON_STATE";

    /// <summary>The value that forces <see cref="RibbonState.HasActiveWorkbook"/> to false.</summary>
    internal const string NoWorkbookValue = "no-workbook";

    /// <summary>The value that forces <see cref="RibbonState.LogAvailable"/> to false.</summary>
    internal const string NoLogValue = "no-log";

    /// <summary>
    /// Applies the requested forced fact to <paramref name="state"/>. In Release
    /// builds this returns <paramref name="state"/> unchanged.
    /// </summary>
    /// <param name="state">The freshly captured snapshot.</param>
    /// <returns>The snapshot the service should publish.</returns>
    internal static RibbonState Apply(RibbonState state)
    {
#if DEBUG
        var requested = Environment.GetEnvironmentVariable(EnvironmentVariableName);
        if (string.IsNullOrWhiteSpace(requested))
        {
            return state;
        }

        if (string.Equals(requested, NoWorkbookValue, StringComparison.OrdinalIgnoreCase))
        {
            return state with { HasActiveWorkbook = false };
        }
        else if (string.Equals(requested, NoLogValue, StringComparison.OrdinalIgnoreCase))
        {
            return state with { LogAvailable = false };
        }

        return state;
#else
        return state;
#endif
    }
}