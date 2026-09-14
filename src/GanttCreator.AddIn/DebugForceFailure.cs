namespace GanttCreator.AddIn;

/// <summary>
/// DEBUG-only force-failure hook. Lets the required Office gate
/// (docs/03-ROADMAP.md R1.4: "forced Excel callback failure shows one safe
/// message and retains usability") demonstrate a forced command failure in a
/// live Excel session without a throwaway spike, and lets later phase exits
/// repeat the demonstration.
/// </summary>
/// <remarks>
/// <para>
/// The environment-variable read lives inside <c>#if DEBUG</c>, so the hook
/// is compiled out of Release builds: <see cref="IsRequested"/> returns
/// <see langword="false"/> unconditionally and no production behaviour or
/// environment probing exists in a Release assembly. The DEBUG-only guarantee
/// is pinned by
/// <c>tests/GanttCreator.Architecture.Tests/DebugHookGuardTests.cs</c> with a
/// positive control.
/// </para>
/// <para>
/// Usage (Visual Studio F5 debug session): set the environment variable to
/// the Ribbon control ID of the command to force, e.g.
/// <c>GANTTCREATOR_FORCE_COMMAND_FAILURE=btnDiagnostics</c>; the next click
/// on that command throws a fixed synthetic exception at the boundary, which
/// produces the one-record/one-dialog failure path.
/// </para>
/// </remarks>
internal static class DebugForceFailure
{
    /// <summary>
    /// The environment variable selecting the command to force. Its value is
    /// compared to the command name case-insensitively; an empty value forces
    /// nothing.
    /// </summary>
    internal const string EnvironmentVariableName = "GANTTCREATOR_FORCE_COMMAND_FAILURE";

    /// <summary>
    /// The fixed synthetic exception message produced by a forced failure.
    /// Internal so tests can assert against the exact message.
    /// </summary>
    internal const string ForcedFailureMessage = "Forced failure (debug force-failure hook).";

    /// <summary>
    /// Throws the fixed synthetic exception when the environment variable
    /// selects <paramref name="commandName"/>. In Release builds this method
    /// is a no-op.
    /// </summary>
    /// <param name="commandName">The command name to compare against the request.</param>
    internal static void ThrowIfRequested(string commandName)
    {
        if (IsRequested(commandName))
        {
            throw new InvalidOperationException(ForcedFailureMessage);
        }
    }

    /// <summary>
    /// Returns whether the environment variable requests a forced failure for
    /// <paramref name="commandName"/>. In Release builds the env-var read is
    /// compiled out and this always returns <see langword="false"/>.
    /// </summary>
    /// <param name="commandName">The command name to compare against the request.</param>
    /// <returns>True when a forced failure is requested for the command.</returns>
    internal static bool IsRequested(string commandName)
    {
#if DEBUG
        if (string.IsNullOrEmpty(commandName))
        {
            return false;
        }

        var requested = Environment.GetEnvironmentVariable(EnvironmentVariableName);
        return !string.IsNullOrWhiteSpace(requested)
            && string.Equals(requested, commandName, StringComparison.OrdinalIgnoreCase);
#else
        _ = commandName;
        return false;
#endif
    }
}
