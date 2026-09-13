namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Pins the DEBUG-only force-failure hook contract (R1.4): the environment
/// variable read in <c>DebugForceFailure</c> must live inside an
/// <c>#if DEBUG</c> guard so no production behaviour or environment probing
/// exists in Release builds. The predicate is checked with a positive
/// control so the guard itself cannot fail vacuously.
/// </summary>
public sealed class DebugHookGuardTests
{
    private const string HookRelativePath = "src/GanttCreator.AddIn/DebugForceFailure.cs";

    [Fact]
    public void DebugForceFailure_hook_keeps_the_env_var_read_inside_ifdef_debug()
    {
        var content = File.ReadAllText(LocateRepoFile(HookRelativePath));

        Assert.True(
            EnvVarReadIsInsideIfdefDebug(content),
            $"The env-var read in {HookRelativePath} must be guarded by `#if DEBUG` " +
            "(DebugForceFailure must be compiled out of Release builds; see " +
            "docs/work-items/R1.4-command-error-boundary.md).");
    }

    [Fact]
    public void Guard_accepts_a_hook_with_ifdef_debug_positive_control()
    {
        const string guarded = """
            internal static class GoodHook
            {
                internal static bool IsRequested(string commandName)
                {
            #if DEBUG
                    var requested = Environment.GetEnvironmentVariable("GANTTCREATOR_FORCE_COMMAND_FAILURE");
                    return requested == commandName;
            #else
                    return false;
            #endif
                }
            }
            """;

        Assert.True(EnvVarReadIsInsideIfdefDebug(guarded));
    }

    [Fact]
    public void Guard_rejects_a_hook_without_ifdef_debug_positive_control()
    {
        const string unguarded = """
            internal static class BadHook
            {
                internal static bool IsRequested(string commandName)
                {
                    var requested = Environment.GetEnvironmentVariable("GANTTCREATOR_FORCE_COMMAND_FAILURE");
                    return requested == commandName;
                }
            }
            """;

        Assert.False(EnvVarReadIsInsideIfdefDebug(unguarded));
    }

    /// <summary>
    /// Returns true when the first <c>Environment.GetEnvironmentVariable</c>
    /// call appears between an opening <c>#if DEBUG</c> and its closing
    /// <c>#endif</c>.
    /// </summary>
    private static bool EnvVarReadIsInsideIfdefDebug(string content)
    {
        var readIndex = content.IndexOf("Environment.GetEnvironmentVariable", StringComparison.Ordinal);
        if (readIndex < 0)
        {
            return false;
        }

        var ifdefIndex = content.LastIndexOf("#if DEBUG", readIndex, StringComparison.Ordinal);
        if (ifdefIndex < 0)
        {
            return false;
        }

        var endifIndex = content.IndexOf("#endif", ifdefIndex, StringComparison.Ordinal);
        return endifIndex > readIndex;
    }

    /// <summary>
    /// Locates a repo file by walking up from the test binary until the
    /// requested relative path exists (same pattern as
    /// <c>ArtifactSourceMarkerTests.LocateTestsRoot</c>).
    /// </summary>
    private static string LocateRepoFile(string relativePath)
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(DebugHookGuardTests).Assembly.Location)!);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, relativePath.Replace('/', Path.DirectorySeparatorChar));
            if (File.Exists(candidate))
            {
                return candidate;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException(
            $"Could not locate '{relativePath}' from the test binary.");
    }
}
