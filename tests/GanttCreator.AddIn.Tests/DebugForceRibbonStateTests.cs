using GanttCreator.AddIn;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the R1.5 DEBUG-only force-ribbon-state hook.
/// The hook's env-var read lives inside <c>#if DEBUG</c>, so
/// <see cref="DebugForceRibbonState.Apply"/> returns its input unchanged in
/// Release builds; the always-on test asserts the Release contract (no mutation)
/// in both configurations.
/// </summary>
public class DebugForceRibbonStateTests
{
    [Fact]
    public void Apply_returns_the_snapshot_unchanged_when_the_environment_variable_is_unset()
    {
        var prior = Environment.GetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName);
        Environment.SetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName, null);
        try
        {
            var state = new RibbonState(true, true);
            var result = DebugForceRibbonState.Apply(state);
            Assert.Equal(state, result);
            Assert.True(result.HasActiveWorkbook);
            Assert.True(result.LogAvailable);
        }
        finally
        {
            Environment.SetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName, prior);
        }
    }

#if DEBUG
    private const string ForceWorkbookValue = "no-workbook";
    private const string ForceLogValue = "no-log";

    private static void SetVar(string? value)
        => Environment.SetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName, value);

    [Fact]
    public void Apply_forces_HasActiveWorkbook_to_false_for_the_workbook_value()
    {
        var prior = Environment.GetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName);
        SetVar(ForceWorkbookValue);
        try
        {
            var state = new RibbonState(true, true);
            var result = DebugForceRibbonState.Apply(state);
            Assert.False(result.HasActiveWorkbook);
            Assert.True(result.LogAvailable);
        }
        finally
        {
            Environment.SetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName, prior);
        }
    }

    [Fact]
    public void Apply_forces_LogAvailable_to_false_for_the_log_value()
    {
        var prior = Environment.GetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName);
        SetVar(ForceLogValue);
        try
        {
            var state = new RibbonState(true, true);
            var result = DebugForceRibbonState.Apply(state);
            Assert.True(result.HasActiveWorkbook);
            Assert.False(result.LogAvailable);
        }
        finally
        {
            Environment.SetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName, prior);
        }
    }

    [Fact]
    public void Apply_returns_the_snapshot_unchanged_for_an_unrecognised_value()
    {
        var prior = Environment.GetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName);
        SetVar("unknown-value");
        try
        {
            var state = new RibbonState(false, false);
            var result = DebugForceRibbonState.Apply(state);
            Assert.Equal(state, result);
        }
        finally
        {
            Environment.SetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName, prior);
        }
    }

    [Fact]
    public void Apply_is_case_insensitive_for_the_known_values()
    {
        var prior = Environment.GetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName);

        SetVar(ForceWorkbookValue.ToUpperInvariant());
        try
        {
            Assert.False(DebugForceRibbonState.Apply(new RibbonState(true, true)).HasActiveWorkbook);
        }
        finally
        {
            Environment.SetEnvironmentVariable(DebugForceRibbonState.EnvironmentVariableName, prior);
        }
    }
#endif
}
