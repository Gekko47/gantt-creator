namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the DEBUG-only force-failure hook (R1.4). The
/// Debug-only facts are wrapped in <c>#if DEBUG</c> because the hook's
/// env-var read is compiled out of Release builds; the always-on fact
/// asserts the Release contract (never requested) in both configurations.
/// </summary>
public class DebugForceFailureTests
{
    [Fact]
    public void IsRequested_returns_false_when_the_environment_variable_is_unset()
    {
        Environment.SetEnvironmentVariable(DebugForceFailure.EnvironmentVariableName, null);
        try
        {
            Assert.False(DebugForceFailure.IsRequested("btnDiagnostics"));
        }
        finally
        {
            Environment.SetEnvironmentVariable(DebugForceFailure.EnvironmentVariableName, null);
        }
    }

#if DEBUG
    private const string RequestedCommand = "btnForceFailure";

    private static void SetRequest(string? value)
        => Environment.SetEnvironmentVariable(DebugForceFailure.EnvironmentVariableName, value);

    [Fact]
    public void IsRequested_returns_true_for_the_matching_command_name()
    {
        SetRequest(RequestedCommand);
        try
        {
            Assert.True(DebugForceFailure.IsRequested(RequestedCommand));
            Assert.True(DebugForceFailure.IsRequested(RequestedCommand.ToUpperInvariant()));
        }
        finally
        {
            SetRequest(null);
        }
    }

    [Fact]
    public void IsRequested_returns_false_for_a_different_command_name()
    {
        SetRequest(RequestedCommand);
        try
        {
            Assert.False(DebugForceFailure.IsRequested("btnOther"));
        }
        finally
        {
            SetRequest(null);
        }
    }

    [Fact]
    public void IsRequested_returns_false_for_an_empty_value()
    {
        SetRequest(string.Empty);
        try
        {
            Assert.False(DebugForceFailure.IsRequested(RequestedCommand));
        }
        finally
        {
            SetRequest(null);
        }
    }

    [Fact]
    public void ThrowIfRequested_throws_the_fixed_message_for_the_requested_command()
    {
        SetRequest(RequestedCommand);
        try
        {
            var thrown = Assert.Throws<InvalidOperationException>(
                () => DebugForceFailure.ThrowIfRequested(RequestedCommand));
            Assert.Equal(DebugForceFailure.ForcedFailureMessage, thrown.Message);
        }
        finally
        {
            SetRequest(null);
        }
    }

    [Fact]
    public void ThrowIfRequested_does_nothing_for_a_different_command_name()
    {
        SetRequest(RequestedCommand);
        try
        {
            DebugForceFailure.ThrowIfRequested("btnOther");
        }
        finally
        {
            SetRequest(null);
        }
    }
#endif
}
