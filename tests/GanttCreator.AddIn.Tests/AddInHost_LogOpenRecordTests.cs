using GanttCreator.AddIn;
using GanttCreator.Core.Logging;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Positive tests for <see cref="AddInHost.LogOpenRecord"/>: a transient write
/// hiccup is retried once, a persistent failure still propagates (never hidden).
/// </summary>
public class AddInHost_LogOpenRecordTests
{
    [Fact]
    public void LogOpenRecord_retries_once_then_propagates_a_persistent_fail()
    {
        // Arrange: a log that fails every write (IsFailed latched true), so
        // the retry fires once and the persistent failure must propagate.
        using (var log = new FailingRollingLog())
        {
            var identity = new AddInIdentity(
                "0.0.0+local.test", "unknown", "x64",
                "GanttCreator.AddIn-AddIn64-packed.xll", "s1-g0h1i2j3k4l5");

            // Act & Assert
            Assert.ThrowsAny<Exception>(() => AddInHost.LogOpenRecord(log, identity));
            Assert.Equal(2, log.WriteCount);
            Assert.True(log.IsFailed);
        }
    }

    [Fact]
    public void LogOpenRecord_retries_once_then_succeeds_after_a_transient_hiccup()
    {
        // Arrange: a log that fails its first write and succeeds on the retry.
        using (var log = new TransientlyFailingRollingLog())
        {
            // Act
            AddInHost.LogOpenRecord(log, new AddInIdentity(
                "0.0.0+local.test", "unknown", "x64",
                "GanttCreator.AddIn-AddIn64-packed.xll", "s1-g0h1i2j3k4l5"));

            // Assert
            Assert.Equal(2, log.WriteCount);
            Assert.False(log.IsFailed);
        }
    }

    private sealed class FailingRollingLog : IRollingLog
    {
        public int WriteCount { get; private set; }
        public bool IsFailed => true;
        public string? LogFilePath { get; } = "gantt-creator-addin.log";

        public void Write(string message)
        {
            WriteCount++;
            throw new InvalidOperationException("log write failed persistently");
        }

        public void Write(string format, params object?[] args) => Write(format);

        public void Dispose() { }
    }

    private sealed class TransientlyFailingRollingLog : IRollingLog
    {
        public int WriteCount { get; private set; }
        public bool IsFailed => false;
        public string? LogFilePath { get; } = "gantt-creator-addin.log";

        public void Write(string message)
        {
            WriteCount++;
            if (WriteCount == 1)
            {
                throw new InvalidOperationException("log write failed transiently");
            }
        }

        public void Write(string format, params object?[] args) => Write(format);

        public void Dispose() { }
    }
}
