using GanttCreator.Core.Logging;
using Xunit;

// CA2000: The latched-failure log in the unavailable-directory test is
// deliberately left undisposed — the test asserts the never-throw contract
// of a log whose directory cannot exist, and a failed log holds no writer
// (RollingLog releases it on latch). The other tests dispose explicitly.
// Scoped to this file; production code still enforces CA2000.
#pragma warning disable CA2000

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Log-location contract tests for <see cref="AddInLogFactory"/>: the
/// product log-tree layout, the temp fallback, and the never-throw
/// behaviour when the target directory is unavailable.
/// </summary>
public class AddInLogFactoryTests
{
    [Fact]
    public void Create_uses_the_given_base_directory_and_writes_a_log_file()
    {
        var baseDirectory = Directory.CreateTempSubdirectory("gantt-r11-").FullName;
        try
        {
            IRollingLog log = AddInLogFactory.Create(baseDirectory);

            log.Write("factory probe");
            log.Dispose();

            var expected = Path.Combine(baseDirectory, "GanttCreator", "logs", "gantt-creator-addin.log");
            Assert.True(File.Exists(expected), $"Expected the log file at '{expected}'.");
            Assert.Contains("factory probe", File.ReadAllText(expected), StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(baseDirectory, recursive: true);
        }
    }

    [Fact]
    public void Create_reports_failure_instead_of_throwing_when_directory_unavailable()
    {
        // Pointing the base directory at an existing *file* makes
        // Directory.CreateDirectory fail inside RollingLog; the contract is
        // a latched failure, never a thrown exception.
        var probeFile = Path.Combine(Path.GetTempPath(), $"gantt-r11-{Guid.NewGuid():N}.tmp");
        File.WriteAllText(probeFile, "blocked");
        try
        {
            IRollingLog log = AddInLogFactory.Create(probeFile);

            Assert.True(log.IsFailed, "An unavailable log directory must latch IsFailed.");
            var exception = Record.Exception(() => log.Write("must not throw"));
            Assert.Null(exception);
        }
        finally
        {
            File.Delete(probeFile);
        }
    }
}

/// <summary>
/// Temp-fallback log tests that share the same log file path
/// under Path.GetTempPath(). They must run sequentially to avoid
/// concurrent file-write collisions.
/// </summary>
#pragma warning disable CA1515, CA1711
[CollectionDefinition("temp-log-fallback")]
public class TempLogFallbackTestsCollection
{
}
#pragma warning restore CA1515, CA1711

[Collection("temp-log-fallback")]
public class TempLogFallbackTests
{
    [Fact]
    public void Create_falls_back_to_temp_when_base_directory_is_missing()
    {
        // Passing null selects the temp fallback inside AddInLogFactory.
        var expectedLogFile = Path.Combine(
            Path.GetTempPath(), "GanttCreator", "logs", "gantt-creator-addin.log");
        var expectedLogDir = Path.GetDirectoryName(expectedLogFile)!;
        if (File.Exists(expectedLogFile)) File.Delete(expectedLogFile);
        if (Directory.Exists(expectedLogDir)) Directory.Delete(expectedLogDir, recursive: true);
        try
        {
            IRollingLog log = AddInLogFactory.Create(null);

            log.Write("missing-dir probe");
            log.Dispose();

            Assert.True(File.Exists(expectedLogFile), $"Expected the log at '{expectedLogFile}'.");
            Assert.True(expectedLogFile.StartsWith(Path.GetTempPath(), StringComparison.Ordinal), "The log must be created beneath the temp path.");
            Assert.Contains("missing-dir probe", File.ReadAllText(expectedLogFile), StringComparison.Ordinal);
        }
        finally
        {
            if (File.Exists(expectedLogFile)) File.Delete(expectedLogFile);
            if (Directory.Exists(expectedLogDir)) Directory.Delete(expectedLogDir, recursive: true);
        }
    }

    [Fact]
    public void Create_falls_back_to_temp_when_base_directory_is_whitespace()
    {
        // Passing whitespace selects the temp fallback inside AddInLogFactory.
        var expectedLogFile = Path.Combine(
            Path.GetTempPath(), "GanttCreator", "logs", "gantt-creator-addin.log");
        var expectedLogDir = Path.GetDirectoryName(expectedLogFile)!;
        if (File.Exists(expectedLogFile)) File.Delete(expectedLogFile);
        if (Directory.Exists(expectedLogDir)) Directory.Delete(expectedLogDir, recursive: true);
        try
        {
            IRollingLog log = AddInLogFactory.Create("   ");

            log.Write("whitespace-dir probe");
            log.Dispose();

            Assert.True(File.Exists(expectedLogFile), $"Expected the log at '{expectedLogFile}'.");
            Assert.True(expectedLogFile.StartsWith(Path.GetTempPath(), StringComparison.Ordinal), "The log must be created beneath the temp path.");
            Assert.Contains("whitespace-dir probe", File.ReadAllText(expectedLogFile), StringComparison.Ordinal);
        }
        finally
        {
            if (File.Exists(expectedLogFile)) File.Delete(expectedLogFile);
            if (Directory.Exists(expectedLogDir)) Directory.Delete(expectedLogDir, recursive: true);
        }
    }
}
