using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Enforces that test-only diagnostic suppressions remain scoped to test
/// projects. CA1707 (identifiers should have correct suffix) must be
/// suppressed in <c>tests/Directory.Build.props</c> only, never in the root
/// <c>Directory.Build.props</c>, so production code still enforces the rule.
/// </summary>
public sealed class NoWarnScopeTests
{
    private static string[] ReadNoWarnValues(string relativePath)
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(NoWarnScopeTests).Assembly.Location)!);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, relativePath);
            if (File.Exists(candidate) && dir.Name != "tests")
            {
                var text = File.ReadAllText(candidate);
                var doc = XDocument.Parse(text);
                var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;
                return doc.Descendants(ns + "NoWarn")
                    .SelectMany(e => (e.Value ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                    .ToArray();
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException($"Could not locate {relativePath} from the test binary.");
    }

    [Fact]
    public void Root_props_does_not_suppress_CA1707()
    {
        var nowarn = ReadNoWarnValues("Directory.Build.props");
        Assert.DoesNotContain("CA1707", nowarn, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_props_suppresses_CA1707()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(NoWarnScopeTests).Assembly.Location)!);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "tests", "Directory.Build.props");
            if (File.Exists(candidate))
            {
                var text = File.ReadAllText(candidate);
                var doc = XDocument.Parse(text);
                var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;
                var nowarn = doc.Descendants(ns + "NoWarn")
                    .SelectMany(e => (e.Value ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                    .ToArray();
                Assert.Contains("CA1707", nowarn, StringComparer.OrdinalIgnoreCase);
                return;
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate tests/Directory.Build.props.");
    }

}
