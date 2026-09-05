using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Enforces that test-only diagnostic suppressions remain scoped to test
/// projects. CA1707 (identifiers should have correct suffix) and
/// IDE0011 (add braces) must be suppressed in
/// <c>tests/Directory.Build.props</c> only, never in the root
/// <c>Directory.Build.props</c>, so production code still enforces the
/// rules (see docs/04-TEST-STRATEGY.md "NoWarn scope" and
/// docs/08-TEST-CHECKLIST.md).
/// </summary>
public sealed class NoWarnScopeTests
{
    /// <summary>
    /// Reads the <c>NoWarn</c> element values from a Directory.Build.props
    /// file. The same shape applies to both root and tests props.
    /// </summary>
    /// <param name="relativePath">Path relative to the walk root. May not
    /// start with the literal string "tests/" because the walk skips
    /// any directory named "tests" so the root props can be located
    /// without being shadowed by tests/Directory.Build.props.</param>
    private static string[] ReadNoWarnValues(string relativePath)
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(NoWarnScopeTests).Assembly.Location)!);
        while (dir is not null)
        {
            // Skip the "tests" directory tree so the root props is not
            // shadowed by tests/Directory.Build.props when the caller
            // asks for the root. The tests props is loaded directly
            // when its path is supplied.
            if (dir.Name == "tests")
            {
                dir = dir.Parent;
                continue;
            }
            var candidate = Path.Combine(dir.FullName, relativePath);
            if (File.Exists(candidate))
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

    /// <summary>One parameter row per (ID) we expect to be scoped to test projects.</summary>
    public static TheoryData<string> TestOnlySuppressionIds() => new()
    {
        "CA1707",
        "IDE0011",
    };

    [Theory]
    [MemberData(nameof(TestOnlySuppressionIds))]
    public void Root_props_does_not_suppress_test_only_id(string id)
    {
        var nowarn = ReadNoWarnValues("Directory.Build.props");
        Assert.DoesNotContain(id, nowarn, StringComparer.OrdinalIgnoreCase);
    }

    [Theory]
    [MemberData(nameof(TestOnlySuppressionIds))]
    public void Test_props_suppresses_test_only_id(string id)
    {
        var nowarn = ReadNoWarnValues(Path.Combine("tests", "Directory.Build.props"));
        Assert.Contains(id, nowarn, StringComparer.OrdinalIgnoreCase);
    }

}
