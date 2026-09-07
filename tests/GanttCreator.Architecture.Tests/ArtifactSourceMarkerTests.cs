using System.Text.RegularExpressions;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// W10 artifact-traceability enforcement (see docs/02-ARCHITECTURE.md
/// build-pipeline artifact contract). A test file that references a
/// <c>bin/</c> or <c>publish/</c> path must carry a
/// <c>// artifact-source:</c> marker comment naming the verify-script
/// step that produces the artifact, so a fresh clone knows which gate
/// must run before the test will pass. The first increment landed in
/// Phase A (AddInAssemblyTests, CoreBoundaryTests); this test makes
/// the rule stick.
/// </summary>
public sealed partial class ArtifactSourceMarkerTests
{
    // Patterns that count as a "test consuming a verify-script artifact".
    // The marker must be a single-line `// artifact-source:` comment, with
    // the producing step name after the colon. Documented in
    // docs/02-ARCHITECTURE.md.
    private static readonly Regex BinPublishRef = BinPublishRefImpl();
    private static readonly Regex MarkerLine = MarkerLineImpl();

    [GeneratedRegex(@"\b(bin|publish)\b[/\\]", RegexOptions.IgnoreCase)]
    private static partial Regex BinPublishRefImpl();

    [GeneratedRegex(@"^\s*//\s*artifact-source\s*:\s*\S", RegexOptions.IgnoreCase)]
    private static partial Regex MarkerLineImpl();

    [Fact]
    public void Every_test_file_referencing_bin_or_publish_has_artifact_source_marker()
    {
        var testsRoot = LocateTestsRoot();
        var files = Directory.GetFiles(testsRoot, "*.cs", SearchOption.AllDirectories);
        var offenders = new List<string>();

        foreach (var file in files)
        {
            // The check is a self-referential positive-control fixture; the
            // file legitimately contains "bin" and "publish" as test
            // inputs. Excluding it keeps the rule honest.
            if (Path.GetFileName(file).Equals("ArtifactSourceMarkerTests.cs", StringComparison.Ordinal))
            {
                continue;
            }

            var text = File.ReadAllText(file);
            if (!BinPublishRef.IsMatch(text))
            {
                continue;
            }

            var hasMarker = false;
            foreach (var line in text.Split('\n'))
            {
                if (MarkerLine.IsMatch(line))
                {
                    hasMarker = true;
                    break;
                }
            }

            if (!hasMarker)
            {
                offenders.Add(RelativeToRepo(file));
            }
        }

        Assert.True(
            offenders.Count == 0,
            "Test files reference bin/ or publish/ but lack the required " +
            "`// artifact-source: <step-name>` marker comment. Add the marker " +
            "to the file (the comment text is free-form; cite the verify-script " +
            "step that produces the artifact). Offenders: " +
            string.Join(", ", offenders));
    }

    // Source strings of the patterns above. Kept in sync with the
    // [GeneratedRegex] attributes; the xUnit analyzer forbids
    // Assert.True(regex.IsMatch(...)) and Assert.Matches(Regex, string) is
    // not in xUnit 2.9. Asserting against the source pattern string is
    // the only analyzer-clean form.
    private const string BinPublishPattern = @"\b(bin|publish)\b[/\\]";
    private const string MarkerLinePattern = @"^\s*//\s*artifact-source\s*:\s*\S";

    [Fact]
    public void Positive_control_synthetic_file_with_marker_is_clean()
    {
        // Build an in-memory file reference and a marker line; the regex
        // check must accept it. Guards the helper itself. Note: the
        // BinPublishPattern requires "bin" or "publish" followed by a
        // path separator, so the synthetic path must use one.
        var sample = "// artifact-source: verify-quick.ps1 -> 'publish AddIn (packed XLL)'"
            + Environment.NewLine
            + "var path = @\"src\\GanttCreator.AddIn\\bin\\Release\";";

        Assert.Matches(BinPublishPattern, sample);
        Assert.Matches(MarkerLinePattern, sample.Split('\n')[0]);
    }

    private static string LocateTestsRoot()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(ArtifactSourceMarkerTests).Assembly.Location)!);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "tests");
            if (Directory.Exists(candidate))
            {
                return candidate;
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate tests/ from the test binary.");
    }

    private static string RelativeToRepo(string path)
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(ArtifactSourceMarkerTests).Assembly.Location)!);
        while (dir is not null)
        {
            var rootCandidate = dir.FullName;
            if (path.StartsWith(rootCandidate, StringComparison.OrdinalIgnoreCase))
            {
                return path[(rootCandidate.Length + 1)..];
            }
            dir = dir.Parent;
        }
        return path;
    }
}