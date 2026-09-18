using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;
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

    // Recognized verify scripts and their step names for artifact-source markers.
    // A marker is valid only when it references a known script AND a known step.
    private static readonly HashSet<string> KnownScripts = new(StringComparer.OrdinalIgnoreCase)
    {
        "verify-quick.ps1",
        "verify.ps1",
    };

    private static readonly HashSet<string> KnownSteps = new(StringComparer.OrdinalIgnoreCase)
    {
        "publish AddIn (packed XLL)",
        "build Release -warnaserror",
    };

    [GeneratedRegex(@"\b(bin|publish)\b(?:[/\\]|"")", RegexOptions.IgnoreCase)]
    private static partial Regex BinPublishRefImpl();

    [GeneratedRegex(@"^\s*//\s*artifact-source\s*:\s*\S+\s*->\s*'", RegexOptions.IgnoreCase)]
    private static partial Regex MarkerLineImpl();

    /// <summary>
    /// Language version pinned to C# 14, the level the repository compiles
    /// with under the .NET 10 SDK. Mirrors the
    /// <c>ScannerLanguageVersion</c> pin in <c>EnvVarGuardScanner</c> so
    /// comment masking cannot drift when the parser package and the
    /// repository language level move apart. The value is 1400, expressed
    /// as a cast because the 4.14.0 package's compile asset has no named
    /// CSharp14 member.
    /// </summary>
    private const LanguageVersion CommentMaskLanguageVersion = (LanguageVersion)1400;

    private static readonly string[] DebugPreprocessorSymbol = new[] { "DEBUG" };

    [Fact]
    public void Every_test_file_referencing_bin_or_publish_has_artifact_source_marker()
    {
        var testsRoot = LocateTestsRoot();
        var files = EnumerateSourceFiles(testsRoot, "*.cs");
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

            if (HasMissingMarker(file))
            {
                offenders.Add(RelativeToRepo(file));
            }
        }

        Assert.True(
            offenders.Count == 0,
            "Test files reference bin/ or publish/ but lack the required " +
            "`// artifact-source: <script> -> '<step>` marker comment. " +
            "The marker must reference a known verify script and step. " +
            "Known scripts: verify-quick.ps1, verify.ps1. " +
            "Known steps: publish AddIn (packed XLL), build Release -warnaserror. " +
            "Offenders: " +
            string.Join(", ", offenders));
    }

    // Source strings of the patterns above. Kept in sync with the
    // [GeneratedRegex] attributes; the xUnit analyzer forbids
    // Assert.True(regex.IsMatch(...)) and Assert.Matches(Regex, string) is
    // not in xUnit 2.9. Asserting against the source pattern string is
    // the only analyzer-clean form.
    private const string BinPublishPattern = @"\b(bin|publish)\b(?:[/\\]|"")";
    private const string MarkerLinePattern = @"^\s*//\s*artifact-source\s*:\s*\S+\s*->\s*'";

    [Fact]
    public void Positive_control_synthetic_file_with_marker_is_clean()
    {
        // Build an in-memory file reference with a marker line; the regex
        // check must accept it. Guards the helper itself. The synthetic
        // path uses a path separator; the quoted-segment (Path.Combine)
        // form is covered by the regression tests below.
        var sample = "// artifact-source: verify-quick.ps1 -> 'publish AddIn (packed XLL)'"
            + Environment.NewLine
            + "var path = @\"src\\GanttCreator.AddIn\\bin\\Release\";";

        Assert.Matches(BinPublishPattern, sample);
        Assert.Matches(MarkerLinePattern, sample.Split('\n')[0]);
    }

    private static bool HasMissingMarker(string filePath)
    {
        // True when the file references a bin/ or publish/ artifact segment
        // (as a path or as a quoted Path.Combine argument) but carries no
        // artifact-source marker line, or the marker references an unknown
        // script or step.
        var text = File.ReadAllText(filePath);
        // The artifact signal lives in code and string literals, so only
        // comments are masked: masking strings would hide the quoted
        // Path.Combine("bin", ...) form the rule exists to catch, while a
        // stray "bin" in a comment must not force a marker.
        if (!BinPublishRef.IsMatch(MaskComments(text)))
        {
            return false;
        }

        foreach (var line in text.Split('\n'))
        {
            var match = MarkerLine.Match(line);
            if (match.Success)
            {
                // Extract the script and step from the marker.
                // Format: // artifact-source: <script> -> '<step>'
                var markerContent = line[match.Index..].TrimStart();
                var colonIdx = markerContent.IndexOf(':', StringComparison.Ordinal);
                var arrowIdx = markerContent.IndexOf("->", StringComparison.Ordinal);
                var quoteStart = markerContent.IndexOf('\'', arrowIdx);
                var quoteEnd = markerContent.IndexOf('\'', quoteStart + 1);

                if (colonIdx > 0 && arrowIdx > colonIdx && quoteStart > arrowIdx && quoteEnd > quoteStart)
                {
                    var script = markerContent[(colonIdx + 1)..arrowIdx].Trim();
                    var step = markerContent[(quoteStart + 1)..quoteEnd];

                    // Valid marker requires both a known script and a known step.
                    if (KnownScripts.Contains(script) && KnownSteps.Contains(step))
                    {
                        return false;
                    }
                }

                // Marker present but invalid (unknown script or step).
                return true;
            }
        }

        return true;
    }

    [Fact]
    public void Regression_quoted_bin_publish_path_combine_segments_without_marker_are_flagged()
    {
        var td = Path.Combine(Path.GetTempPath(), "asm-flag-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(td);
        try
        {
            var file = Path.Combine(td, "PassTests.cs");
            File.WriteAllText(file,
                "class PassTests {" + Environment.NewLine +
                "  static void M() {" + Environment.NewLine +
                "    var binDir = Path.Combine(root, \"bin\", \"Release\");" + Environment.NewLine +
                "    var pubDir = Path.Combine(root, \"publish\", \"out\");" + Environment.NewLine +
                "  }" + Environment.NewLine +
                "}" + Environment.NewLine);

            Assert.True(HasMissingMarker(file),
                "A Path.Combine quoted bin/publish consumer without a marker must be flagged.");
        }
        finally
        {
            Directory.Delete(td, recursive: true);
        }
    }

    [Fact]
    public void Regression_quoted_bin_path_combine_with_marker_is_clean()
    {
        var td = Path.Combine(Path.GetTempPath(), "asm-ok-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(td);
        try
        {
            var file = Path.Combine(td, "PassTests.cs");
            File.WriteAllText(file,
                "// artifact-source: verify-quick.ps1 -> 'build Release -warnaserror'" + Environment.NewLine +
                "class PassTests {" + Environment.NewLine +
                "  static void M() {" + Environment.NewLine +
                "    var binDir = Path.Combine(root, \"bin\", \"Release\");" + Environment.NewLine +
                "  }" + Environment.NewLine +
                "}" + Environment.NewLine);

            Assert.False(HasMissingMarker(file), "A marked Path.Combine bin consumer must be clean.");
        }
        finally
        {
            Directory.Delete(td, recursive: true);
        }
    }

    [Fact]
    public void Regression_marker_with_unknown_script_is_rejected()
    {
        var td = Path.Combine(Path.GetTempPath(), "asm-inv-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(td);
        try
        {
            var file = Path.Combine(td, "FailTests.cs");
            File.WriteAllText(file,
                "// artifact-source: unknown-script.ps1 -> 'some step'" + Environment.NewLine +
                "class FailTests {" + Environment.NewLine +
                "  static void M() {" + Environment.NewLine +
                "    var binDir = Path.Combine(root, \"bin\", \"Release\");" + Environment.NewLine +
                "  }" + Environment.NewLine +
                "}" + Environment.NewLine);

            Assert.True(HasMissingMarker(file),
                "A marker referencing an unknown script must be rejected.");
        }
        finally
        {
            Directory.Delete(td, recursive: true);
        }
    }

    [Fact]
    public void Regression_marker_with_unknown_step_is_rejected()
    {
        var td = Path.Combine(Path.GetTempPath(), "asm-inv2-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(td);
        try
        {
            var file = Path.Combine(td, "FailTests2.cs");
            File.WriteAllText(file,
                "// artifact-source: verify-quick.ps1 -> 'unknown step'" + Environment.NewLine +
                "class FailTests2 {" + Environment.NewLine +
                "  static void M() {" + Environment.NewLine +
                "    var binDir = Path.Combine(root, \"bin\", \"Release\");" + Environment.NewLine +
                "  }" + Environment.NewLine +
                "}" + Environment.NewLine);

            Assert.True(HasMissingMarker(file),
                "A marker referencing an unknown step must be rejected.");
        }
        finally
        {
            Directory.Delete(td, recursive: true);
        }
    }

    [Fact]
    public void Regression_bin_word_in_line_comment_does_not_require_marker()
    {
        var td = Path.Combine(Path.GetTempPath(), "asm-cmt1-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(td);
        try
        {
            var file = Path.Combine(td, "CommentTests.cs");
            File.WriteAllText(file,
                "// clean up the bin/ folder before running" + Environment.NewLine +
                "class CommentTests {" + Environment.NewLine +
                "  static void M() { }" + Environment.NewLine +
                "}" + Environment.NewLine);

            Assert.False(HasMissingMarker(file),
                "A 'bin/' mention inside a line comment must not force a marker.");
        }
        finally
        {
            Directory.Delete(td, recursive: true);
        }
    }

    [Fact]
    public void Regression_publish_word_in_block_comment_does_not_require_marker()
    {
        var td = Path.Combine(Path.GetTempPath(), "asm-cmt2-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(td);
        try
        {
            var file = Path.Combine(td, "BlockCommentTests.cs");
            File.WriteAllText(file,
                "/* publish/ output goes here */" + Environment.NewLine +
                "class BlockCommentTests {" + Environment.NewLine +
                "  static void M() { }" + Environment.NewLine +
                "}" + Environment.NewLine);

            Assert.False(HasMissingMarker(file),
                "A 'publish/' mention inside a block comment must not force a marker.");
        }
        finally
        {
            Directory.Delete(td, recursive: true);
        }
    }

    [Fact]
    public void HasMissingMarker_flags_disabled_preprocessor_branch_bin_reference()
    {
        // Regression: a bin/ or publish/ reference inside a non-DEBUG
        // conditional branch must be detected. The previous MaskComments
        // treated DisabledTextTrivia as comments, masking artifact-source
        // references in disabled branches so the rule could not catch them.
        var source = "#if !DEBUG" + Environment.NewLine +
                     "var path = Path.Combine(\"bin\", \"output\");" + Environment.NewLine +
                     "#endif" + Environment.NewLine;
        var td = Path.Combine(Path.GetTempPath(), "artifact-marker-regression-" + Guid.NewGuid());
        Directory.CreateDirectory(td);
        var file = Path.Combine(td, "NonDebugConditional.cs");
        try
        {
            File.WriteAllText(file, source);
            Assert.True(HasMissingMarker(file),
                "A bin/ reference inside a non-DEBUG conditional branch must be flagged.");
        }
        finally
        {
            Directory.Delete(td, recursive: true);
        }
    }

    [Fact]
    public void HasMissingMarker_flags_disabled_preprocessor_branch_publish_reference()
    {
        // Regression: a publish/ reference inside a non-DEBUG conditional
        // branch must be detected (same root cause as the bin/ case above).
        var source = "#if !DEBUG" + Environment.NewLine +
                     "var path = Path.Combine(\"publish\", \"output\");" + Environment.NewLine +
                     "#endif" + Environment.NewLine;
        var td = Path.Combine(Path.GetTempPath(), "artifact-marker-regression-" + Guid.NewGuid());
        Directory.CreateDirectory(td);
        var file = Path.Combine(td, "NonDebugConditionalPublish.cs");
        try
        {
            File.WriteAllText(file, source);
            Assert.True(HasMissingMarker(file),
                "A publish/ reference inside a non-DEBUG conditional branch must be flagged.");
        }
        finally
        {
            Directory.Delete(td, recursive: true);
        }
    }

    /// <summary>
    /// Masks comment trivia with blanks (newlines preserved) so the
    /// <c>bin</c>/<c>publish</c> reference scan sees code and string
    /// literals but never comment text. String literals are deliberately
    /// kept: the quoted <c>Path.Combine(..., "bin", ...)</c> form is the
    /// primary signal this rule exists to catch. Two parses (default
    /// symbols and DEBUG defined) cover comments inside disabled
    /// preprocessor branches on the common DEBUG-only guard shape.
    /// </summary>
    private static string MaskComments(string source)
    {
        var spansToMask = new List<TextSpan>();
        spansToMask.AddRange(CollectCommentSpans(
            source, new CSharpParseOptions(languageVersion: CommentMaskLanguageVersion)));
        spansToMask.AddRange(CollectCommentSpans(
            source,
            new CSharpParseOptions(
                languageVersion: CommentMaskLanguageVersion,
                preprocessorSymbols: DebugPreprocessorSymbol)));

        var masked = source.ToCharArray();
        int n = source.Length;
        foreach (var span in spansToMask)
        {
            int start = span.Start;
            int end = Math.Min(span.End, n);
            for (int k = start; k < end && k < masked.Length; k++)
            {
                if (masked[k] != '\n')
                {
                    masked[k] = ' ';
                }
            }
        }

        return new string(masked);
    }

    private static List<TextSpan> CollectCommentSpans(string source, CSharpParseOptions parseOptions)
    {
        var spans = new List<TextSpan>();
        var syntaxTree = CSharpSyntaxTree.ParseText(source, parseOptions);
        var root = syntaxTree.GetRoot();

        foreach (var token in root.DescendantTokens(descendIntoTrivia: true))
        {
            foreach (var trivia in token.LeadingTrivia)
            {
                AddCommentSpanIfComment(spans, trivia);
            }

            foreach (var trivia in token.TrailingTrivia)
            {
                AddCommentSpanIfComment(spans, trivia);
            }
        }

        return spans;
    }

    private static void AddCommentSpanIfComment(List<TextSpan> spans, SyntaxTrivia trivia)
    {
        if (trivia.IsKind(SyntaxKind.SingleLineCommentTrivia)
            || trivia.IsKind(SyntaxKind.MultiLineCommentTrivia)
            || trivia.IsKind(SyntaxKind.SingleLineDocumentationCommentTrivia)
            || trivia.IsKind(SyntaxKind.MultiLineDocumentationCommentTrivia))
        {
            spans.Add(new TextSpan(trivia.Span.Start, trivia.Span.Length));
        }
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
            if (Directory.Exists(Path.Combine(dir.FullName, "tests")))
            {
                return Path.GetRelativePath(dir.FullName, path);
            }
            dir = dir.Parent;
        }
        return path;
    }

    /// <summary>
    /// Recursively enumerates source files under <paramref name="root"/>,
    /// excluding build output directories (bin/obj) so results are
    /// independent of build artifacts. Preserves the caller-supplied
    /// search pattern (e.g. "*.cs") and is case-insensitive on the
    /// directory-name exclusion.
    /// </summary>
    private static IEnumerable<string> EnumerateSourceFiles(string root, string searchPattern)
    {
        var exclusions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "bin",
            "obj",
        };

        return EnumerateFilesRecursive(root, searchPattern, exclusions);
    }

    private static IEnumerable<string> EnumerateFilesRecursive(string dir, string searchPattern, HashSet<string> exclusions)
    {
        IEnumerable<string> EnumerateDir(string current)
        {
            // Skip excluded directories entirely so we never descend into
            // build output trees.
            var name = Path.GetFileName(current);
            if (exclusions.Contains(name))
            {
                return Enumerable.Empty<string>();
            }

            var files = Directory.EnumerateFiles(current, searchPattern);
            var subDirs = Directory.EnumerateDirectories(current);

            return files.Concat(subDirs.SelectMany(EnumerateDir));
        }

        return EnumerateDir(dir);
    }
}
