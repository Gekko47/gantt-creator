namespace GanttCreator.Architecture.Tests;

// R1.4 DEBUG-only environment-variable guard. Drives EnvVarGuardScanner across
// the whole src/ tree and proves the scanner accepts genuine DEBUG-only reads
// and rejects failure modes (plain #if DEBUG with #else, ambiguous conditions
// such as DEBUG || FEATURE, !DEBUG, unguarded reads, call mentions in
// comments/strings, and malformed/unbalanced directives). The test assembly
// under test is loaded from its bin/ output via Assembly.Location, so the
// artifact-source marker points at the verify-script build step that produces
// it (docs/04-TEST-STRATEGY.md build-pipeline traceability).
// artifact-source: verify-quick.ps1 -> 'build Release -warnaserror'
public sealed class DebugHookGuardTests
{
    private const string HookRelativePath = "src/GanttCreator.AddIn/DebugForceFailure.cs";

    [Fact]
    public void DebugForceFailure_hook_keeps_the_env_var_read_inside_ifdef_debug()
    {
        var findings = EnvVarGuardScanner.Scan(File.ReadAllText(LocateRepoFile(HookRelativePath)));

        Assert.Empty(findings);
    }

    [Fact]
    public void Repository_wide_no_env_var_read_survives_into_release()
    {
        var root = FindRepoRoot();
        var src = Path.Combine(root, "src");
        var files = Directory.EnumerateFiles(src, "*.cs", SearchOption.AllDirectories)
            .Where(p => !IsBuildOrGenerated(p))
            .ToList();

        Assert.Contains(files, f => f.EndsWith("GanttCreator.AddIn" + Path.DirectorySeparatorChar + "DebugForceFailure.cs", StringComparison.Ordinal)
            || f.EndsWith("GanttCreator.AddIn/DebugForceFailure.cs", StringComparison.Ordinal));

        var findings = new List<string>();
        foreach (var file in files)
        {
            var local = EnvVarGuardScanner.Scan(File.ReadAllText(file));
            findings.AddRange(local.Select(f => $"{Path.GetRelativePath(root, file)}: {f}"));
        }

        Assert.Empty(findings);
    }

    // __POS__
    #region Positive controls

    [Fact]
    public void Scanner_accepts_plain_ifdef_debug()
    {
        const string source = """
            internal static class GoodHook
            {
                internal static bool IsRequested(string commandName)
            #if DEBUG
                {
                    var requested = Environment.GetEnvironmentVariable("X");
                    return requested == commandName;
            #else
                    return false;
            #endif
                }
            }
            """;

        Assert.Empty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_accepts_ifdef_debug_and_feature()
    {
        const string source = """
            #if DEBUG && FEATURE
                    var requested = Environment.GetEnvironmentVariable("X");
            #endif
            """;

        Assert.Empty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_accepts_nested_ifdef_debug_under_any_parent()
    {
        const string source = """
            #if PARENT
            #if DEBUG
                    var requested = Environment.GetEnvironmentVariable("X");
            #endif
            #endif
            """;

        Assert.Empty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_ignores_call_mention_in_line_comment()
    {
        const string source = """
            // Environment.GetEnvironmentVariable("X")
            var requested = Environment.GetEnvironmentVariable("X");
            """;

        var findings = EnvVarGuardScanner.Scan(source);
        Assert.Single(findings);
        Assert.Contains("Line 2", findings[0], StringComparison.Ordinal);
    }

    [Fact]
    public void Scanner_ignores_call_mention_in_block_comment()
    {
        const string source = """
            /* Environment.GetEnvironmentVariable("X") */
            var requested = Environment.GetEnvironmentVariable("X");
            """;

        var findings = EnvVarGuardScanner.Scan(source);
        Assert.Single(findings);
        Assert.Contains("Line 2", findings[0], StringComparison.Ordinal);
    }

    [Fact]
    public void Scanner_ignores_call_mention_in_string_literals()
    {
        const string source = "var a = \"Environment.GetEnvironmentVariable(\";\r\n"
            + "var b = @\"Environment.GetEnvironmentVariable(\";\r\n"
            + "var c = $\"prefix {1} Environment.GetEnvironmentVariable(\";\r\n"
            + "var raw = \"\"\"\r\n"
            + "    Environment.GetEnvironmentVariable(\r\n"
            + "\"\"\";\r\n";

        Assert.Empty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_accepts_call_mention_in_char_literal()
    {
        const string source = """
            var c = '"';
            var requested = Environment.GetEnvironmentVariable("X");
            """;

        var findings = EnvVarGuardScanner.Scan(source);
        Assert.Single(findings);
        Assert.Contains("Line 2", findings[0], StringComparison.Ordinal);
    }

    #endregion

    // __ENVVAR__
    #region Env-var call variants

    [Fact]
    public void Scanner_rejects_plural_env_var_read()
    {
        const string source = """
            var all = Environment.GetEnvironmentVariables();
            """;

        Assert.NotEmpty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_rejects_expand_env_var_read()
    {
        const string source = """
            var path = Environment.ExpandEnvironmentVariables("%PATH%");
            """;

        Assert.NotEmpty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_rejects_spaced_dot_env_var_read()
    {
        const string source = """
            var requested = Environment . GetEnvironmentVariable("X");
            """;

        Assert.NotEmpty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_accepts_plural_read_inside_ifdef_debug()
    {
        const string source = """
            #if DEBUG
                    var all = Environment.GetEnvironmentVariables();
            #endif
            """;

        Assert.Empty(EnvVarGuardScanner.Scan(source));
    }

    // __MULTILINE__
    #region Multiline env-var call detection

    [Fact]
    public void Scanner_rejects_multiline_unguarded_read_on_correct_line()
    {
        const string source = """
            var requested =
                Environment.GetEnvironmentVariable(
                    "X");
            """;

        var findings = EnvVarGuardScanner.Scan(source);
        Assert.Contains(findings, f => f.Contains("Line 2", StringComparison.Ordinal)
            && f.Contains("active guard: True", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_accepts_multiline_read_inside_ifdef_debug()
    {
        const string source = """
            #if DEBUG
                var requested =
                    Environment.GetEnvironmentVariable(
                        "X");
            #endif
            """;

        Assert.Empty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_rejects_multiline_call_just_because_it_is_multiline()
    {
        const string source = """
            #if DEBUG || FEATURE
                var requested =
                    Environment.GetEnvironmentVariable(
                        "X");
            #endif
            """;

        Assert.NotEmpty(EnvVarGuardScanner.Scan(source));
    }

    #endregion

    // __BARE__
    #region Bare GetEnvironmentVariable detection

    [Fact]
    public void Scanner_rejects_bare_get_env_var_read()
    {
        const string source = """
            var requested = GetEnvironmentVariable("X");
            """;

        var findings = EnvVarGuardScanner.Scan(source);
        Assert.Contains(findings, f => f.Contains("Line 1", StringComparison.Ordinal)
            && f.Contains("active guard: True", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_accepts_bare_read_inside_ifdef_debug()
    {
        const string source = """
            #if DEBUG
                var requested = GetEnvironmentVariable("X");
            #endif
            """;

        Assert.Empty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_accepts_bare_read_with_optional_environment_prefix_and_whitespace()
    {
        const string source = """
            #if DEBUG
                var requested = Environment   .   GetEnvironmentVariable("X");
            #endif
            """;

        Assert.Empty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_rejects_bare_read_in_else_branch()
    {
        const string source = """
            #if DEBUG
                return false;
            #else
                var requested = GetEnvironmentVariable("X");
            #endif
            """;

        Assert.Contains(
            EnvVarGuardScanner.Scan(source),
            f => f.Contains("active guard", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_rejects_bare_read_when_condition_is_debug_or_something()
    {
        const string source = """
            #if DEBUG || FEATURE
                var requested = GetEnvironmentVariable("X");
            #endif
            """;

        Assert.NotEmpty(EnvVarGuardScanner.Scan(source));
    }

    #endregion

    #endregion

    // __NEG__
    #region Negative controls

    [Fact]
    public void Scanner_rejects_unguarded_read()
    {
        const string source = """
            var requested = Environment.GetEnvironmentVariable("X");
            """;

        var findings = EnvVarGuardScanner.Scan(source);
        Assert.Contains(findings, f => f.Contains("Line 1", StringComparison.Ordinal)
            && f.Contains("active guard: True", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_rejects_debug_or_feature()
    {
        const string source = """
            #if DEBUG || FEATURE
                    var requested = Environment.GetEnvironmentVariable("X");
            #endif
            """;

        Assert.NotEmpty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_rejects_not_debug()
    {
        const string source = """
            #if !DEBUG
                    var requested = Environment.GetEnvironmentVariable("X");
            #endif
            """;

        Assert.NotEmpty(EnvVarGuardScanner.Scan(source));
    }

    [Fact]
    public void Scanner_rejects_else_of_ifdef_debug()
    {
        const string source = """
            #if DEBUG
                    return false;
            #else
                    var requested = Environment.GetEnvironmentVariable("X");
            #endif
            """;

        var findings = EnvVarGuardScanner.Scan(source);
        Assert.Contains(findings, f => f.Contains("Line 4", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_rejects_elif_feature_after_ifdef_debug()
    {
        const string source = """
            #if DEBUG
                    return false;
            #elif FEATURE
                    var requested = Environment.GetEnvironmentVariable("X");
            #endif
            """;

        Assert.NotEmpty(EnvVarGuardScanner.Scan(source));
    }

    #endregion

    // __STRUCT__
    #region Directive-structure controls

    [Fact]
    public void Scanner_rejects_if_without_expression()
    {
        const string source = """
            #if
                    var requested = Environment.GetEnvironmentVariable("X");
            #endif
            """;

        var findings = EnvVarGuardScanner.Scan(source);
        Assert.Contains(findings, f => f.Contains("no expression", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_rejects_elif_without_matching_if()
    {
        const string source = """
            #elif DEBUG
                    return false;
            #endif
            """;

        Assert.Contains(
            EnvVarGuardScanner.Scan(source),
            f => f.Contains("without a matching #if", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_rejects_else_without_matching_if()
    {
        const string source = """
            #else
                    return false;
            #endif
            """;

        Assert.Contains(
            EnvVarGuardScanner.Scan(source),
            f => f.Contains("without a matching #if", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_rejects_endif_without_matching_if()
    {
        const string source = """
            var x = 1;
            #endif
            """;

        Assert.Contains(
            EnvVarGuardScanner.Scan(source),
            f => f.Contains("without a matching #if", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_rejects_unterminated_if()
    {
        const string source = """
            #if DEBUG
                    var requested = Environment.GetEnvironmentVariable("X");
            """;

        Assert.Contains(
            EnvVarGuardScanner.Scan(source),
            f => f.Contains("unterminated #if", StringComparison.Ordinal));
    }

    [Fact]
    public void Scanner_rejects_malformed_expression()
    {
        const string source = """
            #if DEBUG &&
                    var requested = Environment.GetEnvironmentVariable("X");
            #endif
            """;

        Assert.Contains(
            EnvVarGuardScanner.Scan(source),
            f => f.Contains("malformed preprocessor expression", StringComparison.Ordinal));
    }

    #endregion

    // __INFRA__
    #region Infrastructure

    private static bool IsBuildOrGenerated(string path)
    {
        var normalized = path.Replace('\\', '/');
        if (normalized.Contains("/obj/", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("/bin/", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var file = Path.GetFileName(path);
        if (file.EndsWith(".Designer.cs", StringComparison.OrdinalIgnoreCase)
            || file.EndsWith(".g.cs", StringComparison.OrdinalIgnoreCase)
            || file.EndsWith(".AssemblyInfo.cs", StringComparison.OrdinalIgnoreCase)
            || file.EndsWith(".GlobalUsings.g.cs", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(DebugHookGuardTests).Assembly.Location)!);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "GanttCreator.slnx"))
                && Directory.Exists(Path.Combine(dir.FullName, "src"))
                && Directory.Exists(Path.Combine(dir.FullName, "tests")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the repository root from the test binary.");
    }

    private static string LocateRepoFile(string relativePath)
    {
        var root = FindRepoRoot();
        var candidate = Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar));
        if (File.Exists(candidate))
        {
            return candidate;
        }

        throw new DirectoryNotFoundException($"Could not locate '{relativePath}' from the test binary.");
    }

    #endregion
}
