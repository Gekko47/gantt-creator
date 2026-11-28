using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// ADR-0020 D4: the status bar's visibility is the user's Excel setting, so only the
/// application-state scope may hide it. A command that calls the scope's visibility
/// toggle makes the bar disappear and reappear around its work, which is the
/// show/hide flicker this guard exists to prevent.
/// </summary>
/// <remarks>
/// <para>
/// The Type-change heal was the offender: it called the toggle before healing a row,
/// so every <c>Type</c> edit hid the bar and restored it on the same macro. The
/// refresh path had already dropped the same call for the same reason (ADR-0020 D4),
/// and nothing structural stopped the heal from re-adopting it. This test is that
/// structure.
/// </para>
/// <para>
/// The scan is SYNTAX-based, like the neighbouring authority tests: an identifier
/// token is matched, so a comment or a string literal naming the member is not a
/// mention. It also matches the declaration token, which is the only shape the two
/// scope files use.
/// </para>
/// </remarks>
public sealed class StatusBarVisibilityAuthorityTests
{
    /// <summary>The scope member that writes <c>Application.DisplayStatusBar</c>.</summary>
    private const string _visibilityMember = "SuppressStatusBar";

    /// <summary>Production source roots that could reach the Excel application object.</summary>
    private static readonly string[] _productionRoots =
    [
        "src/GanttCreator.AddIn",
        "src/GanttCreator.Office",
    ];

    /// <summary>
    /// The files that necessarily declare and implement the member, so naming it is
    /// their job rather than a competing caller.
    /// </summary>
    private static readonly string[] _scopeDeclarationPaths =
    [
        "src/GanttCreator.Office/ExcelApplicationStateScope.cs",
        "src/GanttCreator.Office/IApplicationStateScope.cs",
    ];

    /// <summary>
    /// No production file outside the scope may hide the status bar.
    /// </summary>
    [Fact]
    public void Only_the_state_scope_hides_the_status_bar()
    {
        string repositoryRoot = FindRepositoryRoot();
        List<string> callers = [];

        foreach (string root in _productionRoots)
        {
            foreach (string file in EnumerateProductionFiles(repositoryRoot, root))
            {
                if (_scopeDeclarationPaths.Contains(file, StringComparer.Ordinal))
                {
                    continue;
                }

                if (Mentions(file, repositoryRoot, _visibilityMember))
                {
                    callers.Add(file);
                }
            }
        }

        Assert.True(
            callers.Count == 0,
            "Status-bar visibility is the user's Excel setting (ADR-0020 D4), so hiding it is the "
            + "state scope's job alone. These product files toggle it: " + string.Join(", ", callers));
    }

    /// <summary>
    /// The scope files still name the member, so the guard above passes because it
    /// excludes the right files rather than because the matcher finds nothing.
    /// </summary>
    [Fact]
    public void The_scope_declarations_do_name_the_member()
    {
        string repositoryRoot = FindRepositoryRoot();

        foreach (string path in _scopeDeclarationPaths)
        {
            Assert.True(
                Mentions(path, repositoryRoot, _visibilityMember),
                path + " no longer names " + _visibilityMember + "; the allowlist above is stale.");
        }
    }

    /// <summary>
    /// The heal command was the offending caller, and it must stay clean.
    /// </summary>
    [Fact]
    public void The_type_heal_command_does_not_hide_the_status_bar() =>
        Assert.False(
            Mentions(
                "src/GanttCreator.AddIn/TypeChangeHealCommand.cs",
                FindRepositoryRoot(),
                _visibilityMember),
            "The Type heal must not hide the status bar (ADR-0020 D4).");

    /// <summary>
    /// The matcher detects a declaration and a call, and ignores comments and strings.
    /// </summary>
    /// <param name="source">Inline source parsed without a compilation.</param>
    /// <param name="expected">Whether the member is used in code.</param>
    [Theory]
    [InlineData("class C { void SuppressStatusBar() { } }", true)]
    [InlineData("class C { void M(IApplicationStateScope s) { s.SuppressStatusBar(); } }", true)]
    [InlineData("class C { void M() { /* s.SuppressStatusBar(); */ } }", false)]
    [InlineData("class C { string s = \"SuppressStatusBar\"; }", false)]
    public void The_matcher_detects_calls_but_not_text(string source, bool expected)
    {
        SyntaxNode root = CSharpSyntaxTree.ParseText(source).GetRoot();

        Assert.Equal(expected, MentionsIn(root, _visibilityMember));
    }

    private static bool Mentions(string relativePath, string repositoryRoot, string member) =>
        MentionsIn(
            CSharpSyntaxTree.ParseText(
                File.ReadAllText(Path.Combine(repositoryRoot, relativePath))).GetRoot(),
            member);

    /// <summary>
    /// Whether the syntax tree uses the named member, ignoring comments and string literals.
    /// </summary>
    /// <param name="root">The parsed source root.</param>
    /// <param name="member">The identifier to look for.</param>
    /// <returns><see langword="true"/> when an identifier token names the member.</returns>
    /// <remarks>
    /// An identifier token covers both the declaration in the scope files and a call
    /// such as <c>scope.SuppressStatusBar()</c>. A comment is trivia and a string
    /// literal is a different token kind, so neither can satisfy or trip the guard.
    /// </remarks>
    private static bool MentionsIn(SyntaxNode root, string member) =>
        root.DescendantTokens()
            .Any(token => token.IsKind(SyntaxKind.IdentifierToken) && token.ValueText == member);

    private static IEnumerable<string> EnumerateProductionFiles(string repositoryRoot, string root)
    {
        string directory = Path.Combine(repositoryRoot, root);

        // A missing root FAILS rather than being skipped: skipping is how a renamed or
        // deleted project would stop being scanned while the guard stayed green.
        Assert.True(
            Directory.Exists(directory),
            "Production root '" + root + "' does not exist under the repository root; "
            + "a missing root would be silently skipped rather than scanned.");

        foreach (string file in Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories))
        {
            yield return Path.GetRelativePath(repositoryRoot, file).Replace('\\', '/');
        }
    }

    /// <summary>
    /// Locates the repository root by walking up for <c>src/</c> and the solution file.
    /// </summary>
    /// <returns>The absolute repository root path.</returns>
    private static string FindRepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (Directory.Exists(Path.Combine(directory.FullName, "src"))
                && File.Exists(Path.Combine(directory.FullName, "GanttCreator.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not locate the repository root (a directory holding src/ and GanttCreator.slnx).");
    }
}
