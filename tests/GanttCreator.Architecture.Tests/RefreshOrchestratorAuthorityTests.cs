using System.IO;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// R4.8A: the refresh orchestrator and the scene-request factory own the pipeline
/// and the scene inputs, and nothing else re-derives them.
/// </summary>
/// <remarks>
/// <para>
/// D5's real risk is not that the orchestrator does too little but that it does not
/// stay the <em>only</em> place. A Ribbon callback that resolves the preset, or a
/// renderer that re-derives the plot range, would compile and run and disagree at
/// some margin — presenting as a layout bug rather than as duplicated logic. These
/// guards make the ownership structural.
/// </para>
/// <para>
/// The scans are SYNTAX-based via Roslyn, for the reason
/// <c>PlotGeometryAuthorityTests</c> documents: a substring search is defeated by
/// formatting in ways that would leave a guard green while examining nothing.
/// </para>
/// </remarks>
public sealed class RefreshOrchestratorAuthorityTests
{
    /// <summary>Production source roots scanned for a competing authority.</summary>
    private static readonly string[] ProductionRoots =
    [
        "src/GanttCreator.AddIn",
        "src/GanttCreator.Office",
    ];

    /// <summary>The single file permitted to drive a whole refresh.</summary>
    private const string OrchestratorPath = "src/GanttCreator.Office/GanttRefreshOrchestrator.cs";

    /// <summary>The single file permitted to call the shape reconciler.</summary>
    private const string ReconcilerPath = "src/GanttCreator.Office/ShapeReconciler.cs";

    /// <summary>The single production caller of the scene-request factory.</summary>
    private const string FactoryPath = "src/GanttCreator.Office/GanttRefreshOrchestrator.cs";

    /// <summary>
    /// The files permitted to name the factory WITHOUT being the orchestrator.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A file cannot avoid naming the type it implements, the interface it implements,
    /// or the object it constructs, so those references are declarations rather than
    /// competing authorities. Everything else that names the type is a second
    /// pipeline in waiting and is reported.
    /// </para>
    /// <list type="bullet">
    /// <item><description>
    /// <c>ISceneBuildRequestFactory.cs</c> — the port's own declaration.
    /// </description></item>
    /// <item><description>
    /// <c>ExcelSceneBuildRequestFactory.cs</c> — the port's implementation, which
    /// necessarily names both the class it declares and the interface it implements.
    /// </description></item>
    /// <item><description>
    /// <c>RefreshSheetCommand.cs</c> — <b>the composition root</b>, and the one
    /// allowed caller. D1 requires the Ribbon command to construct the live
    /// orchestrator and its collaborators; that is where the single instance is
    /// assembled, and it is stated in the command's own remarks. It is allowlisted
    /// rather than left undetected, so the exemption is visible and a THIRD caller
    /// still fails this test.
    /// </description></item>
    /// </list>
    /// </remarks>
    private static readonly string[] FactoryDeclarationPaths =
    [
        "src/GanttCreator.Office/ISceneBuildRequestFactory.cs",
        "src/GanttCreator.Office/ExcelSceneBuildRequestFactory.cs",
        "src/GanttCreator.AddIn/RefreshSheetCommand.cs",
    ];

    /// <summary>
    /// Only the orchestrator drives a whole refresh, so the pipeline order lives in
    /// one place.
    /// </summary>
    /// <remarks>
    /// A second caller would not be a second pipeline — it would be two places to keep
    /// in step, and the one that drifts is the one someone edits when adding a step.
    /// </remarks>
    [Fact]
    public void Only_the_orchestrator_calls_the_shape_reconciler()
    {
        var repositoryRoot = FindRepositoryRoot();
        var callers = new List<string>();

        foreach (string root in ProductionRoots)
        {
            foreach (string file in EnumerateProductionFiles(repositoryRoot, root))
            {
                if (file == ReconcilerPath)
                {
                    continue;
                }

                if (Mentions(file, "ShapeReconciler.Reconcile"))
                {
                    callers.Add(file);
                }
            }
        }

        Assert.True(
            callers is [OrchestratorPath] or [],
            "The shape reconciliation must be driven by the orchestrator alone, but these call it: "
            + string.Join(", ", callers));
    }

    /// <summary>
    /// Only the orchestrator builds a scene request, so D5's "the factory owns the
    /// scene inputs" is enforced rather than stated.
    /// </summary>
    [Fact]
    public void Only_the_orchestrator_creates_a_scene_build_request()
    {
        var repositoryRoot = FindRepositoryRoot();
        var callers = new List<string>();

        foreach (string root in ProductionRoots)
        {
            foreach (string file in EnumerateProductionFiles(repositoryRoot, root))
            {
                if (file == FactoryPath || FactoryDeclarationPaths.Contains(file, StringComparer.Ordinal))
                {
                    continue;
                }

                // The type is referenced to construct a request anywhere: a Ribbon
                // callback that built one would be a second authority for the plot
                // bounds, the preset, and the range all at once.
                if (Mentions(file, "ExcelSceneBuildRequestFactory") || Mentions(file, "ISceneBuildRequestFactory"))
                {
                    callers.Add(file);
                }
            }
        }

        Assert.True(
            callers is [],
            "The scene-request factory must be used by the orchestrator and the declared composition root alone, but these reference it: "
            + string.Join(", ", callers));
    }

    /// <summary>
    /// The bare-name branch really does detect the shapes that matter, proved against
    /// inline source.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the positive control for the guard above, and it exists because that
    /// guard was silently weak.</b> Its bare-name branch originally iterated
    /// <c>MemberAccessExpressionSyntax</c> and read <c>access.Name</c>, so it could
    /// only match <c>Foo.Bar</c>. It therefore never saw <c>new Foo()</c>, a
    /// <c>Foo</c>-typed parameter, or a cast — and the real
    /// <c>RefreshSheetCommand</c> constructs the factory with <c>new</c>. The
    /// production test passed only because it was looking in the wrong place, and the
    /// composition root had to be added to an allowlist when the branch was fixed.
    /// </para>
    /// <para>
    /// Without this test the same regression could return silently: a matcher that
    /// finds nothing looks identical to a matcher that is working.
    /// </para>
    /// </remarks>
    [Theory]
    // The shape that actually occurred in production and was missed.
    [InlineData("class C { void M() { var f = new ExcelSceneBuildRequestFactory(); } }")]
    // The other two ways a type is named without a dotted access.
    [InlineData("class C { void M(ISceneBuildRequestFactory f) { } }")]
    [InlineData("class C { void M(object o) { var f = (ISceneBuildRequestFactory)o; } }")]
    // A local of the concrete type.
    [InlineData("class C { void M() { ExcelSceneBuildRequestFactory f = null; } }")]
    public void The_bare_name_branch_detects_a_type_reference_that_is_not_a_member_access(string source)
    {
        SyntaxNode root = CSharpSyntaxTree.ParseText(source).GetRoot();

        Assert.True(
            MentionsIn(root, "ExcelSceneBuildRequestFactory")
                || MentionsIn(root, "ISceneBuildRequestFactory"),
            "the matcher missed a plain type reference in: " + source);
    }

    /// <summary>
    /// The counterweight: a comment or a string literal naming the type is NOT a
    /// mention, and neither is an unrelated name.
    /// </summary>
    /// <remarks>
    /// A guard that matched raw text would flag both of these, which is why the check
    /// is syntactic. Asserting it keeps the fix from being made by broadening to
    /// substrings.
    /// </remarks>
    [Fact]
    public void A_comment_or_string_naming_the_type_is_not_a_mention()
    {
        const string Commented = "class C { /* ExcelSceneBuildRequestFactory */ void M() { } }";
        const string InAString = "class C { string s = \"ExcelSceneBuildRequestFactory\"; void M() { } }";
        const string Unrelated = "class C { void M() { ShapeReconciler.Reconcile(); } }";

        Assert.False(MentionsIn(CSharpSyntaxTree.ParseText(Commented).GetRoot(), "ExcelSceneBuildRequestFactory"));
        Assert.False(MentionsIn(CSharpSyntaxTree.ParseText(InAString).GetRoot(), "ExcelSceneBuildRequestFactory"));
        Assert.False(MentionsIn(CSharpSyntaxTree.ParseText(Unrelated).GetRoot(), "ExcelSceneBuildRequestFactory"));
    }

    /// <summary>
    /// The orchestrator itself holds no geometry derivation of its own.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The orchestrator's job is ordering. The moment it also decides a preset, a
    /// plot width, or a date range, it has a second authority for an input D5 assigns
    /// to the factory, and the two are free to disagree.
    /// </para>
    /// <para>
    /// The rule is stated on the syntax a derivation would contain — a call to the
    /// plot resolver, the size-preset catalogue, or the time scale — rather than on
    /// any particular spelling, so renaming a local does not defeat it.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_orchestrator_derives_no_scene_geometry_of_its_own()
    {
        string source = File.ReadAllText(Path.Combine(FindRepositoryRoot(), OrchestratorPath));

        foreach (string forbidden in new[]
                 {
                     "PlotGeometryResolver",
                     "SizePresets",
                     "TimeScale.",
                     "PlotGeometryOutcome",
                 })
        {
            Assert.DoesNotContain(
                forbidden,
                source,
                StringComparison.Ordinal);
        }
    }

    /// <summary>The orchestrator reaches no COM type, so the whole pipeline is testable without Excel.</summary>
    [Fact]
    public void The_orchestrator_is_free_of_office_references()
    {
        string source = File.ReadAllText(Path.Combine(FindRepositoryRoot(), OrchestratorPath));

        Assert.DoesNotContain("Interop.Excel", source, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Microsoft.Office.Interop", source, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The ribbon command holds no pipeline: it names no validator, no reconciler, and
    /// no scene builder, so D1's "the command is not the orchestrator" is structural.
    /// </summary>
    [Fact]
    public void The_refresh_command_holds_no_pipeline_of_its_own()
    {
        string source = File.ReadAllText(
            Path.Combine(FindRepositoryRoot(), "src/GanttCreator.AddIn/RefreshSheetCommand.cs"));

        foreach (string forbidden in new[]
                 {
                     "GanttRowValidator",
                     "ShapeReconciler",
                     "SceneBuilder",
                     "DurationCalculator",
                     "ShapeBuilder",
                 })
        {
            Assert.DoesNotContain(forbidden, source, StringComparison.Ordinal);
        }
    }

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
    /// Whether the file <em>uses</em> the named member — invoked, or referenced as a
    /// method group — ignoring comments and string literals.
    /// </summary>
    /// <param name="relativePath">The repository-relative file path.</param>
    /// <param name="member">The member to look for: either a bare type name
    /// (<c>ExcelSceneBuildRequestFactory</c>) or a dotted member
    /// (<c>ShapeReconciler.Reconcile</c>).</param>
    /// <returns><see langword="true"/> when the member is used.</returns>
    /// <remarks>
    /// <para>
    /// This deliberately matches a reference, not only an invocation. The first
    /// version looked for an <see cref="InvocationExpressionSyntax"/> whose expression
    /// text ended with the member name, and a probe proved it vacuous: a file that
    /// merely named <c>ShapeReconciler.Reconcile</c> — as a method group, a delegate,
    /// or a log message about it — was reported clean. A guard that can be defeated by
    /// removing two parentheses is not a guard.
    /// </para>
    /// <para>
    /// Matching a plain identifier reference is still far better than a substring
    /// search, because a mention inside a comment or a string literal does not parse
    /// as an identifier. That is the evasion this shape does <em>not</em> permit, and
    /// the reason the check is syntactic at all.
    /// </para>
    /// </remarks>
    private static bool Mentions(string relativePath, string member)
    {
        string source = File.ReadAllText(Path.Combine(FindRepositoryRoot(), relativePath));
        return MentionsIn(CSharpSyntaxTree.ParseText(source).GetRoot(), member);
    }

    /// <summary>
    /// Whether the syntax tree <em>uses</em> the named member.
    /// </summary>
    /// <param name="root">The parsed source root.</param>
    /// <param name="member">
    /// The member to look for: either a bare type name
    /// (<c>ExcelSceneBuildRequestFactory</c>) or a dotted member
    /// (<c>ShapeReconciler.Reconcile</c>).
    /// </param>
    /// <returns><see langword="true"/> when the member is used.</returns>
    /// <remarks>
    /// <para>
    /// This deliberately matches a reference, not only an invocation. The first
    /// version looked for an <see cref="InvocationExpressionSyntax"/> whose expression
    /// text ended with the member name, and a probe proved it vacuous: a file that
    /// merely named <c>ShapeReconciler.Reconcile</c> — as a method group, a delegate,
    /// or a log message about it — was reported clean. A guard that can be defeated by
    /// removing two parentheses is not a guard.
    /// </para>
    /// <para>
    /// Matching a plain identifier reference is still far better than a substring
    /// search, because a mention inside a comment or a string literal does not parse
    /// as an identifier. That is the evasion this shape does <em>not</em> permit, and
    /// the reason the check is syntactic at all.
    /// </para>
    /// <para>
    /// <b>A bare name is matched on every identifier in the tree</b>, not only on
    /// member-access names. This branch used to iterate
    /// <see cref="MemberAccessExpressionSyntax"/> and read <c>access.Name</c>, so it
    /// could only ever find <c>Foo.Bar</c> and missed the three shapes that matter
    /// most for a type name: <c>new Foo()</c> (an
    /// <see cref="ObjectCreationExpressionSyntax"/>, whose type is not a member
    /// access), a parameter or local declared as <c>Foo</c> (an
    /// <see cref="IdentifierNameSyntax"/> in a type position), and a cast. The real
    /// <c>RefreshSheetCommand</c> constructs the factory with <c>new</c>, so the guard
    /// was reporting a composition root it claimed to forbid as clean — for the wrong
    /// reason.
    /// </para>
    /// </remarks>
    private static bool MentionsIn(SyntaxNode root, string member)
    {
        // A dotted name is matched on two real identifiers — the receiver type and the
        // member name — rather than on one dotted string. A bare name is matched as a
        // whole identifier.
        string[] parts = member.Split('.');
        var bare = parts.Length == 1;
        string typeName = bare ? string.Empty : parts[^2];
        string memberName = parts[^1];

        foreach (SyntaxNode node in root.DescendantNodes())
        {
            if (bare)
            {
                if (node is IdentifierNameSyntax identifier
                    && identifier.Identifier.ValueText == memberName)
                {
                    return true;
                }

                continue;
            }

            if (node is MemberAccessExpressionSyntax access
                && access.Name.Identifier.ValueText == memberName
                && access.Expression.ToString().EndsWith(typeName, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
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
