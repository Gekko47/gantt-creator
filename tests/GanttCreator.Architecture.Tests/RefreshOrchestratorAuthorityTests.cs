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
    /// <c>RefreshSheetCommand.cs</c> is the composition root, and it is handled by
    /// <see cref="ComposesTheFactoryWithoutDrivingIt"/> rather than by this list —
    /// see <see cref="CompositionRootPath"/> for why a blanket exemption is too much.
    /// </description></item>
    /// </list>
    /// </remarks>
    private static readonly string[] FactoryDeclarationPaths =
    [
        "src/GanttCreator.Office/ISceneBuildRequestFactory.cs",
        "src/GanttCreator.Office/ExcelSceneBuildRequestFactory.cs",
    ];

    /// <summary>
    /// The composition root: the one file permitted to name the factory without
    /// being the orchestrator.
    /// </summary>
    /// <remarks>
    /// <para>
    /// D1 requires the Ribbon command to construct the live orchestrator and its
    /// collaborators, so it must construct the factory. It must not <em>drive</em> it:
    /// a Ribbon callback that called <c>Create</c> would decide the preset, the plot
    /// bounds, and the range itself, which is the second authority D5 exists to
    /// prevent, and it would sit in the file people edit when adding a feature.
    /// </para>
    /// <para>
    /// A whole-file exemption could not tell those two apart, so the composition root
    /// gets its own check: every mention must be the type of an object creation, and
    /// no member of the factory may be invoked.
    /// </para>
    /// </remarks>
    private const string CompositionRootPath = "src/GanttCreator.AddIn/RefreshSheetCommand.cs";

    /// <summary>The factory members that would make the command a second pipeline.</summary>
    private static readonly string[] FactoryMembers =
    [
        "Create",
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
                // The composition root is allowed to CONSTRUCT the factory and nothing
                // more, so it is checked rather than skipped: a blanket exemption would
                // also permit `factory.Create(...)`, which is the second pipeline.
                if (file == CompositionRootPath)
                {
                    if (!ComposesTheFactoryWithoutDrivingIt(file))
                    {
                        callers.Add(file);
                    }

                    continue;
                }

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
    /// The composition root may CONSTRUCT the factory and may not DRIVE it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the positive test for the narrowed exemption.</b> The composition
    /// root used to be skipped wholesale, which permitted it to call
    /// <c>Create</c> — a second authority for the plot bounds, the preset, and the
    /// range, in the file people edit when adding a feature. Constructing the factory
    /// is required by D1 and stays permitted; calling it does not.
    /// </para>
    /// <para>
    /// Both directions are asserted, so a check that simply reported "false" would
    /// fail the first case and one that reported "true" would fail the others.
    /// </para>
    /// </remarks>
    [Theory]
    // Permitted: construction, including the real command's nested collaborator.
    [InlineData("class C { void M() { var f = new ExcelSceneBuildRequestFactory(); } }", true)]
    [InlineData("class C { void M() { var f = new ISceneBuildRequestFactory(); } }", true)]
    // Refused: driving the factory is the orchestrator's job alone.
    [InlineData("class C { void M() { var f = new ExcelSceneBuildRequestFactory(); f.Create(null, null, null, null); } }", false)]
    [InlineData("class C { void M() { var f = new ISceneBuildRequestFactory(); f.Create(null, null, null, null); } }", false)]
    [InlineData("class C { void M() { ISceneBuildRequestFactory f; f.Create(null, null, null, null); } }", false)]
    public void The_composition_root_may_construct_the_factory_but_may_not_drive_it(string source, bool expected)
    {
        SyntaxNode root = CSharpSyntaxTree.ParseText(source).GetRoot();

        Assert.Equal(expected, ComposesFactoryWithoutDrivingIt(root));
    }

    /// <summary>
    /// The real composition root passes the narrowed check, so the guard above is not
    /// green because it examines nothing.
    /// </summary>
    [Fact]
    public void The_real_composition_root_only_constructs_the_factory()
    {
        Assert.True(
            ComposesTheFactoryWithoutDrivingIt(CompositionRootPath),
            CompositionRootPath + " names the factory somewhere other than an object creation, "
            + "or invokes a factory member. Constructing it is D1's composition root; driving it is not.");
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
    /// <summary>
    /// Whether the composition root may name the factory: every mention must be the
    /// type of an <see cref="ObjectCreationExpressionSyntax"/>, and no factory member
    /// may be invoked.
    /// </summary>
    /// <param name="relativePath">The repository-relative file path.</param>
    /// <returns><see langword="true"/> when the file only constructs the factory.</returns>
    /// <remarks>
    /// <para>
    /// This is a separate check rather than one more allowlisted path because the two
    /// permissions are different. Constructing the factory is required by D1 and is a
    /// declaration; calling <c>Create</c> on it is the scene-input authority D5 gives
    /// to the orchestrator alone, and a whole-file exemption could not tell them apart.
    /// </para>
    /// <para>
    /// A mention is treated as a construction when the identifier is the type of an
    /// <c>ObjectCreationExpression</c>, or the type of a generic/qualified one. A
    /// mention anywhere else — a field, a parameter, a local declaration, a cast — is
    /// a reference the composition root has no reason to hold.
    /// </para>
    /// </remarks>
    private static bool ComposesTheFactoryWithoutDrivingIt(string relativePath)
    {
        string source = File.ReadAllText(Path.Combine(FindRepositoryRoot(), relativePath));
        return ComposesFactoryWithoutDrivingIt(CSharpSyntaxTree.ParseText(source).GetRoot());
    }

    /// <summary>The syntax-level form of the composition-root check, so a test can drive it inline.</summary>
    /// <param name="root">The parsed source root.</param>
    /// <returns><see langword="true"/> when the tree only constructs the factory.</returns>
    private static bool ComposesFactoryWithoutDrivingIt(SyntaxNode root)
    {
        HashSet<string> factoryTypes =
        [
            "ExcelSceneBuildRequestFactory",
            "ISceneBuildRequestFactory",
        ];

        foreach (SyntaxNode node in root.DescendantNodes())
        {
            if (node is not IdentifierNameSyntax identifier
                || !factoryTypes.Contains(identifier.Identifier.ValueText))
            {
                continue;
            }

            if (!IsCreationType(identifier))
            {
                return false;
            }
        }

        // A local, parameter, or field of the factory type is how a call would
        // actually be written: `f.Create(...)`, not `ISceneBuildRequestFactory.Create`.
        // The receivers are therefore resolved to their declared types rather than
        // matched as dotted text, which a one-letter local defeats.
        HashSet<string> factoryLocals = LocalsOfType(root, factoryTypes);

        foreach (SyntaxNode node in root.DescendantNodes())
        {
            if (node is InvocationExpressionSyntax invocation
                && invocation.Expression is MemberAccessExpressionSyntax access
                && FactoryMembers.Contains(access.Name.Identifier.ValueText)
                && IsFactoryReceiver(access.Expression, factoryTypes, factoryLocals))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Names of every local, parameter, and field declared with one of the types.</summary>
    /// <param name="root">The parsed source root.</param>
    /// <param name="types">The type names to look for.</param>
    /// <returns>The declared names.</returns>
    private static HashSet<string> LocalsOfType(SyntaxNode root, HashSet<string> types)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);

        foreach (SyntaxNode node in root.DescendantNodes())
        {
            switch (node)
            {
                case VariableDeclaratorSyntax declarator
                    when declarator.Parent is VariableDeclarationSyntax declaration
                        && (MentionsType(declaration.Type, types) || InitialisesFactory(declarator, types)):
                    _ = names.Add(declarator.Identifier.ValueText);
                    break;

                case ParameterSyntax parameter when MentionsType(parameter.Type!, types):
                    _ = names.Add(parameter.Identifier.ValueText);
                    break;

                case FieldDeclarationSyntax field when MentionsType(field.Declaration.Type, types):
                    foreach (VariableDeclaratorSyntax declarator in field.Declaration.Variables)
                    {
                        _ = names.Add(declarator.Identifier.ValueText);
                    }

                    break;

                default:
                    break;
            }
        }

        return names;
    }

    /// <summary>Whether the declarator's initialiser constructs or casts one of the types.</summary>
    /// <param name="declarator">The variable declarator.</param>
    /// <param name="types">The factory type names.</param>
    /// <returns><see langword="true"/> when the initialiser yields a factory.</returns>
    /// <remarks>
    /// <c>var</c> hides the type from the declaration, and the real command composes
    /// through an object-creation argument rather than a local — but a <c>var</c> local
    /// holding a factory is the ordinary way a caller would end up invoking
    /// <c>Create</c>, so the initialiser is inspected as well.
    /// </remarks>
    private static bool InitialisesFactory(VariableDeclaratorSyntax declarator, HashSet<string> types) =>
        declarator.Initializer is { } initializer
        && initializer.DescendantNodesAndSelf()
            .OfType<TypeSyntax>()
            .Any(type => MentionsType(type, types));

    /// <summary>Whether the given type syntax names one of the types.</summary>
    /// <param name="type">The declared type syntax.</param>
    /// <param name="types">The type names to look for.</param>
    /// <returns><see langword="true"/> when the declaration names one of them.</returns>
    private static bool MentionsType(TypeSyntax type, HashSet<string> types) =>
        type.DescendantNodesAndSelf()
            .OfType<IdentifierNameSyntax>()
            .Any(identifier => types.Contains(identifier.Identifier.ValueText));

    /// <summary>Whether the receiver of a member access is a factory instance.</summary>
    /// <param name="receiver">The receiver expression.</param>
    /// <param name="types">The factory type names.</param>
    /// <param name="locals">Names known to hold a factory instance.</param>
    /// <returns><see langword="true"/> when the receiver is a factory.</returns>
    private static bool IsFactoryReceiver(
        ExpressionSyntax receiver,
        HashSet<string> types,
        HashSet<string> locals) =>
        receiver is IdentifierNameSyntax identifier && locals.Contains(identifier.Identifier.ValueText)
        || types.Contains(receiver.ToString());

    /// <summary>Whether the identifier is the type of an object creation.</summary>
    /// <param name="identifier">The identifier naming a factory type.</param>
    /// <returns><see langword="true"/> when it appears as a constructed type.</returns>
    private static bool IsCreationType(IdentifierNameSyntax identifier) =>
        identifier.Parent is ObjectCreationExpressionSyntax creation
        && creation.Type == identifier;

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
