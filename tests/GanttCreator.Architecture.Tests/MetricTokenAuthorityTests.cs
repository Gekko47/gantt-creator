using System.IO;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// The token catalogue is the <b>single authority</b> for a metric default and for a
/// setting key. A production file that reads a metric from the settings map, or that
/// writes its own numeric fallback, is a second authority — and it is one that fails
/// silently, because the wrong number still renders a chart.
/// </summary>
/// <remarks>
/// <para>
/// This guard exists because of three defects found in a senior QA review of R4, all
/// of which passed a fully green suite:
/// </para>
/// <list type="number">
/// <item>The scene-request factory read every layout metric with
/// <c>ReadDouble(settings, tokenName, fallback)</c>, but the settings map carries only
/// the approved <em>setting</em> keys and no metric name is one of them — so every
/// lookup missed and the fallback was returned. Six of eight metrics disagreed with
/// the catalogue.</item>
/// <item>One metric was read under the key <c>DelineatorStackGapPt</c>, which is not a
/// token at all; the token is <c>StackGapPt</c>.</item>
/// <item>The orchestrator passed literal row heights <c>(15, 6, 6)</c> against
/// catalogue defaults of <c>18 / 18 / 9</c>.</item>
/// </list>
/// <para>
/// The scan is SYNTAX-based via Roslyn rather than a substring search, on the same
/// reasoning as <see cref="PlotGeometryAuthorityTests"/>: a substring form is defeated
/// by ordinary formatting and variable naming, and a guard that can be defeated while
/// reporting itself as enforcing the rule is worse than no guard.
/// </para>
/// </remarks>
public sealed class MetricTokenAuthorityTests
{
    /// <summary>Production source roots scanned for an out-of-authority metric read.</summary>
    private static readonly string[] ProductionRoots =
    [
        "src/GanttCreator.AddIn",
        "src/GanttCreator.Core",
        "src/GanttCreator.Office",
        "src/GanttCreator.Raster",
    ];

    /// <summary>
    /// The helpers whose second argument names a settings key or a metric token. Only
    /// these are examined, so an unrelated <c>TryGetValue</c> elsewhere is not flagged.
    /// </summary>
    private static readonly string[] SettingsReaders =
    [
        "ReadString",
        "ReadBool",
        "ReadDouble",
    ];

    /// <summary>
    /// Every settings key the catalogue defines, read from the catalogue source itself.
    /// </summary>
    /// <param name="repositoryRoot">The repository root.</param>
    /// <returns>The approved setting keys.</returns>
    /// <remarks>
    /// <para>
    /// The keys are parsed out of <c>GanttCatalogues.cs</c> rather than taken from a
    /// compiled reference, because this project intentionally carries
    /// <em>no</em> <c>ProjectReference</c> to any GanttCreator assembly — its csproj
    /// says so. Adding one to make a guard compile would weaken a deliberate
    /// boundary, so the source is read instead.
    /// </para>
    /// <para>
    /// Reading the source also means the guard cannot drift from a stale binary: it
    /// checks the same text a developer would edit.
    /// </para>
    /// </remarks>
    private static HashSet<string> ApprovedSettingKeys(string repositoryRoot)
    {
        var catalogue = Path.Combine(
            repositoryRoot, "src", "GanttCreator.Core", "GanttCatalogues.cs");
        Assert.True(File.Exists(catalogue), $"The catalogue source '{catalogue}' was not found.");

        HashSet<string> keys = new(StringComparer.Ordinal);

        // The `Settings` collection is initialised as a collection expression of
        // `new("Key", "value")` pairs. Every two-argument object creation inside it is
        // therefore exactly one setting key, and the first argument is the key.
        SyntaxNode root = CSharpSyntaxTree.ParseText(File.ReadAllText(catalogue)).GetRoot();
        foreach (CollectionExpressionSyntax collection in root.DescendantNodes().OfType<CollectionExpressionSyntax>())
        {
            string? owner = OwnerMemberName(collection);
            if (!string.Equals(owner, "Settings", StringComparison.Ordinal))
            {
                continue;
            }

            foreach (ExpressionElementSyntax element in collection.Elements)
            {
                // `new("Key", "value")` is a TARGET-TYPED creation, which parses as
                // ImplicitObjectCreationExpression rather than
                // ObjectCreationExpression. Reading only the latter finds nothing at
                // all -- a guard that silently matches zero and reports itself as
                // enforcing the rule.
                BaseArgumentListSyntax? arguments = element.Expression switch
                {
                    ImplicitObjectCreationExpressionSyntax implicitCreate => implicitCreate.ArgumentList,
                    ObjectCreationExpressionSyntax create => create.ArgumentList,
                    _ => null,
                };

                if (arguments is { Arguments.Count: >= 1 }
                    && arguments.Arguments[0].Expression is LiteralExpressionSyntax key
                    && key.IsKind(SyntaxKind.StringLiteralExpression))
                {
                    _ = keys.Add(key.Token.ValueText);
                }
            }
        }

        Assert.NotEmpty(keys);
        return keys;
    }

    /// <summary>
    /// Gets the name of the member a collection expression is initialising.
    /// </summary>
    /// <param name="collection">The collection expression.</param>
    /// <returns>The member name, or <see langword="null"/> when it cannot be determined.</returns>
    private static string? OwnerMemberName(CollectionExpressionSyntax collection)
    {
        // Walk up through the initialiser to the assignment's left-hand member access,
        // so `Settings { get; } = [ ... ]` yields "Settings".
        for (SyntaxNode? node = collection.Parent; node is not null; node = node.Parent)
        {
            // `Settings { get; } = [ ... ]` — the property declaration is the nearest
            // enclosing declaration, and its identifier is the collection's name.
            if (node is PropertyDeclarationSyntax property)
            {
                return property.Identifier.ValueText;
            }
        }

        return null;
    }

    /// <summary>
    /// No production file reads a metric through the settings map, and every settings
    /// key it reads is one the catalogue defines.
    /// </summary>
    [Fact]
    public void Settings_reads_name_keys_the_catalogue_defines()
    {
        var repositoryRoot = FindRepositoryRoot();
        HashSet<string> settings = ApprovedSettingKeys(repositoryRoot);
        var offenders = new List<string>();

        foreach (var root in ProductionRoots)
        {
            var directory = Path.Combine(repositoryRoot, root);
            Assert.True(
                Directory.Exists(directory),
                $"Production root '{root}' does not exist; a missing root would be silently skipped rather than scanned.");

            foreach (var file in Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(repositoryRoot, file).Replace('\\', '/');

                // SettingsKeyReads takes SOURCE, not a path. Passing the path would
                // parse the path string itself -- a scan that examined nothing while
                // reporting itself green, which is the exact failure mode this guard
                // is written to prevent.
                foreach ((string Key, int Line) in SettingsKeyReads(File.ReadAllText(file)))
                {
                    if (!settings.Contains(Key))
                    {
                        offenders.Add(
                            $"{relative}:{Line}: reads setting '{Key}', which GanttCatalogues.Settings does not define");
                    }
                }
            }
        }

        Assert.Empty(offenders);
    }

    /// <summary>
    /// The positive control: the key-extraction scan detects a read of a key the
    /// catalogue does not define, in whatever formatting the caller used.
    /// </summary>
    /// <remarks>
    /// Without this, a scan that silently matched nothing would report the rule as
    /// enforced while examining nothing — the failure mode the plot-geometry guard
    /// documents at length.
    /// </remarks>
    [Fact]
    public void The_key_scan_detects_an_unknown_setting_key()
    {
        const string Sneaky = """
            class Sneaky
            {
                string Read(IReadOnlyDictionary<string, string> settings) =>
                    ReadString(
                        settings,
                        "TotallyMadeUpKey");
            }
            """;

        List<(string Key, int Line)> reads = SettingsKeyReads(Sneaky);

        Assert.Contains(reads, read => read.Key == "TotallyMadeUpKey");
    }

    /// <summary>
    /// The positive control for the reader-name detection: a reader spelled with
    /// different whitespace and a multi-line call is still found.
    /// </summary>
    [Fact]
    public void The_scan_detects_a_reformatted_settings_read()
    {
        const string Sneaky = """
            class Sneaky
            {
                double Read(IReadOnlyDictionary<string, string> settings) =>
                    ReadDouble( settings ,
                        "AnotherUnknownKey" , 0 , 3 );
            }
            """;

        List<(string Key, int Line)> reads = SettingsKeyReads(Sneaky);

        Assert.Contains(reads, read => read.Key == "AnotherUnknownKey");
    }

    /// <summary>
    /// Finds every <c>ReadString</c>/<c>ReadBool</c>/<c>ReadDouble</c> call whose second
    /// argument is a string literal, and returns that literal with its line.
    /// </summary>
    /// <param name="code">The C# source to scan.</param>
    /// <returns>Each key the source reads, with the 1-based line it appears on.</returns>
    /// <remarks>
    /// Syntax-based on purpose. A substring scan for <c>"TitleBandHeightPt"</c> would
    /// equally match a <em>comparison</em> or a message string, so a file that never
    /// reads the key could satisfy the rule — the same defect the existing
    /// <c>AssignedMutationMembers</c> note records for <c>.RowHeight</c>.
    /// </remarks>
    private static List<(string Key, int Line)> SettingsKeyReads(string code)
    {
        SyntaxNode root = CSharpSyntaxTree.ParseText(code).GetRoot();
        var found = new List<(string Key, int Line)>();

        foreach (InvocationExpressionSyntax invocation in root.DescendantNodes().OfType<InvocationExpressionSyntax>())
        {
            string? name = Name(invocation);
            if (name is null || !SettingsReaders.Contains(name, StringComparer.Ordinal))
            {
                continue;
            }

            ArgumentListSyntax? arguments = invocation.ArgumentList;
            if (arguments is null || arguments.Arguments.Count < 2)
            {
                continue;
            }

            ExpressionSyntax key = arguments.Arguments[1].Expression;
            if (key is not LiteralExpressionSyntax literal || !literal.IsKind(SyntaxKind.StringLiteralExpression))
            {
                continue;
            }

            found.Add((
                literal.Token.ValueText,
                literal.GetLocation().GetLineSpan().StartLinePosition.Line + 1));
        }

        return found;
    }

    /// <summary>
    /// Gets the invoked method's simple name, whether written bare or as a member
    /// access, or <see langword="null"/> when the callee is not a simple identifier.
    /// </summary>
    /// <param name="invocation">The invocation to name.</param>
    /// <returns>The simple name, or null.</returns>
    private static string? Name(InvocationExpressionSyntax invocation)
    {
        ExpressionSyntax expression = invocation.Expression;
        SimpleNameSyntax? name = expression switch
        {
            MemberAccessExpressionSyntax member => member.Name,
            IdentifierNameSyntax identifier => identifier,
            MemberBindingExpressionSyntax binding => binding.Name,
            _ => null,
        };

        return name?.Identifier.ValueText;
    }

    /// <summary>
    /// Walks up from the test assembly to the repository root, which holds the
    /// <c>src/</c> roots.
    /// </summary>
    /// <returns>The repository root.</returns>
    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (Directory.Exists(Path.Combine(directory.FullName, "src", "GanttCreator.Core")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "The repository root (the ancestor containing src/GanttCreator.Core) was not found from "
            + AppContext.BaseDirectory + ".");
    }
}