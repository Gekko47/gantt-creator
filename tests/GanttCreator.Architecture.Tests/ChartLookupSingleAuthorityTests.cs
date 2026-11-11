using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// R5.17e: one predicate owns the "is this the Gantt data table?" comparison.
/// </summary>
/// <remarks>
/// <para>
/// Before this guard the comparison was written out inline in twenty-odd Office
/// adapters, and the copies had drifted into two different comparison strengths.
/// A substring search for the schema name would not have caught that drift,
/// because every copy contained the same text; the guard therefore has to find
/// the <em>comparison</em>, which is why the scan is syntax-based via Roslyn in
/// the same shape as <c>NoSheetActivationTests</c>.
/// </para>
/// <para>
/// The guard is deliberately narrow: it flags a <c>string.Equals</c> call that
/// mentions <see cref="Core.GanttTableSchema"/>'s table name. Writing the name
/// into a table (<c>table.Name = GanttTableSchema.TableName</c>), passing it to a
/// lookup helper, or mentioning it in a log message is not a second authority and
/// is left alone.
/// </para>
/// </remarks>
public sealed class ChartLookupSingleAuthorityTests
{
    /// <summary>The one production file allowed to compare the table name.</summary>
    private const string ContractPath = "src/GanttCreator.Core/GanttChartContract.cs";

    /// <summary>The text that marks a comparison as a table-identity lookup.</summary>
    private const string SchemaNameText = "GanttTableSchema.TableName";

    /// <summary>
    /// No production file outside <see cref="ContractPath"/> compares a table name
    /// against the schema's table name.
    /// </summary>
    [Fact]
    public void Only_the_chart_contract_compares_a_table_name_to_the_schema()
    {
        string root = FindRepositoryRoot();
        List<string> offenders = [];

        foreach (string file in Directory.EnumerateFiles(
                     Path.Combine(root, "src"), "*.cs", SearchOption.AllDirectories))
        {
            string relative = Path.GetRelativePath(root, file).Replace('\\', '/');
            if (string.Equals(relative, ContractPath, StringComparison.Ordinal))
            {
                continue;
            }

            foreach (InvocationExpressionSyntax invocation in Nodes(File.ReadAllText(file))
                         .OfType<InvocationExpressionSyntax>())
            {
                if (!ComparesSchemaTableName(invocation))
                {
                    continue;
                }

                int line = invocation.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
                offenders.Add($"{relative}:{line}");
            }
        }

        Assert.Empty(offenders);
    }

    /// <summary>
    /// Positive control: the detector fires on a stub that re-inlines the
    /// comparison, so a green main test cannot mean "the scan examined nothing".
    /// </summary>
    [Fact]
    public void The_guard_detects_an_inline_comparison_in_a_stub()
    {
        const string stub = """
            static class Sample
            {
                internal static bool IsGantt(string? name) =>
                    string.Equals(name, GanttTableSchema.TableName, StringComparison.Ordinal);
            }
            """;

        Assert.Contains(Nodes(stub).OfType<InvocationExpressionSyntax>(), ComparesSchemaTableName);
    }

    /// <summary>
    /// Negative control: a write of the schema name is not a lookup and must stay
    /// legal, otherwise the guard would force writers through a predicate that
    /// does not fit them.
    /// </summary>
    [Fact]
    public void The_guard_allows_writing_the_schema_name()
    {
        const string stub = """
            static class Sample
            {
                internal static void Name(object table) =>
                    table.Name = GanttTableSchema.TableName;
            }
            """;

        Assert.DoesNotContain(Nodes(stub).OfType<InvocationExpressionSyntax>(), ComparesSchemaTableName);
    }

    /// <summary>Whether the invocation is a <c>string.Equals</c> mentioning the table name.</summary>
    /// <param name="invocation">The invocation to inspect.</param>
    /// <returns><see langword="true"/> when both conditions hold.</returns>
    private static bool ComparesSchemaTableName(InvocationExpressionSyntax invocation) =>
        IsStringEquals(invocation)
        && invocation.ArgumentList.Arguments.Any(argument =>
            argument.Expression.ToString().Contains(SchemaNameText, StringComparison.Ordinal));

    /// <summary>Whether the invocation's receiver is the <c>string</c> type.</summary>
    /// <param name="invocation">The invocation to inspect.</param>
    /// <returns><see langword="true"/> for a static call on <c>string</c>.</returns>
    private static bool IsStringEquals(InvocationExpressionSyntax invocation) =>
        invocation.Expression is MemberAccessExpressionSyntax access
        && string.Equals(access.Expression.ToString(), "string", StringComparison.Ordinal)
        && string.Equals(access.Name.Identifier.ValueText, "Equals", StringComparison.Ordinal);

    /// <summary>Parses source text and yields every node in its syntax tree.</summary>
    /// <param name="text">The source text.</param>
    /// <returns>Every descendant node.</returns>
    private static IEnumerable<SyntaxNode> Nodes(string text) =>
        CSharpSyntaxTree.ParseText(text).GetRoot().DescendantNodes();

    /// <summary>Walks up from the test binaries to the repository root.</summary>
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
