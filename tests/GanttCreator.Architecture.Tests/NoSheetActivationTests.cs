using System.IO;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// R4.8 D2/D3: a refresh never activates another worksheet and never adds one.
/// </summary>
/// <remarks>
/// <para>
/// D3's "zero activate-operations" is stated about the <em>refresh</em> path, and
/// that qualifier carries real weight rather than being a way to soften the rule.
/// <c>ExcelInsertedRowSelector</c> does call <c>worksheet.Activate()</c>, and it
/// is correct that it does: <c>Range.Select</c> only works on the active sheet, so
/// the AddRow command must activate the Gantt worksheet to put the cursor on the row
/// it just inserted. A blanket ban on <c>Activate</c> would have flagged working,
/// deliberate code and invited someone to "fix" it by breaking cursor placement.
/// </para>
/// <para>
/// So the rule is an allowlist rather than a ban, and the allowlist is deliberately
/// explicit about the one file that is exempt and why. A second exemption means the
/// rule was wrong, not that the guard should grow.
/// </para>
/// <para>
/// The scan is SYNTAX-based via Roslyn for the same reason
/// <c>PlotGeometryAuthorityTests</c> is: a substring search for <c>".Activate("</c>
/// is defeated by whitespace and line breaks in ways that would leave the guard
/// reporting the contract as enforced while examining nothing.
/// </para>
/// </remarks>
public sealed class NoSheetActivationTests
{
    /// <summary>Production source roots scanned for a sheet activation.</summary>
    private static readonly string[] ProductionRoots =
    [
        "src/GanttCreator.AddIn",
        "src/GanttCreator.Office",
    ];

    /// <summary>
    /// Files permitted to activate a worksheet, with the reason on record.
    /// </summary>
    /// <remarks>
    /// <c>ExcelInsertedRowSelector</c> is the AddRow command's cursor placement, not
    /// a render. It is exempt because <c>Range.Select</c> is documented to require an
    /// active sheet, and removing the activation would leave the user with no
    /// feedback after adding a row. It is exempt from the <em>refresh</em> rule
    /// because AddRow is not a refresh; it is not an exemption from "select what you
    /// just changed".
    /// </remarks>
    private static readonly string[] PermittedActivators =
    [
        "src/GanttCreator.Office/ExcelInsertedRowSelector.cs",
    ];

    /// <summary>
    /// No production file outside <see cref="PermittedActivators"/> activates a
    /// worksheet.
    /// </summary>
    /// <remarks>
    /// This is the guard that makes "Refresh never switches the active sheet" a
    /// structural property rather than an intention. A refresh that called
    /// <c>Activate</c> to make some COM operation legal would silently move the
    /// user's cursor and scroll their view — a small, easy-to-miss regression that
    /// no functional test would catch, because the chart would still render.
    /// </remarks>
    [Fact]
    public void No_refresh_path_activates_a_worksheet()
    {
        var repositoryRoot = FindRepositoryRoot();
        var offenders = new List<string>();

        foreach (var root in ProductionRoots)
        {
            var directory = Path.Combine(repositoryRoot, root);
            // A missing root FAILS rather than being skipped. Skipping is how a
            // renamed or deleted project would silently stop being scanned, leaving
            // the guard green while covering nothing.
            Assert.True(
                Directory.Exists(directory),
                $"Production root '{root}' does not exist under the repository root; "
                + "a missing root would be silently skipped rather than scanned.");

            foreach (var file in Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(repositoryRoot, file).Replace('\\', '/');
                if (PermittedActivators.Contains(relative, StringComparer.Ordinal))
                {
                    continue;
                }

                if (InvokesActivate(file))
                {
                    offenders.Add(relative);
                }
            }
        }

        Assert.True(
            offenders.Count == 0,
            "A refresh path must never activate a worksheet, but these call Activate(): "
            + string.Join(", ", offenders)
            + ". AddRow's cursor placement (ExcelInsertedRowSelector) is the only "
            + "permitted exemption; if another file needs one, the rule is wrong and "
            + "should be re-decided rather than the allowlist extended.");
    }

    /// <summary>
    /// D2: no production file creates or deletes a worksheet.
    /// </summary>
    /// <remarks>
    /// The add-in owns exactly one helper sheet, <c>_GanttCreatorConfig</c>, and it
    /// is created by the initialiser alone. A refresh that added or removed a sheet
    /// would change the user's workbook structure as a side effect of pressing
    /// Refresh, and a "cleanup" that deleted an unowned sheet would take the user's
    /// own data with it. This is the row the test-strategy rule "no extra worksheet
    /// creation" refers to, asserted structurally across every refresh path rather
    /// than per adapter.
    /// </remarks>
    [Fact]
    public void Only_the_initialiser_creates_a_worksheet()
    {
        var repositoryRoot = FindRepositoryRoot();
        var offenders = new List<string>();

        foreach (var root in ProductionRoots)
        {
            var directory = Path.Combine(repositoryRoot, root);
            foreach (var file in Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(repositoryRoot, file).Replace('\\', '/');
                if (relative == "src/GanttCreator.Office/ExcelWorkbookInitialiser.cs")
                {
                    continue;
                }

                if (InvokesWorksheetAdditionOrRemoval(file))
                {
                    offenders.Add(relative);
                }
            }
        }

        Assert.True(
            offenders.Count == 0,
            "Only ExcelWorkbookInitialiser may create or delete a worksheet, but these call "
            + "Worksheets.Add or Worksheets.Delete: " + string.Join(", ", offenders));
    }

    /// <summary>
    /// Determines whether the file invokes <c>Activate</c> on something that is
    /// actually a worksheet.
    /// </summary>
    /// <param name="file">The file to inspect.</param>
    /// <returns><see langword="true"/> when a worksheet is activated.</returns>
    /// <remarks>
    /// <para>
    /// Matching the member name alone was too coarse, and the suite caught it on the
    /// first run: <c>AddInHost.cs</c> calls <c>RibbonStateService.Instance.Activate()</c>,
    /// which enables the Ribbon and has nothing to do with Excel's object model. A
    /// name-only match would have demanded an exemption for it, and an exemption list
    /// containing "the ribbon service" would have told a later reader the rule was
    /// about something it is not.
    /// </para>
    /// <para>
    /// The receiver is therefore required to look like a worksheet. This cannot be
    /// fully type-checked without a semantic model over the interop assembly, so the
    /// rule is syntactic and stated as such: a worksheet activation written against
    /// an unrecognised receiver name would evade it. The alternatives were worse —
    /// a substring scan is defeated by formatting, and a full compilation against
    /// the interop reference is not available in this test project.
    /// </para>
    /// </remarks>
    private static bool InvokesActivate(string file)
    {
        foreach (SyntaxNode node in Descendants(file))
        {
            // The null-conditional form is matched as well as the plain one, and the
            // probe is why: a first version handled only MemberAccessExpression, so
            // `worksheet.Activate()` was caught and `worksheet?.Activate()` sailed
            // straight through — the guard was green while a violating file sat in
            // the tree. An evasion that costs one character is not one worth
            // shipping a guard that can be defeated by.
            //
            // The two forms have genuinely different shapes. `a.B()` is an
            // InvocationExpression over a MemberAccessExpression. `a?.B()` is a
            // ConditionalAccessExpression whose WhenNotNull is an InvocationExpression
            // over a *MemberBinding*Expression, with the receiver held on the
            // ConditionalAccess itself. Matching only the first shape silently
            // misses the second, which the second probe run demonstrated.
            switch (node)
            {
                case InvocationExpressionSyntax { Expression: MemberAccessExpressionSyntax access }
                    when access.Name.Identifier.ValueText == "Activate"
                        && IsWorksheetReceiver(access.Expression.ToString()):
                    return true;

                case ConditionalAccessExpressionSyntax conditional
                    when conditional.WhenNotNull is InvocationExpressionSyntax
                    {
                        Expression: MemberBindingExpressionSyntax binding
                    }
                        && binding.Name.Identifier.ValueText == "Activate"
                        && IsWorksheetReceiver(conditional.Expression.ToString()):
                    return true;

                default:
                    break;
            }
        }

        return false;
    }

    /// <summary>
    /// Determines whether the receiver of the call names a worksheet.
    /// </summary>
    /// <param name="receiver">The receiver expression's text.</param>
    /// <returns><see langword="true"/> when the receiver is worksheet-shaped.</returns>
    private static bool IsWorksheetReceiver(string receiver) =>
        receiver.EndsWith("worksheet", StringComparison.OrdinalIgnoreCase)
        || receiver.EndsWith("sheet", StringComparison.OrdinalIgnoreCase)
        || receiver.EndsWith("Worksheets", StringComparison.Ordinal)
        || receiver.EndsWith("Sheets", StringComparison.Ordinal);

    /// <summary>
    /// Determines whether the file adds or deletes a worksheet through
    /// <c>Worksheets.Add</c> or <c>Worksheets.Delete</c>.
    /// </summary>
    /// <param name="file">The file to inspect.</param>
    /// <returns><see langword="true"/> when a worksheet is added or removed.</returns>
    private static bool InvokesWorksheetAdditionOrRemoval(string file)
    {
        foreach (var node in Descendants(file))
        {
            if (node is InvocationExpressionSyntax invocation
                && invocation.Expression is MemberAccessExpressionSyntax access
                && access.Name.Identifier.ValueText is "Add" or "Delete")
            {
                // Only on a receiver that is recognisably a Worksheets collection.
                // A plain "Delete" on a shape or a row is ordinary mutation and must
                // not be flagged, or this guard would fire on the entire renderer.
                var receiver = access.Expression.ToString();
                if (receiver.EndsWith("Worksheets", StringComparison.Ordinal)
                    || receiver.EndsWith("Sheets", StringComparison.Ordinal))
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>Parses the file and yields its descendant syntax nodes.</summary>
    /// <param name="file">The file to parse.</param>
    /// <returns>Every node in the file's syntax tree.</returns>
    private static IEnumerable<SyntaxNode> Descendants(string file)
    {
        var tree = CSharpSyntaxTree.ParseText(File.ReadAllText(file));
        return tree.GetRoot().DescendantNodes();
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
