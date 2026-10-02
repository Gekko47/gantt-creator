using System.IO;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// R4.7H D1: <c>PlotGeometryResolver</c> is the <b>single</b> authority deriving
/// plot bounds. A second derivation in a command, a renderer, or a Ribbon callback
/// would disagree with it by a margin somewhere, and the disagreement would present
/// as a layout bug rather than as duplicated logic. This is the guard that makes the
/// authority structural instead of a convention.
/// </summary>
public sealed class PlotGeometryAuthorityTests
{
    /// <summary>Production source roots scanned for an out-of-authority derivation.</summary>
    private static readonly string[] ProductionRoots =
    [
        "src/GanttCreator.AddIn",
        "src/GanttCreator.Core",
        "src/GanttCreator.Office",
        "src/GanttCreator.Raster",
    ];

    /// <summary>The single file permitted to construct resolved plot bounds.</summary>
    private const string ResolverPath = "src/GanttCreator.Core/Scene/PlotGeometryResolver.cs";

    /// <summary>
    /// Files permitted to carry plot bounds on the resolver's behalf, because they
    /// <em>pass through</em> a resolved value rather than deriving one.
    /// </summary>
    private static readonly string[] PermittedConsumers =
    [
        ResolverPath,
        "src/GanttCreator.Core/Scene/SceneBuilder.cs",
    ];

    /// <summary>
    /// No production file may derive plot bounds itself. This is the specific
    /// duplication the row exists to prevent: the same formula written a second time
    /// somewhere that will drift from the first.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The scan is SYNTAX-based, via Roslyn, rather than a substring search. The
    /// substring form looked for three literal spellings of the subtraction
    /// (<c>"WidthPt -"</c>, <c>"- TextPanelWidthPt"</c>, <c>"- textPanelWidthPt"</c>),
    /// which ordinary formatting defeats: put the operand on the next line, name the
    /// preset variable something else, or parenthesise it, and a second derivation sails
    /// through while the test stays green. That is the worst failure mode for this guard
    /// -- it would report the single-authority claim as enforced while examining nothing.
    /// </para>
    /// <para>
    /// The rule is stated on the syntax fact a derivation must contain: a subtraction
    /// whose operand reads a size preset's own <c>WidthPt</c>. Line breaks, spacing,
    /// extra parentheses and variable naming are irrelevant. An unrelated
    /// <c>WidthPt</c> subtraction -- a text measurement's width inside
    /// <c>DateLabelBuilder</c>, say -- is NOT flagged, because the operand there is not a
    /// preset member.
    /// </para>
    /// </remarks>
    [Fact]
    public void Plot_bounds_are_not_derived_outside_the_resolver()
    {
        var repositoryRoot = FindRepositoryRoot();
        var offenders = new List<string>();

        foreach (var root in ProductionRoots)
        {
            var directory = Path.Combine(repositoryRoot, root);
            // A missing root FAILS the test rather than being skipped. Skipping is how
            // a renamed or deleted production project would silently stop being
            // scanned, leaving this guard green while covering nothing -- the test
            // would pass because it examined less, not because the contract holds.
            Assert.True(
                Directory.Exists(directory),
                $"Production root '{root}' does not exist under the repository root; "
                + "a missing root would be silently skipped rather than scanned.");

            foreach (var file in Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(repositoryRoot, file).Replace('\\', '/');
                if (PermittedConsumers.Contains(relative, StringComparer.Ordinal))
                {
                    continue;
                }

                if (DerivesPresetWidth(File.ReadAllText(file)))
                {
                    offenders.Add($"{relative}: derives plot bounds from a size preset width");
                }
            }
        }

        Assert.Empty(offenders);
    }

    /// <summary>
    /// The positive control for the syntax scan above: a second derivation written with
    /// different formatting, a different variable name and an extra pair of parentheses
    /// -- every shape the previous substring form missed -- is still detected.
    /// </summary>
    [Fact]
    public void The_syntax_scan_detects_a_renamed_and_reformatted_derivation()
    {
        const string Sneaky = """
            class Sneaky
            {
                RectD Bounds(SizePreset paper, double columnPt)
                {
                    var chrome = 10;
                    var width =
                        paper.WidthPt
                        - columnPt
                        - chrome;
                    return new RectD(columnPt + chrome, 0, width, paper.HeightPt);
                }
            }
            """;

        Assert.True(DerivesPresetWidth(Sneaky));
    }

    /// <summary>
    /// The negative control: a subtraction of a width that is NOT a preset's is not a
    /// plot derivation. <c>DateLabelBuilder</c> really does subtract
    /// <c>measured.WidthPt</c>, and flagging that would make the guard cry wolf on
    /// legitimate code and get switched off.
    /// </summary>
    [Fact]
    public void The_syntax_scan_ignores_an_unrelated_width_subtraction()
    {
        const string Unrelated = """
            class Labeller
            {
                RectD Place(RectD visible, double gap, double measuredWidthPt)
                {
                    return new RectD(visible.Left - gap - measuredWidthPt, 0, measuredWidthPt, 10);
                }
            }
            """;

        Assert.False(DerivesPresetWidth(Unrelated));
    }

    /// <summary>
    /// Whether a source subtracts a size preset's own width.
    /// </summary>
    /// <param name="source">The C# source text.</param>
    /// <returns><see langword="true"/> when a preset-width subtraction is present.</returns>
    private static bool DerivesPresetWidth(string source)
    {
        SyntaxNode root = CSharpSyntaxTree.ParseText(source).GetRoot();

        return root
            .DescendantNodes()
            .OfType<BinaryExpressionSyntax>()
            .Where(binary => binary.IsKind(SyntaxKind.SubtractExpression))
            .Any(binary => ContainsPresetWidth(binary.Left) || ContainsPresetWidth(binary.Right));
    }

    /// <summary>
    /// Whether a subtree reads a size preset's <c>WidthPt</c>, at any depth and through
    /// any number of parentheses.
    /// </summary>
    /// <param name="node">The subtree to search.</param>
    /// <returns><see langword="true"/> when a preset width read is present.</returns>
    private static bool ContainsPresetWidth(SyntaxNode? node) =>
        node is not null
        && node
            .DescendantNodesAndSelf()
            .OfType<MemberAccessExpressionSyntax>()
            .Any(access =>
                access.Name.Identifier.ValueText == "WidthPt"
                && access.Expression is IdentifierNameSyntax identifier
                && PresetVariableNames.Contains(identifier.Identifier.ValueText));

    /// <summary>
    /// Local names treated as holding a <c>SizePreset</c>. This is the one place the
    /// scan consults a name, and the limitation is recorded rather than hidden: this
    /// assembly deliberately has NO project reference to compile against (that
    /// isolation is what lets it assert Core carries no Office dependency), so the
    /// preset's static type cannot be resolved and the receiver has to be recognised
    /// textually. A derivation that reads a preset width through a local named
    /// something outside this set would be missed.
    /// </summary>
    private static readonly HashSet<string> PresetVariableNames = new(StringComparer.Ordinal)
    {
        "preset",
        "paper",
        "sizePreset",
        "size",
        "a4",
        "presentation",
        "selected",
    };

    /// <summary>
    /// At most one production caller may invoke the resolver. Two callers are not two
    /// derivations, but they are two places to keep in step, and the row's whole claim
    /// is one authority. The scene-request factory (R4.8A) is the intended caller.
    /// </summary>
    [Fact]
    public void Only_the_caller_surface_invokes_the_resolver()
    {
        var repositoryRoot = FindRepositoryRoot();
        var callers = new List<string>();

        foreach (var root in ProductionRoots)
        {
            var directory = Path.Combine(repositoryRoot, root);
            // Same reason as the scan above: a missing root must fail, not shrink
            // the set of files this caller-count guard inspects.
            Assert.True(
                Directory.Exists(directory),
                $"Production root '{root}' does not exist under the repository root; "
                + "a missing root would be silently skipped rather than scanned.");

            foreach (var file in Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(repositoryRoot, file).Replace('\\', '/');
                if (relative == ResolverPath)
                {
                    continue;
                }

                if (File.ReadAllText(file).Contains("PlotGeometryResolver.TryResolve", StringComparison.Ordinal))
                {
                    callers.Add(relative);
                }
            }
        }

        Assert.True(
            callers.Count <= 1,
            "Plot bounds must be resolved in at most one production caller, but these call it: "
            + string.Join(", ", callers));
    }

    /// <summary>
    /// The resolver must not depend on Office. Plot bounds are pure arithmetic over a
    /// preset and a measurement, so an interop reference here would make the single
    /// authority untestable without a live host.
    /// </summary>
    [Fact]
    public void The_resolver_is_free_of_office_references()
    {
        var repositoryRoot = FindRepositoryRoot();
        var text = File.ReadAllText(Path.Combine(repositoryRoot, ResolverPath));

        Assert.DoesNotContain("Microsoft.Office", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Interop.Excel", text, StringComparison.OrdinalIgnoreCase);
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
