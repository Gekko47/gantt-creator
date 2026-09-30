using System.IO;

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
    /// No production file may apply the D2 subtraction inline. This is the specific
    /// duplication the row exists to prevent: the same formula written a second time
    /// somewhere that will drift from the first.
    /// </summary>
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

                var text = File.ReadAllText(file);
                if (text.Contains("WidthPt -", StringComparison.Ordinal)
                    || text.Contains("- TextPanelWidthPt", StringComparison.Ordinal)
                    || text.Contains("- textPanelWidthPt", StringComparison.Ordinal))
                {
                    offenders.Add($"{relative}: derives plot width by inline subtraction");
                }
            }
        }

        Assert.Empty(offenders);
    }

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
