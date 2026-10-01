using System.IO;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// R4.7I slice 1 D1: <c>GanttSheetLayout</c> is the <b>single</b> authority for the
/// live sheet's row layout and its plot anchor. This guard makes that structural
/// rather than a convention.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this guard exists.</b> Before ADR-0030, four files each built the plot
/// anchor's <c>refersTo</c> string independently — the initialiser, the integrity
/// checker, the plot-anchor repairer, and a test helper — each with its own copy of
/// the A1 column conversion and its own hardcoded <c>$1</c> row literal. ADR-0030
/// adds a reserved row above the table, which moves the header to row 2.
/// </para>
/// <para>
/// <b>The failure this prevents is not hypothetical.</b> Had the initialiser been
/// updated to write <c>$O$2</c> while the integrity checker still expected
/// <c>$O$1</c>, the integrity check would report <c>PlotAnchorDisagreement</c> on
/// every healthy workbook: a false integrity finding manufactured by the fix, and
/// Repair would "correct" a value that was already right. Four copies of one rule is
/// what made the disagreement reachable.
/// </para>
/// </remarks>
public sealed class SheetLayoutAuthorityTests
{
    /// <summary>Production source roots scanned for an out-of-authority anchor.</summary>
    private static readonly string[] ProductionRoots =
    [
        "src/GanttCreator.AddIn",
        "src/GanttCreator.Core",
        "src/GanttCreator.Office",
        "src/GanttCreator.Raster",
    ];

    /// <summary>The one file permitted to build the anchor's <c>refersTo</c>.</summary>
    private const string LayoutPath = "src/GanttCreator.Core/GanttSheetLayout.cs";

    /// <summary>
    /// No production file may build a plot-anchor <c>refersTo</c> string itself.
    /// </summary>
    /// <remarks>
    /// The scan is TEXT-based for the interpolated-string form, because the thing
    /// being forbidden is a <em>literal</em>: a string that builds
    /// <c>='Sheet'!$X$1</c> without consulting <c>GanttSheetLayout</c>. Roslyn
    /// syntax matching is not used here because the offending construct is a
    /// perfectly ordinary interpolation, and there is no syntax shape that
    /// distinguishes "an anchor built from a literal" from "an anchor built from a
    /// helper" other than the presence of the helper call.
    /// <para>
    /// The pattern requires BOTH the anchor's defining shape (<c>!$</c>) and a
    /// row-literal, and it is applied to <b>code with comments removed</b>. The first
    /// version scanned raw text and flagged <c>ExcelTypeOptionsMaterialiser</c>,
    /// which contains no anchor — only a doc comment quoting
    /// <c>='_GanttCreatorConfig'!$B$2:$B$17</c> to explain how Excel rewrites a
    /// TypeOptions reference on read-back.
    /// </para>
    /// <para>
    /// <b>Only comments are stripped — string literals are deliberately KEPT.</b>
    /// The second version also blanked string literals, which made the guard
    /// vacuous: an anchor is <em>always</em> built inside a string, so blanking
    /// strings removes exactly what the pattern searches for. That version passed
    /// while a reintroduced literal sat in the checker — only the call-site test
    /// caught it. A guard that cannot fail on the thing it names is worse than no
    /// guard, because it reports the contract as enforced.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_plot_anchor_is_not_built_outside_the_layout_authority()
    {
        string repositoryRoot = FindRepositoryRoot();
        var offenders = new List<string>();

        foreach (string root in ProductionRoots)
        {
            string directory = Path.Combine(repositoryRoot, root);

            // A missing root FAILS the test rather than being skipped. Skipping is how
            // a renamed or deleted production project would silently stop being
            // scanned, leaving this guard green while covering nothing.
            Assert.True(
                Directory.Exists(directory),
                $"Production root '{root}' does not exist under the repository root; "
                + "a missing root would be silently skipped rather than scanned.");

            foreach (string file in Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories))
            {
                string relative = Path.GetRelativePath(repositoryRoot, file).Replace('\\', '/');
                if (relative == LayoutPath)
                {
                    continue;
                }

                string code = StripComments(File.ReadAllText(file));
                bool buildsAnAnchor = code.Contains("!$", StringComparison.Ordinal);
                bool pinsARow = code.Contains("$1", StringComparison.Ordinal)
                    || code.Contains("$2", StringComparison.Ordinal);

                if (buildsAnAnchor && pinsARow)
                {
                    offenders.Add(relative);
                }
            }
        }

        Assert.True(
            offenders.Count == 0,
            "The plot anchor must be built by GanttSheetLayout alone; these files build one "
            + "from their own literal: " + string.Join(", ", offenders));
    }

    /// <summary>
    /// Blanks comment bodies, preserving line structure so a reported offender still
    /// reads as source. String literals are left intact on purpose — see the test's
    /// remarks.
    /// </summary>
    /// <param name="source">The C# source to filter.</param>
    /// <returns>The same text with comment bodies replaced by spaces.</returns>
    private static string StripComments(string source)
    {
        var result = new System.Text.StringBuilder(source.Length);
        int index = 0;

        while (index < source.Length)
        {
            char current = source[index];

            // Line comment.
            if (current == '/' && index + 1 < source.Length && source[index + 1] == '/')
            {
                while (index < source.Length && source[index] != '\n')
                {
                    _ = result.Append(' ');
                    index++;
                }

                continue;
            }

            // Block comment.
            if (current == '/' && index + 1 < source.Length && source[index + 1] == '*')
            {
                while (index < source.Length
                    && !(source[index] == '*' && index + 1 < source.Length && source[index + 1] == '/'))
                {
                    _ = result.Append(source[index] == '\n' ? '\n' : ' ');
                    index++;
                }

                for (int i = 0; i < 2 && index < source.Length; i++)
                {
                    _ = result.Append(' ');
                    index++;
                }

                continue;
            }

            _ = result.Append(current);
            index++;
        }

        return result.ToString();
    }

    /// <summary>
    /// The three former call sites must all route through the authority.
    /// </summary>
    /// <remarks>
    /// The scan above proves nobody builds an anchor <em>from a literal</em>. This
    /// proves the opposite direction — that the sites which must use the authority
    /// actually do — because "no offenders found" is equally consistent with all three
    /// call sites having been deleted.
    /// </remarks>
    [Fact]
    public void Every_anchor_call_site_routes_through_the_layout_authority()
    {
        string repositoryRoot = FindRepositoryRoot();

        string[] requiredCallSites =
        [
            "src/GanttCreator.Office/ExcelWorkbookInitialiser.cs",
            "src/GanttCreator.Office/ExcelConfigIntegrityChecker.cs",
            "src/GanttCreator.Office/ExcelPlotAnchorRepairer.cs",
        ];

        foreach (string relative in requiredCallSites)
        {
            string path = Path.Combine(repositoryRoot, relative);
            Assert.True(File.Exists(path), $"Expected call site '{relative}' does not exist.");

            string text = File.ReadAllText(path);
            Assert.True(
                text.Contains("GanttSheetLayout.BuildPlotAnchorRefersTo", StringComparison.Ordinal),
                $"'{relative}' must build its anchor through GanttSheetLayout. A local "
                + "copy of the rule here is exactly the duplication D1 exists to prevent, "
                + "and the three of them drifted once already.");
        }
    }

    /// <summary>
    /// The header row must come from the authority, not a literal.
    /// </summary>
    /// <remarks>
    /// The header row was worksheet row 1 by convention, written as
    /// <c>Cells[1, 1]</c> in the initialiser. ADR-0030's reserved row moves it to
    /// row 2, so the index has to be read from the layout authority or the table is
    /// created one row above where the anchor and the bands say it is.
    /// </remarks>
    [Fact]
    public void The_header_row_is_taken_from_the_layout_authority()
    {
        string repositoryRoot = FindRepositoryRoot();
        string path = Path.Combine(repositoryRoot, "src/GanttCreator.Office/ExcelWorkbookInitialiser.cs");

        Assert.True(File.Exists(path), "Expected the workbook initialiser to exist.");

        string text = File.ReadAllText(path);
        Assert.True(
            text.Contains("GanttSheetLayout.HeaderRowIndex", StringComparison.Ordinal),
            "The header range must be built from GanttSheetLayout.HeaderRowIndex. A literal "
            + "row index here would place the table one row away from the anchor and the bands.");
        Assert.DoesNotContain(
            "Cells[1, 1]",
            text,
            StringComparison.Ordinal);
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
