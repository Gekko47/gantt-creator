namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Pins which COM accessor the R5.3 parent/child mutator uses to address one
/// cell of a body row, and proves the live gate exists.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a source-shape guard is the only thing that can catch this.</b> The
/// defect this pins was invisible to every behavioural test in the repository.
/// <c>ExcelParentChildMutator.GetCellAt</c> is an <c>internal virtual</c> seam,
/// so the contract harness overrides it with a delegate and the production
/// implementation — the one line that was wrong — never executed. The mock was
/// tested; the production line was not.
/// </para>
/// <para>
/// <b>The measured fact.</b> On a live Excel 16.0 <c>tblGanttData</c> body row
/// (<c>$A$5:$P$5</c>, one row x sixteen columns, obtained from
/// <c>DataBodyRange.Rows[i]</c>), the two-argument form of
/// <see cref="Excel.Range"/>'s own indexer — <c>row[1, columnIndex]</c>, i.e.
/// the <c>Item</c>/<c>_Default</c> property — is refused by the host with
/// <c>COMException 0x800A03EC</c>, and so are <c>rows[2, 10]</c> and
/// <c>row.Range[1, 10]</c>. Every equivalent form routed through <c>.Cells</c>
/// — <c>row.Cells[1, 10]</c>, <c>rows.Cells[2, 10]</c>,
/// <c>dataBody.Cells[2, 10]</c> — is accepted. The two forms are
/// indistinguishable by the shape of the range they address, so the guard pins
/// the accessor used, not the range it resolves to. That is the same discipline
/// the R5.3 OutlineLevel work recorded for <c>Worksheet.Rows[...]</c> versus
/// <c>Worksheet.Range[...]</c>.
/// </para>
/// <para>
/// <b>How it reached the user.</b> The refusal surfaced as
/// <c>ParentChildMutationRefusalReason.WriteFailed</c> on both verbs, whose
/// message reads "Excel refused a write, so the table may be out of order."
/// <c>GetCellAt</c> has exactly one caller, <c>WriteCells</c>, which serves both
/// demote's planned cells and promote's resequenced cells above the moved
/// window — which is why promoting and demoting failed identically.
/// </para>
/// </remarks>
public class ParentChildCellAccessorTests
{
    /// <summary>The adapter whose cell accessor is pinned.</summary>
    private const string MutatorPath = "src/GanttCreator.Office/ExcelParentChildMutator.cs";

    /// <summary>The adapter that already writes <c>ParentId</c> the accepted way.</summary>
    private const string RepairerPath = "src/GanttCreator.Office/ExcelGanttRowIdentityRepairer.cs";

    /// <summary>
    /// <c>GetCellAt</c> must address its cell through <c>.Cells</c>.
    /// </summary>
    /// <remarks>
    /// Only comments are stripped, never string literals, for the reason
    /// <c>SheetLayoutAuthorityTests</c> records: a guard that cannot fail on the
    /// thing it names is worse than no guard. Here the pattern is a method body,
    /// so no literal is searched for at all — but the doc comment quotes the
    /// offending <c>row[1, columnIndex]</c> verbatim to explain the refusal, so
    /// an unstripped scan would flag the very file that documents the fix.
    /// </remarks>
    [Fact]
    public void The_parent_child_cell_accessor_goes_through_cells()
    {
        string code = StripComments(File.ReadAllText(Resolve(MutatorPath)));

        Assert.DoesNotContain("row[1, columnIndex]", code, StringComparison.Ordinal);
        Assert.Contains("row.Cells[1, columnIndex]", code, StringComparison.Ordinal);
    }

    /// <summary>
    /// The guard must not be vacuous: the accepted form is the one the
    /// live-proven <c>ParentId</c> writer uses, so both adapters agree.
    /// </summary>
    /// <remarks>
    /// "No offenders found" is equally consistent with the method having been
    /// deleted or renamed, which would leave this guard green while covering
    /// nothing. This pins the guard's shape against the sibling adapter that
    /// writes the same column in production and is live-gated.
    /// </remarks>
    [Fact]
    public void The_accepted_form_is_the_one_the_live_proven_repairer_uses()
    {
        string code = StripComments(File.ReadAllText(Resolve(RepairerPath)));

        Assert.Contains("cells[row, column]", code, StringComparison.Ordinal);
    }

    /// <summary>
    /// The live gate must exist, because no mock can catch this defect.
    /// </summary>
    /// <remarks>
    /// A source-shape guard proves the accessor is the accepted one; only a live
    /// exercise proves the write lands. R5.3 shipped with its live promote/demote
    /// gate unrun — recorded as the outstanding gate in <c>docs/STATUS.md</c> —
    /// which is precisely why this defect reached a user. Deleting or renaming
    /// the live test must fail here rather than silently dropping the only real
    /// coverage.
    /// </remarks>
    [Fact]
    public void A_live_promote_and_demote_gate_exists()
    {
        string path = Resolve(
            "tests/GanttCreator.Office.IntegrationTests/ParentChildMutationIntegrationTests.cs");

        Assert.True(File.Exists(path), $"The live parent/child gate '{path}' does not exist.");
        string code = File.ReadAllText(path);
        Assert.Contains("Promote()", code, StringComparison.Ordinal);
        Assert.Contains("Demote()", code, StringComparison.Ordinal);
    }

    /// <summary>Resolves a repository-relative path from the test output directory.</summary>
    /// <param name="relative">The slash-separated repository-relative path.</param>
    /// <returns>The absolute path.</returns>
    private static string Resolve(string relative)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (Directory.Exists(Path.Combine(directory.FullName, "src"))
                && File.Exists(Path.Combine(directory.FullName, "GanttCreator.slnx")))
            {
                return Path.Combine(
                    directory.FullName,
                    relative.Replace('/', Path.DirectorySeparatorChar));
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not locate the repository root (a directory holding src/ and GanttCreator.slnx).");
    }

    /// <summary>
    /// Blanks comment bodies, preserving line structure. String literals are left
    /// intact on purpose — see <c>SheetLayoutAuthorityTests</c>.
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

            if (current == '/' && index + 1 < source.Length && source[index + 1] == '/')
            {
                while (index < source.Length && source[index] != '\n')
                {
                    _ = result.Append(' ');
                    index++;
                }

                continue;
            }

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
}

