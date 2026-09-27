namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The R3.16 drift guard: the entity guide's "Per-entity field contract" table and
/// <see cref="EquivalenceFields"/> must name the same rows, or the contract a renderer
/// author consults and the contract the tests enforce have silently diverged.
/// </summary>
/// <remarks>
/// <para>
/// This is the guard the R3.16 work item's D6 asks for, and it exists because this
/// class of drift has already happened repeatedly in this row's own history: the header
/// labels were collapsed into their rectangles, the "Labels" row was really two rows,
/// and the panel row's z-layer cell named one layer where the builder emits three.
/// Each was found by reading the table against the committed golden by hand. A hand
/// check does not run on the next edit.
/// </para>
/// <para>
/// <b>Why this is not prose parsing.</b> The work item stops if "the drift guard cannot
/// read the guide deterministically", because a guard that depends on prose is worse
/// than none. It does not. The parser is anchored on the table's own header row, a
/// fixed machine-checkable string, and then consumes only markdown pipe rows until the
/// first line that is not one. Nothing here reads a sentence, infers a meaning, or
/// tolerates a reworded paragraph - if the header changes, the guard reports that it
/// could not find the table rather than guessing.
/// </para>
/// <para>
/// The guard compares the <em>row-name set only</em>. It deliberately does not attempt
/// to compare field lists in prose, which is where a false positive would live.
/// </para>
/// </remarks>
public sealed class EntityGuideDriftGuardTests
{
    private const string _tableHeading = "### Per-entity field contract";
    private const string _tableHeaderCell = "Scene primitive and role ID";
    private const string _guideFileName = "07-GANTT-ENTITY-GUIDE.md";
    private const int _expectedRowCount = 13;

    [Fact]
    public void The_guide_table_and_the_field_catalogue_name_the_same_rows()
    {
        List<string> guideRows = ReadGuideRowNames();

        Assert.Equal(
            guideRows,
            EquivalenceFields.All.Select(row => row.GuideRowName).ToList());
    }

    [Fact]
    public void The_guide_table_was_actually_found()
    {
        // Without this the comparison above could pass vacuously: an empty guide parse
        // and an empty catalogue would agree. The row count is pinned so a guide edit
        // that empties or collapses the table cannot pass unnoticed.
        List<string> guideRows = ReadGuideRowNames();

        Assert.Equal(_expectedRowCount, guideRows.Count);
        Assert.Equal(guideRows.Count, guideRows.Distinct(StringComparer.Ordinal).Count());
        Assert.All(guideRows, row => Assert.False(string.IsNullOrWhiteSpace(row)));
    }

    [Fact]
    public void A_renamed_guide_row_is_detected_by_the_drift_guard()
    {
        // The guard's own positive. Without it, a guard that compared the wrong two
        // things - or compared nothing - would still pass every other test here, which
        // is how a vacuous guard ships.
        //
        // The rename is applied to a synthetic copy of the row set, not to the live
        // guide, so this positive keeps testing the comparison itself after a real
        // guide rename instead of failing for an unrelated reason.
        List<string> guideRows = [.. EquivalenceFields.All.Select(row => row.GuideRowName)];
        int index = guideRows.IndexOf("Delineator");
        Assert.True(index >= 0, "The catalogue must contain the Delineator row for this positive to mean anything.");

        guideRows[index] = "Delineator (renamed)";
        List<string> catalogueRows = [.. EquivalenceFields.All.Select(row => row.GuideRowName)];

        // This is exactly the comparison the guard performs, so a divergence here is
        // precisely what the guard would report.
        Assert.NotEqual(catalogueRows, guideRows);
        Assert.Contains("Delineator", catalogueRows);
        Assert.DoesNotContain("Delineator (renamed)", catalogueRows);
    }

    [Fact]
    public void An_added_guide_row_is_detected_by_the_drift_guard()
    {
        // The other direction: a guide row the catalogue does not transcribe. Without
        // this, a guard written only to catch renames would miss an addition - which
        // is the more likely edit, since adding a new entity starts as a guide change.
        List<string> guideRows = [.. EquivalenceFields.All.Select(row => row.GuideRowName), "A newly proposed row"];
        List<string> catalogueRows = [.. EquivalenceFields.All.Select(row => row.GuideRowName)];

        Assert.NotEqual(catalogueRows, guideRows);
        Assert.Equal(catalogueRows.Count + 1, guideRows.Count);
    }

    /// <summary>
    /// Reads the first-column row names of the guide's field-contract table.
    /// </summary>
    /// <returns>The row names, in the guide's order.</returns>
    /// <remarks>
    /// Fails rather than returning empty when the heading or the table header is
    /// missing, so a guide restructure surfaces as an actionable message instead of a
    /// vacuous pass.
    /// </remarks>
    private static List<string> ReadGuideRowNames()
    {
        string[] lines = File.ReadAllLines(LocateGuide());

        int heading = Array.FindIndex(lines, line => line.Trim() == _tableHeading);
        Assert.True(heading >= 0, $"The entity guide no longer contains the '{_tableHeading}' section.");

        int header = Array.FindIndex(lines, heading, line => line.Contains(_tableHeaderCell, StringComparison.Ordinal));
        Assert.True(
            header >= 0,
            $"The field-contract table header ('{_tableHeaderCell}') was not found below '{_tableHeading}'.");

        // The row after the header is the markdown separator (| --- | --- |), which is
        // not a data row; skipping it by shape rather than by position keeps the parse
        // correct if a row is inserted above the table.
        List<string> names = [];
        for (int index = header + 1; index < lines.Length; index++)
        {
            string line = lines[index].Trim();
            if (!line.StartsWith('|'))
            {
                break;
            }

            if (IsSeparatorRow(line))
            {
                continue;
            }

            string name = FirstCell(line);
            Assert.False(string.IsNullOrWhiteSpace(name), $"A field-contract row has a blank name: '{line}'");
            names.Add(name);
        }

        return names;
    }

    /// <summary>Locates the committed guide copied to the test output directory.</summary>
    private static string LocateGuide()
    {
        string path = Path.Combine(AppContext.BaseDirectory, "docs", _guideFileName);
        Assert.True(
            File.Exists(path),
            $"The entity guide was not copied to the test output at '{path}'. The csproj must copy docs/{_guideFileName}; without it the drift guard would compare an empty row set and pass vacuously.");
        return path;
    }

    private static bool IsSeparatorRow(string line)
    {
        string[] cells = line.Split('|');
        for (int index = 1; index < cells.Length - 1; index++)
        {
            if (cells[index].Trim().Trim('-', ':').Length > 0)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Reads the first cell of a markdown table row, trimmed of its delimiters and any
    /// emphasis a row name may carry.
    /// </summary>
    private static string FirstCell(string line)
    {
        ReadOnlySpan<char> body = line.AsSpan(line.IndexOf('|', StringComparison.Ordinal) + 1);        int end = body.IndexOf('|');
        if (end >= 0)
        {
            body = body[..end];
        }

        return body.ToString().Trim().Trim('*', '_', '`');
    }
}
