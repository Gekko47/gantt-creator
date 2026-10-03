namespace GanttCreator.Core;

/// <summary>
/// The <b>single authority</b> for the live Gantt worksheet's vertical layout and
/// its plot anchor (R4.7I slice 1 D1, ADR-0030 D4/D5).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why one authority.</b> The header row used to be worksheet row 1 by
/// convention, and three independent files each built the plot anchor's
/// <c>refersTo</c> string with a hardcoded <c>$1</c> literal: the initialiser, the
/// integrity checker, and the plot-anchor repairer — each with its own copy of the
/// A1 column conversion. ADR-0030 adds a reserved row above the table, which moves
/// the header to row 2. Editing three literals in step is exactly the arrangement
/// that made the disagreement possible: had the initialiser written <c>$O$2</c>
/// while the checker still expected <c>$O$1</c>, the integrity check would report
/// <c>PlotAnchorDisagreement</c> on every healthy workbook — a false finding
/// manufactured by the fix. One helper makes them incapable of disagreeing.
/// </para>
/// <para>
/// <b>The reserved row is one row, not two.</b> The owner asked for two rows above
/// the table, one for the year band and one for the period band. ADR-0030 D4/D5
/// implements <em>one</em> inserted row, because entity guide §4 already requires
/// the data-panel header's height to align with the period header's bottom — the
/// header row <em>is</em> the period band's row. A second inserted row would strand
/// the header between the period band and the first lane, leaving a visible empty
/// strip one header-row tall over the plot columns. The divergence is recorded in
/// ADR-0030 rather than taken silently; changing it is a guide amendment.
/// </para>
/// <para>
/// Pure and Office-free: this is arithmetic over the schema, testable without a
/// live host.
/// </para>
/// </remarks>
public static class GanttSheetLayout
{
    /// <summary>
    /// How many rows are reserved above the table's header row.
    /// </summary>
    /// <remarks>
    /// <b>Two</b>, per ADR-0031 D1: a top padding row and the row carrying the table
    /// title and the year band. The period band still needs no row of its own
    /// (ADR-0030 D5) because it shares the header row.
    /// </remarks>
    public const int ReservedRowCount = 2;

    /// <summary>
    /// The top padding row: worksheet row 1, above everything the add-in draws. It
    /// exists so the chart's top margin is a real row the user can see and align
    /// with, rather than a sub-row sliver of chart chrome.
    /// </summary>
    public const int TopPaddingRowIndex = 1;

    /// <summary>
    /// The reserved row directly above the header: it carries the table title across
    /// the panel columns and the year band across the plot columns.
    /// </summary>
    public const int ReservedRowIndex = TopPaddingRowIndex + 1;

    /// <summary>
    /// The table's header row, and therefore the row the period band sits on.
    /// </summary>
    /// <remarks>
    /// <b>Counted from the block's start, not from its last member.</b> This read
    /// <c>ReservedRowIndex + ReservedRowCount</c>, which was correct only while
    /// <see cref="ReservedRowIndex"/> sat at the <em>first</em> row of the reserved
    /// block. Adding the top padding row inside that block made the two overlap, and
    /// the sum then skipped a row: the header landed on 4 while the reserved rows
    /// were 1 and 2, stranding an empty row between the title row and the header.
    /// The block is contiguous by definition, so the header is simply the row after
    /// the block's last row — which is <see cref="ReservedRowIndex"/> itself now
    /// that it is the last member.
    /// </remarks>
    public const int HeaderRowIndex = ReservedRowIndex + 1;

    /// <summary>
    /// The worksheet row the first body row of the table occupies.
    /// </summary>
    /// <remarks>
    /// <b>Stated as an invariant rather than a literal.</b> The plot's top edge is
    /// this row's top edge (ADR-0030 D1), so a change to
    /// <see cref="ReservedRowCount"/> must move it without anyone editing a second
    /// place.
    /// </remarks>
    public const int FirstBodyRowIndex = HeaderRowIndex + 1;

    /// <summary>
    /// Gets the one-based worksheet row of the reserved title row.
    /// </summary>
    public static int TitleRowIndex => ReservedRowIndex;

    /// <summary>
    /// Gets the one-based worksheet column carrying the table title: the column above
    /// <c>Description</c>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This was the first AUTHORING column (D), and the title was then written across
    /// a 1-by-5 <c>Resize</c> range so it would "read as one heading across the
    /// visible columns". That premise is false: assigning <c>Value2</c> to a
    /// multi-cell range writes the value into <em>every</em> cell, so the title
    /// appeared once per column instead of once. The owner asked for a single title
    /// cell directly above <c>Description</c>, which is also the column the table is
    /// read by.
    /// </para>
    /// <para>
    /// Columns A-C are <see cref="GanttColumnAccess.EngineHidden"/> bookkeeping. They
    /// are hidden but still occupy worksheet positions, so a title in column A would
    /// sit above columns the user cannot see.
    /// </para>
    /// </remarks>
    public static int TitleColumnIndex
    {
        get
        {
            IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
            for (var i = 0; i < columns.Count; i++)
            {
                if (string.Equals(columns[i].Name, "Description", StringComparison.Ordinal))
                {
                    return i + 1;
                }
            }

            throw new InvalidOperationException("The table schema has no Description column to title.");
        }
    }

    /// <summary>
    /// Gets the one-based worksheet column the table title starts in: the column above
    /// <c>Description</c>.
    /// </summary>
    /// <remarks>
    /// Retained as the start of the title's span. It now equals
    /// <see cref="TitleColumnIndex"/> because the title is ONE cell, not a row-wide
    /// heading.
    /// </remarks>
    public static int TitleStartColumnIndex => TitleColumnIndex;

    /// <summary>
    /// Gets the one-based worksheet column the table title ends in: the same single
    /// column it starts in.
    /// </summary>
    public static int TitleEndColumnIndex => TitleColumnIndex;

    /// <summary>
    /// Gets how many worksheet columns the table title spans: exactly one.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This was five. It feeds the <c>Resize</c> that built the written range, and a
    /// span greater than one is what wrote the title into every visible column.
    /// </para>
    /// <para>
    /// A title that is genuinely wider than its column should be expressed with a
    /// merge or a left-aligned overflow into empty cells, which is a presentation
    /// decision; writing the value into each cell is not that.
    /// </para>
    /// </remarks>
    public static int TitleColumnSpan => 1;

    /// <summary>
    /// Gets the one-based worksheet row of the reserved ANCHOR row: the row directly
    /// below the last body row (ADR-0038 D1).
    /// </summary>
    /// <param name="bodyRowCount">The number of body rows in the table.</param>
    /// <returns>The anchor row's one-based worksheet index.</returns>
    /// <remarks>
    /// <para>
    /// <b>This row exists to be an anchor, not to be seen.</b> Excel resizes a shape
    /// when a worksheet row is inserted only if the insertion point falls strictly
    /// BELOW the shape's <c>BottomRightCell</c> row; at or above it, the shape slides
    /// and keeps its height. Every insert the add-in performs targets one row past
    /// the body, so with the bands' bottom edge exactly on the body/anchor boundary
    /// that bottom anchor resolves INTO the body and a top insert leaves the plot
    /// unpainted (measured 2026-10-03, <c>scripts/probe-anchor-row-height.ps1</c>).
    /// Painting the bands down THROUGH this row gives their bottom a cell anchor one
    /// row below the insertion point, which is what makes them stretch.
    /// </para>
    /// <para>
    /// <b>Why the top could not be solved this way, and why that matters here.</b>
    /// ADR-0037 lifted the top edge a sub-row amount into the header row instead,
    /// because a row above the plot would paint over whatever the user put there.
    /// The two edges are therefore NOT symmetric and must not be "symmetrised": the
    /// top is an overlap into an existing row, the bottom is a reserved row of its
    /// own. A future change that mirrors one onto the other re-breaks the other.
    /// </para>
    /// <para>
    /// <b>Derived, not a constant</b>, for the same reason
    /// <see cref="FirstBodyRowIndex"/> is: it sits below a body whose length the user
    /// controls, so a published constant would be right for exactly one table length.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="bodyRowCount"/> is not positive. A table with no
    /// body has no last activity row, so there is no row for this to name.
    /// </exception>
    public static int AnchorRowIndex(int bodyRowCount)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(bodyRowCount, 1);
        return FirstBodyRowIndex + bodyRowCount;
    }

    /// <summary>
    /// Gets the one-based worksheet row of the bottom padding row: the row directly
    /// below the reserved anchor row (ADR-0038 D1).
    /// </summary>
    /// <param name="bodyRowCount">The number of body rows in the table.</param>
    /// <returns>The bottom padding row's one-based worksheet index.</returns>
    /// <remarks>
    /// <para>
    /// <b>Derived, not a constant.</b> The top padding row is a fixed index because
    /// it sits above the table, but this one sits below a body whose length the user
    /// controls. Publishing it as a constant would be a second source of truth that
    /// is wrong for every table except one particular length - the same defect
    /// <see cref="FirstBodyRowIndex"/> is derived to avoid.
    /// </para>
    /// <para>
    /// <b>It moved down one row in ADR-0038.</b> It used to be
    /// <c>FirstBodyRowIndex + bodyRowCount</c>, i.e. the row directly below the body,
    /// because the padding row was the only reserved row there. The anchor row now
    /// occupies that position and the padding row sits below it, so the index is
    /// <see cref="AnchorRowIndex"/> plus one. Every existing expectation written as
    /// "the row directly below the body" is now naming the anchor row instead, which
    /// is the change that makes the old insert path correct for the bottom.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="bodyRowCount"/> is not positive. A table with no
    /// body has no last activity row, so there is no row for this to name, and
    /// returning a plausible index would place the chart's bottom margin inside the
    /// table.
    /// </exception>
    public static int BottomPaddingRowIndex(int bodyRowCount) =>
        AnchorRowIndex(bodyRowCount) + 1;

    /// <summary>
    /// Resolves and <b>verifies</b> the reserved bottom padding row from the body's
    /// MEASURED worksheet rows (ADR-0035 D2).
    /// </summary>
    /// <param name="firstBodyRow">The body's first worksheet row, as the host reported it.</param>
    /// <param name="lastBodyRow">The body's last worksheet row, as the host reported it.</param>
    /// <param name="observedPaddingRow">
    /// The row the caller believes is the padding row, or <see langword="null"/>
    /// when no such row could be identified.
    /// </param>
    /// <param name="paddingRowIsEmpty">
    /// Whether the observed padding row was verified to carry no content, or
    /// <see langword="null"/> when emptiness was not established.
    /// </param>
    /// <returns>The resolved row, or a typed refusal.</returns>
    /// <remarks>
    /// <para>
    /// <b>Why this exists at all.</b> <see cref="BottomPaddingRowIndex"/> derives
    /// the row arithmetically from the body length, and two adapters acted on that
    /// arithmetic without ever checking it against the sheet: the row-height
    /// normaliser WROTE a height to it and the panel measurement READ it as the
    /// chart's bottom margin. Nothing established that the row existed, that it was
    /// the one below the table, or that it was empty -- so the add-in was resizing
    /// and measuring an ordinary user row, and a user who typed in it had their row
    /// silently changed.
    /// </para>
    /// <para>
    /// <b>Why the inputs are measured rather than derived.</b> Passing the derived
    /// index back in would make the check vacuous -- it would compare the layout
    /// authority against itself and always agree. The caller must pass what the
    /// host reported, which is the only way a divergence between the assumed layout
    /// and the real table can be detected at all.
    /// </para>
    /// <para>
    /// <b>Why each input can refuse.</b> A non-positive or inverted span means there
    /// is no body to sit below. A non-contiguous observed row means the sheet does
    /// not have the layout the add-in believes in. An occupied row means the user
    /// owns it. All three are reported rather than absorbed, because each is a
    /// workbook state Initialise or Repair can act on.
    /// </para>
    /// </remarks>
    public static BottomPaddingResolution ResolveBottomPaddingRow(
        int firstBodyRow,
        int lastBodyRow,
        int? observedPaddingRow,
        bool? paddingRowIsEmpty)
    {
        if (firstBodyRow < 1 || lastBodyRow < 1)
        {
            return BottomPaddingResolution.Refused(BottomPaddingRefusalReason.NonPositiveRowIndex);
        }

        if (lastBodyRow < firstBodyRow)
        {
            return BottomPaddingResolution.Refused(BottomPaddingRefusalReason.InvertedBodySpan);
        }

        // The ONLY row this may name is the one below the reserved ANCHOR row, which
        // is itself the one immediately below the body. Any other candidate means the
        // sheet disagrees with the add-in's model of it.
        //
        // This was `lastBodyRow + 1` until ADR-0038 inserted the anchor row between
        // the body and the padding. Leaving it would resolve a padding row that is
        // really the anchor row -- a row the add-in would then write the margin height
        // into, which is precisely the defect ADR-0035 D2 exists to prevent.
        var expected = lastBodyRow + 2;
        if (observedPaddingRow is not { } observed || observed != expected)
        {
            return BottomPaddingResolution.Refused(BottomPaddingRefusalReason.PaddingRowNotContiguous);
        }

        // Emptiness is checked, not assumed. `null` -- "not established" -- refuses,
        // because treating an unknown as empty is the substitution that produced the
        // original defect.
        return paddingRowIsEmpty is true
            ? BottomPaddingResolution.Ok(observed)
            : BottomPaddingResolution.Refused(BottomPaddingRefusalReason.PaddingRowOccupied);
    }

    /// <summary>
    /// Resolves and <b>verifies</b> the reserved anchor row from the body's MEASURED
    /// worksheet rows (ADR-0038 D6).
    /// </summary>
    /// <param name="firstBodyRow">The body's first worksheet row, as the host reported it.</param>
    /// <param name="lastBodyRow">The body's last worksheet row, as the host reported it.</param>
    /// <param name="observedAnchorRow">
    /// The row the caller believes is the anchor row, or <see langword="null"/> when no
    /// such row could be identified.
    /// </param>
    /// <param name="anchorRowIsEmpty">
    /// Whether the observed anchor row was verified to carry no content, or
    /// <see langword="null"/> when emptiness was not established.
    /// </param>
    /// <returns>The resolved row, or a typed refusal.</returns>
    /// <remarks>
    /// <para>
    /// <b>Verified for the same reason the padding row is.</b> The anchor row is
    /// written to by the row-height normaliser and read by the panel measurement, and
    /// both would be corrupting an ordinary user row if the add-in's model of the
    /// sheet were wrong. The two rows are resolved independently rather than as one
    /// combined result so a caller cannot verify one and silently assume the other --
    /// which is how the padding row came to be consumed in the first place.
    /// </para>
    /// <para>
    /// <b>Its own refusal reasons, not the padding row's.</b> The names differ
    /// (<c>AnchorRowNotContiguous</c>, <c>AnchorRowOccupied</c>) because the two rows
    /// fail for different workbook states and the diagnostics must say which row the
    /// user has to fix.
    /// </para>
    /// </remarks>
    public static AnchorRowResolution ResolveAnchorRow(
        int firstBodyRow,
        int lastBodyRow,
        int? observedAnchorRow,
        bool? anchorRowIsEmpty)
    {
        if (firstBodyRow < 1 || lastBodyRow < 1)
        {
            return AnchorRowResolution.Refused(AnchorRowRefusalReason.NonPositiveRowIndex);
        }

        if (lastBodyRow < firstBodyRow)
        {
            return AnchorRowResolution.Refused(AnchorRowRefusalReason.InvertedBodySpan);
        }

        // The ONLY row this may name is the one immediately below the body: every
        // insert the add-in performs targets it, so a shape whose bottom anchor does
        // not resolve into it does not stretch.
        var expected = lastBodyRow + 1;
        if (observedAnchorRow is not { } observed || observed != expected)
        {
            return AnchorRowResolution.Refused(AnchorRowRefusalReason.AnchorRowNotContiguous);
        }

        // "Not established" refuses, exactly as it does for the padding row. A row
        // whose contents the host would not report is not an empty row.
        return anchorRowIsEmpty is true
            ? AnchorRowResolution.Ok(observed)
            : AnchorRowResolution.Refused(AnchorRowRefusalReason.AnchorRowOccupied);
    }

    /// <summary>
    /// Gets the one-based worksheet column of the plot anchor: the column one to the
    /// right of the table's last column.
    /// </summary>
    public static int PlotAnchorColumnIndex => GanttTableSchema.Default.Columns.Count + 1;

    /// <summary>
    /// Builds the plot anchor's <c>refersTo</c> string: the cell one column right of
    /// the table's last column, <b>on the header row</b>.
    /// </summary>
    /// <param name="sheetName">
    /// The Gantt worksheet's label. May contain an apostrophe, which is doubled
    /// inside the quoted reference.
    /// </param>
    /// <returns>The absolute A1 reference, for example <c>='Gantt Data'!$O$2</c>.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown when <paramref name="sheetName"/> is null, empty, or whitespace. An
    /// anchor naming no sheet is not a repairable value, so it is refused rather
    /// than written as a string that could never resolve.
    /// </exception>
    public static string BuildPlotAnchorRefersTo(string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        var escaped = sheetName.Replace("'", "''", StringComparison.Ordinal);
        return string.Create(
            System.Globalization.CultureInfo.InvariantCulture,
            $"='{escaped}'!${ToA1Column(PlotAnchorColumnIndex)}${HeaderRowIndex}");
    }

    /// <summary>
    /// Builds a one-based column/row pair's A1 cell reference, for example
    /// <c>(4, 2)</c> becomes <c>D2</c>.
    /// </summary>
    /// <param name="columnIndex">The one-based column index.</param>
    /// <param name="rowIndex">The one-based row index.</param>
    /// <returns>The A1 reference.</returns>
    /// <remarks>
    /// <b>Why the title test needs this.</b> The owner asked for the title in
    /// <c>D2:H2</c>. Writing that expectation as two string literals would pass even
    /// if the layout's own column derivation were wrong, because both the code and
    /// the test would be reading the same hand-typed letters. Deriving the reference
    /// from the same indices the layout publishes makes the test check the derivation
    /// instead of restating it.
    /// </remarks>
    public static string ToA1Cell(int columnIndex, int rowIndex)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(rowIndex, 1);
        return string.Create(
            System.Globalization.CultureInfo.InvariantCulture,
            $"{ToA1Column(columnIndex)}{rowIndex}");
    }

    /// <summary>
    /// Converts a one-based column index to its A1 letters, so <c>1</c> becomes
    /// <c>A</c> and <c>27</c> becomes <c>AA</c>.
    /// </summary>
    /// <param name="columnIndex">The one-based column index.</param>
    /// <returns>The A1 column letters.</returns>
    /// <remarks>
    /// <b>Why it is not <c>((char)('A' + index - 1))</c>.</b> That single-cast form
    /// is correct only to column 26 and emits a punctuation character from column 27
    /// onward, producing a reference that names no cell. The schema has 14 columns
    /// today, so the bug would be dormant until an approved schema ADR adds enough
    /// columns to reach <c>AA</c> — and an anchor written at that point would be
    /// wrong in a way nothing would catch.
    /// </remarks>
    public static string ToA1Column(int columnIndex)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(columnIndex, 1);

        var builder = new System.Text.StringBuilder();
        var remaining = columnIndex;
        while (remaining > 0)
        {
            var digit = (remaining - 1) % 26;
            _ = builder.Insert(0, (char)('A' + digit));
            remaining = (remaining - 1) / 26;
        }

        return builder.ToString();
    }
}

/// <summary>Why a measured table body could not yield a reserved bottom padding row.</summary>
public enum BottomPaddingRefusalReason
{
    /// <summary>No reason; the row resolved.</summary>
    None = 0,

    /// <summary>The body's first or last worksheet row was not a positive index.</summary>
    NonPositiveRowIndex = 1,

    /// <summary>The body's last row precedes its first, so the span is not a body.</summary>
    InvertedBodySpan = 2,

    /// <summary>
    /// The row below the body is not the row the caller measured, so the add-in
    /// would be writing to a row it does not own.
    /// </summary>
    PaddingRowNotContiguous = 3,

    /// <summary>
    /// The reserved padding row carries content, so it is a user row rather than
    /// the chart's bottom margin.
    /// </summary>
    PaddingRowOccupied = 4,
}

/// <summary>
/// The measured outcome of resolving the reserved bottom padding row
/// (ADR-0035 D2).
/// </summary>
/// <param name="Succeeded">Whether the row was resolved and verified.</param>
/// <param name="Row">The padding row, or <see langword="null"/> on refusal.</param>
/// <param name="Refusal">The typed refusal reason on refusal.</param>
public sealed record BottomPaddingResolution(
    bool Succeeded,
    int? Row,
    BottomPaddingRefusalReason? Refusal)
{
    /// <summary>Creates a successful resolution.</summary>
    /// <param name="row">The verified padding row.</param>
    /// <returns>A successful outcome.</returns>
    public static BottomPaddingResolution Ok(int row) => new(true, row, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal.</returns>
    public static BottomPaddingResolution Refused(BottomPaddingRefusalReason refusal) => new(false, null, refusal);
}

/// <summary>Why the reserved anchor row could not be resolved (ADR-0038 D6).</summary>
/// <remarks>
/// Deliberately a separate type from <see cref="BottomPaddingRefusalReason"/> rather
/// than two more members of it: the two rows fail for different workbook states and a
/// diagnostic that cannot name which row is wrong is a diagnostic the user cannot act
/// on.
/// </remarks>
public enum AnchorRowRefusalReason
{
    /// <summary>No reason; the row resolved.</summary>
    None = 0,

    /// <summary>The body's first or last worksheet row was not a positive index.</summary>
    NonPositiveRowIndex = 1,

    /// <summary>The body's last row precedes its first, so the span is not a body.</summary>
    InvertedBodySpan = 2,

    /// <summary>
    /// The row directly below the body is not the row the caller measured, so the
    /// add-in would be writing to a row it does not own.
    /// </summary>
    AnchorRowNotContiguous = 3,

    /// <summary>
    /// The reserved anchor row carries content, so it is a user row rather than the
    /// plot's bottom anchor.
    /// </summary>
    AnchorRowOccupied = 4,
}

/// <summary>
/// The measured outcome of resolving the reserved anchor row (ADR-0038 D6).
/// </summary>
/// <param name="Succeeded">Whether the row was resolved and verified.</param>
/// <param name="Row">The anchor row, or <see langword="null"/> on refusal.</param>
/// <param name="Refusal">The typed refusal reason on refusal.</param>
public sealed record AnchorRowResolution(
    bool Succeeded,
    int? Row,
    AnchorRowRefusalReason? Refusal)
{
    /// <summary>Creates a successful resolution.</summary>
    /// <param name="row">The verified anchor row.</param>
    /// <returns>A successful outcome.</returns>
    public static AnchorRowResolution Ok(int row) => new(true, row, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal.</returns>
    public static AnchorRowResolution Refused(AnchorRowRefusalReason refusal) => new(false, null, refusal);
}
