using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelColumnPresentationRestorer"/> (R4.7C D1/D2,
/// R4.8A D2): a correct sheet writes nothing, and every refusal path fires with no
/// write.
/// </summary>
/// <remarks>
/// <para>
/// <b>Idempotence is the load-bearing claim.</b> This adapter runs on every Refresh,
/// so an implementation that rewrote a column whose flag already matched would mark
/// the workbook dirty every time — a defect a user meets as a workbook that always
/// asks to save. <c>Range.Hidden</c> is declared <see cref="object"/> in the PIA, so
/// a naive <c>!= true</c> compares boxed references and is always unequal. That is
/// exactly the bug the write-count assertion below pins.
/// </para>
/// </remarks>
public sealed class ExcelColumnPresentationRestorerTests
{
    /// <summary>
    /// Overrides the COM seams so the test can supply a table and record writes with
    /// no live Excel host.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="table">The table, or null when the table is missing.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="recorded">The flag writes the test double records.</param>
    /// <remarks>
    /// Not sealed, so <c>TestableRestorerWithDrift</c> can override
    /// <c>GetColumnAt</c> and produce a genuinely wrong flag to repair.
    /// </remarks>
    private class TestableRestorer(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject? table,
        Excel.Worksheet? worksheet,
        List<string> recorded)
        : ExcelColumnPresentationRestorer(application, guard)
    {
        /// <summary>The shared write log, so a subclass records into the test's list.</summary>
        protected List<string> Recorded { get; } = recorded;

        internal override bool TryFindGanttTable(
            Excel.Sheets sheets,
            out Excel.ListObject? resolvedTable,
            out Excel.Worksheet? resolvedWorksheet)
        {
            resolvedTable = table;
            resolvedWorksheet = table is null ? null : worksheet;
            return table is not null;
        }

        internal override Excel.ListColumns GetTableColumns(Excel.ListObject target)
        {
            Mock<Excel.ListColumns> columns = new();
            columns.Setup(c => c.Count).Returns(LiveColumnCount);
            return columns.Object;
        }

        /// <summary>How many columns the fake live table reports.</summary>
        /// <remarks>
        /// Overridable so the missing-column and extra-column cases can change the
        /// live count without duplicating the collection setup.
        /// </remarks>
        internal virtual int LiveColumnCount => GanttTableSchema.Default.Columns.Count;

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            GanttTableColumn definition = GanttTableSchema.Default.Columns[index - 1];

            // Both flags start already-correct, so only a genuinely wrong value
            // becomes a write. That makes the write-count assertion below a statement
            // about the comparison rather than about the loop.
            Mock<Excel.Range> entire = new();
            entire.SetupGet(c => c.Hidden).Returns(definition.IsHidden);
            entire.SetupSet(c => c.Hidden = It.IsAny<object>())
                .Callback((object _) => Recorded.Add("Hidden:" + index));

            Mock<Excel.Range> range = new();
            range.SetupGet(r => r.Locked).Returns(definition.IsLocked);
            range.SetupSet(r => r.Locked = It.IsAny<object>())
                .Callback((object _) => Recorded.Add("Locked:" + index));
            range.SetupGet(r => r.EntireColumn).Returns(entire.Object);

            Mock<Excel.ListColumn> column = new();
            column.SetupGet(c => c.Range).Returns(range.Object);

            // The adapter matches by name, so the double must carry the schema's
            // name or every lookup misses.
            column.SetupGet(c => c.Name).Returns(definition.Name);
            return column.Object;
        }

        internal override Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;
    }

    private static Excel.Worksheet Worksheet()
    {
        Mock<Excel.Worksheet> worksheet = new();
        return worksheet.Object;
    }

    private static Mock<Excel.Application> Application()
    {
        Mock<Excel.Application> application = new();
        Mock<Excel.Workbook> workbook = new();
        Mock<Excel.Sheets> sheets = new();
        workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }

    private static Mock<IWorksheetProtectionGuard> Guard(ProtectionGuardOutcome outcome)
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        guard.Setup(g => g.Query()).Returns(outcome);
        guard.Setup(g => g.QueryTarget(It.IsAny<Excel.Worksheet>())).Returns(outcome);
        return guard;
    }

    /// <summary>
    /// A sheet already carrying the schema's classification writes nothing, so a
    /// Refresh does not mark the workbook dirty.
    /// </summary>
    /// <remarks>
    /// This is the positive test for the comparison the adapter exists to get right.
    /// A mutation that replaced the unboxed <c>ReadHiddenFlag</c> call with a direct
    /// <c>!= wanted</c> against the boxed property would make every column appear
    /// drifted, and this count would become the schema's column count rather than 0.
    /// </remarks>
    [Fact]
    public void An_already_correct_sheet_writes_nothing()
    {
        List<string> recorded = [];
        TestableRestorer restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.ColumnsRestored);
        Assert.Empty(recorded);
    }

    /// <summary>
    /// A column whose flag has drifted is rewritten, so the restore is not vacuous.
    /// </summary>
    /// <remarks>
    /// The counterpart to the idempotence test. Without it, an adapter that wrote
    /// nothing at all would satisfy "writes nothing" and the guarantee would be
    /// untested — the same vacuity this review found twice already.
    /// </remarks>
    [Fact]
    public void A_drifted_hidden_flag_is_restored()
    {
        List<string> recorded = [];
        int firstEngineColumn = IndexOfFirstEngineColumn();

        // Every column reports Hidden = false, so each engine column is drifted.
        TestableRestorerWithDrift restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded,
            firstEngineColumn);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.True(outcome.Succeeded);

        // The drifted column is rewritten, and written to the schema's value rather
        // than to whatever it already held.
        Assert.Contains(recorded, entry => entry.StartsWith("Hidden:" + firstEngineColumn + "=", StringComparison.Ordinal));
        Assert.Contains("Hidden:" + firstEngineColumn + "=True", recorded);
    }

    /// <summary>The one-based index of the first engine-hidden column in the schema.</summary>
    private static int IndexOfFirstEngineColumn()
    {
        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        for (var index = 0; index < columns.Count; index++)
        {
            if (columns[index].IsHidden)
            {
                return index + 1;
            }
        }

        throw new InvalidOperationException("The schema declares no engine-hidden column.");
    }

    /// <summary>
    /// A restorer whose columns all report <c>Hidden = false</c>, so every
    /// engine-hidden column is genuinely drifted and must be rewritten.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="table">The table.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="recorded">The flag writes recorded.</param>
    /// <param name="driftedColumn">The one-based index reporting a wrong flag.</param>
    /// <remarks>
    /// The write log is taken from the base class's <c>Recorded</c> property rather
    /// than captured here: capturing a primary-constructor parameter that is also
    /// passed to the base constructor is the double-capture CS9107 warns about, and
    /// passing a fresh list would have recorded into a list the test never reads.
    /// </remarks>
    private sealed class TestableRestorerWithDrift(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject table,
        Excel.Worksheet worksheet,
        List<string> recorded,
        int driftedColumn)
        : TestableRestorer(application, guard, table, worksheet, recorded)
    {
        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            if (index != driftedColumn)
            {
                return base.GetColumnAt(columns, index);
            }

            Mock<Excel.Range> entire = new();
            entire.SetupGet(c => c.Hidden).Returns(false);
            entire.SetupSet(c => c.Hidden = It.IsAny<object>())
                .Callback((object value) => Recorded.Add("Hidden:" + index + "=" + value));

            Mock<Excel.Range> range = new();
            range.SetupGet(r => r.Locked)
                .Returns(GanttTableSchema.Default.Columns[index - 1].IsLocked);
            range.SetupGet(r => r.EntireColumn).Returns(entire.Object);

            Mock<Excel.ListColumn> column = new();
            column.SetupGet(c => c.Range).Returns(range.Object);
            column.SetupGet(c => c.Name)
                .Returns(GanttTableSchema.Default.Columns[index - 1].Name);
            return column.Object;
        }
    }

    /// <summary>
    /// A restorer whose live columns are in a DIFFERENT order from the schema.
    /// </summary>
    /// <remarks>
    /// The base double reports each column's flags already correct, so a
    /// misclassification cannot be detected by a write. What is detected here is
    /// WHICH column was written: under a positional walk the adapter applied
    /// <c>Id</c>'s hidden/locked state to whatever column happened to sit first.
    /// </remarks>
    private sealed class TestableRestorerReordered(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject table,
        Excel.Worksheet worksheet,
        List<string> recorded)
        : TestableRestorer(application, guard, table, worksheet, recorded)
    {
        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            // Reverse the LIVE order only. Each column still reports its own schema
            // flags, so a name-matching adapter writes nothing and a positional one
            // writes to the wrong columns.
            int schemaIndex = GanttTableSchema.Default.Columns.Count - (index - 1);
            GanttTableColumn definition = GanttTableSchema.Default.Columns[schemaIndex - 1];

            Mock<Excel.Range> entire = new();
            entire.SetupGet(c => c.Hidden).Returns(definition.IsHidden);
            entire.SetupSet(c => c.Hidden = It.IsAny<object>())
                .Callback((object value) => Recorded.Add($"Hidden:{definition.Name}={value}"));

            Mock<Excel.Range> range = new();
            range.SetupGet(r => r.Locked).Returns(definition.IsLocked);
            range.SetupSet(r => r.Locked = It.IsAny<object>())
                .Callback((object value) => Recorded.Add($"Locked:{definition.Name}={value}"));
            range.SetupGet(r => r.EntireColumn).Returns(entire.Object);

            Mock<Excel.ListColumn> column = new();
            column.SetupGet(c => c.Range).Returns(range.Object);
            column.SetupGet(c => c.Name).Returns(definition.Name);
            return column.Object;
        }

        /// <summary>The live column name at a one-based live position.</summary>
        internal static string LiveColumnNameAt(int index)
        {
            int schemaIndex = GanttTableSchema.Default.Columns.Count - (index - 1);
            return GanttTableSchema.Default.Columns[schemaIndex - 1].Name;
        }
    }

    /// <summary>
    /// A restorer whose live table omits the schema's last column.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="table">The table.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="recorded">The flag writes recorded.</param>
    private sealed class TestableRestorerMissingColumn(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject table,
        Excel.Worksheet worksheet,
        List<string> recorded)
        : TestableRestorer(application, guard, table, worksheet, recorded)
    {
        internal override int LiveColumnCount => GanttTableSchema.Default.Columns.Count - 1;
    }

    /// <summary>
    /// A restorer whose live table carries every schema column PLUS a user column
    /// the schema does not declare.
    /// </summary>
    /// <param name="application">The application object.</param>
    /// <param name="guard">The protection guard.</param>
    /// <param name="table">The table.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="recorded">The flag writes recorded.</param>
    private sealed class TestableRestorerWithExtraColumn(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject table,
        Excel.Worksheet worksheet,
        List<string> recorded)
        : TestableRestorer(application, guard, table, worksheet, recorded)
    {
        internal override int LiveColumnCount => GanttTableSchema.Default.Columns.Count + 1;

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            // The extra column is appended LAST, so every schema column keeps its
            // correct name and the only difference is the trailing surplus.
            if (index > GanttTableSchema.Default.Columns.Count)
            {
                Mock<Excel.Range> extraRange = new();
                extraRange.SetupGet(r => r.Locked).Returns(true);
                extraRange.SetupSet(r => r.Locked = It.IsAny<object>())
                    .Callback((object value) => Recorded.Add($"Locked:UserColumn={value}"));

                Mock<Excel.Range> extraEntire = new();
                extraEntire.SetupGet(c => c.Hidden).Returns(false);
                extraEntire.SetupSet(c => c.Hidden = It.IsAny<object>())
                    .Callback((object value) => Recorded.Add($"Hidden:UserColumn={value}"));

                extraRange.SetupGet(r => r.EntireColumn).Returns(extraEntire.Object);

                Mock<Excel.ListColumn> extraColumn = new();
                extraColumn.SetupGet(c => c.Range).Returns(extraRange.Object);
                extraColumn.SetupGet(c => c.Name).Returns("UserColumn");
                return extraColumn.Object;
            }

            return base.GetColumnAt(columns, index);
        }
    }

    /// <summary>
    /// Columns are matched by NAME, so a live table in a different order is
    /// restored correctly and an already-correct one writes nothing.
    /// </summary>
    /// <remarks>
    /// <b>This is the regression test for the positional walk.</b> Under a positional
    /// match the adapter read <c>schema[index - 1]</c> and applied it to whatever
    /// column sat at <c>index</c>, so a reordered table had its <c>Description</c>
    /// hidden and its <c>Id</c> left visible: no refusal, no warning, and a Refresh
    /// reporting success while making the authoring surface worse.
    /// </remarks>
    [Fact]
    public void Reordered_live_columns_are_matched_by_name_and_write_nothing()
    {
        List<string> recorded = [];
        TestableRestorerReordered restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.ColumnsRestored);

        // Non-vacuity: the live order really did differ from the schema's, so
        // "wrote nothing" cannot be explained by the orders matching.
        Assert.NotEqual(
            GanttTableSchema.Default.Columns[0].Name,
            TestableRestorerReordered.LiveColumnNameAt(1));
        Assert.Empty(recorded);
    }

    /// <summary>
    /// A schema column the live table lacks is a typed refusal, not a silent skip.
    /// </summary>
    /// <remarks>
    /// <b>The positive test for the new refusal.</b> Skipping the missing column
    /// would leave the table half-restored and still report success, telling the user
    /// nothing while a column kept whatever state it had.
    /// </remarks>
    [Fact]
    public void A_missing_schema_column_refuses_with_its_own_typed_reason()
    {
        List<string> recorded = [];
        TestableRestorerMissingColumn restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.SchemaColumnMissing, outcome.Refusal);
        Assert.Equal(0, outcome.ColumnsRestored);
        Assert.Empty(recorded);
    }

    /// <summary>
    /// A restorer whose AUTHORING column carries a mixed lock state, which Excel
    /// reports as null.
    /// </summary>
/// <param name="application">The application object.</param>
/// <param name="guard">The protection guard.</param>
/// <param name="table">The table.</param>
/// <param name="worksheet">The resolved worksheet.</param>
/// <param name="recorded">The flag writes recorded.</param>
/// <remarks>
/// The mixed-lock fixture. Excel returns null from <c>Range.Locked</c> for a range
/// whose cells do not agree, which is what a user produces by locking part of a
/// column. An AUTHORING column is the case that matters: its desired state is
/// <see langword="false"/>, so a reader that defaults an unknown value to false
/// matches, skips the write, and reports success over a range it never repaired.
/// </remarks>
private sealed class TestableRestorerMixedLockAuthoringColumn(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject table,
        Excel.Worksheet worksheet,
        List<string> recorded)
        : TestableRestorer(application, guard, table, worksheet, recorded)
    {
        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            GanttTableColumn definition = GanttTableSchema.Default.Columns[index - 1];
            if (!definition.IsLocked)
            {
                Mock<Excel.Range> entire = new();
                entire.SetupGet(c => c.Hidden).Returns(definition.IsHidden);
                entire.SetupSet(c => c.Hidden = It.IsAny<object>())
                    .Callback((object value) => Recorded.Add($"Hidden:{definition.Name}={value}"));

                Mock<Excel.Range> range = new();
                // The mixed state, exactly as Excel reports it. The PIA declares
                // `Locked` as a non-nullable `object` even though Excel genuinely
                // returns null for a range whose cells disagree, so the null has to
                // be forced here rather than passed naturally.
#pragma warning disable CS8625 // Deliberate: the host returns null for a mixed range.
                range.SetupGet(r => r.Locked).Returns((object?)null);
#pragma warning restore CS8625
                range.SetupSet(r => r.Locked = It.IsAny<object>())
                    .Callback((object value) => Recorded.Add($"Locked:{definition.Name}={value}"));
                range.SetupGet(r => r.EntireColumn).Returns(entire.Object);

                Mock<Excel.ListColumn> column = new();
                column.SetupGet(c => c.Range).Returns(range.Object);
                column.SetupGet(c => c.Name).Returns(definition.Name);
                return column.Object;
            }

            return base.GetColumnAt(columns, index);
        }
    }

    /// <summary>
    /// A mixed lock state on an AUTHORING column is repaired, not reported as correct.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The positive test for the <see langword="null"/> lock state. <c>ReadLockedFlag</c>
    /// used to default an unrecognised host value to <see langword="false"/>. For a
    /// LOCKED column that defaulted to drift and was written, which is why the defect
    /// stayed invisible: the authoring columns are the only ones whose desired state is
    /// <see langword="false"/>, and only they could report success over a mixed range.
    /// </para>
    /// <para>
    /// Every authoring column is made mixed, so the assertion is on the count as well as
    /// the content. Non-vacuity: <c>An_already_correct_sheet_writes_nothing</c> pins
    /// zero writes when the same columns report genuine booleans, so a write here can
    /// only come from the null state.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_mixed_lock_state_on_an_authoring_column_is_repaired_not_reported_correct()
    {
        List<string> recorded = [];
        TestableRestorerMixedLockAuthoringColumn restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        // Non-vacuity: the fixture really did leave authoring columns mixed.
        GanttTableColumn[] authoring =
        [
            .. GanttTableSchema.Default.Columns.Where(c => !c.IsLocked),
        ];
        Assert.NotEmpty(authoring);
        Assert.Equal(authoring.Length, outcome.ColumnsRestored);

        // Each one is written with its schema's UNLOCKED state, and the locked engine
        // columns are untouched -- an unknown state is drift, not "make everything locked".
        Assert.All(authoring, column => Assert.Contains($"Locked:{column.Name}=False", recorded));
        Assert.Equal(
            authoring.Length,
            recorded.Count(entry => entry.StartsWith("Locked:", StringComparison.Ordinal)));
    }

    /// <summary>
    /// A missing schema column found BEFORE a drifted one writes nothing (ADR-0008 D4).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The positive test for the two-pass ordering. The pre-fix adapter resolved and
    /// wrote inside one loop, so a column missing at position <em>N</em> refused only
    /// after positions 1..N-1 had been corrected -- reporting "nothing restored" over a
    /// worksheet it had already half-classified, and repeating the half-write on every
    /// Refresh. The fixture makes both drifted flags on the FIRST column differ from the
    /// schema, so a single-pass adapter records two writes before the refusal.
    /// </para>
    /// <para>
    /// Non-vacuity: the dropped-last-column fixture above takes the same path with
    /// nothing to write, and <c>A_drifted_hidden_flag_is_restored</c> shows the same
    /// drift is written when every column resolves.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_missing_column_refuses_before_writing_any_drifted_column()
    {
        List<string> recorded = [];
        TestableRestorerMissingMiddleColumn restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.SchemaColumnMissing, outcome.Refusal);
        Assert.Equal(0, outcome.ColumnsRestored);
        Assert.Empty(recorded);
    }

    /// <summary>
    /// A restorer whose live table is missing a column that is NOT the last one, and
    /// whose FIRST column has drifted.
    /// </summary>
/// <param name="application">The application object.</param>
/// <param name="guard">The protection guard.</param>
/// <param name="table">The table.</param>
/// <param name="worksheet">The resolved worksheet.</param>
/// <param name="recorded">The flag writes recorded.</param>
/// <remarks>
/// The ordering fixture. <c>TestableRestorerMissingColumn</c> drops the LAST schema
/// column, which is the one case a resolve-inside-the-write-loop handles correctly
/// by accident: nothing precedes it, so nothing had been written yet. Dropping an
/// EARLY column while a later one has drifted is the case that exposes the defect.
/// </remarks>
private sealed class TestableRestorerMissingMiddleColumn(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.ListObject table,
        Excel.Worksheet worksheet,
        List<string> recorded)
        : TestableRestorer(application, guard, table, worksheet, recorded)
    {
        /// <summary>Two live columns: the first, and one the schema does not name.</summary>
        internal override int LiveColumnCount => 2;

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
        {
            if (index != 1)
            {
                Mock<Excel.ListColumn> unrecognised = new();
                unrecognised.SetupGet(c => c.Name).Returns("NotASchemaColumn");
                return unrecognised.Object;
            }

            // The first schema column, DRIFTED on both flags, so resolving and
            // writing in a single pass would mutate it before the refusal.
            GanttTableColumn definition = GanttTableSchema.Default.Columns[0];
            Mock<Excel.Range> entire = new();
            entire.SetupGet(c => c.Hidden).Returns(false);
            entire.SetupSet(c => c.Hidden = It.IsAny<object>())
                .Callback((object value) => Recorded.Add($"Hidden:{definition.Name}={value}"));

            Mock<Excel.Range> range = new();
            range.SetupGet(r => r.Locked).Returns(!definition.IsLocked);
            range.SetupSet(r => r.Locked = It.IsAny<object>())
                .Callback((object value) => Recorded.Add($"Locked:{definition.Name}={value}"));
            range.SetupGet(r => r.EntireColumn).Returns(entire.Object);

            Mock<Excel.ListColumn> column = new();
            column.SetupGet(c => c.Range).Returns(range.Object);
            column.SetupGet(c => c.Name).Returns(definition.Name);
            return column.Object;
        }
    }

    /// <summary>
/// A restorer whose live table carries every schema column PLUS a user column
/// the schema does not declare.
    /// </summary>
    /// <remarks>
    /// The counterweight to the refusal above, so the two cannot be confused. The
    /// adapter restores the add-in's own classification and does not own a user's
    /// extra column, so that column is simply never visited.
    /// </remarks>
    [Fact]
    public void A_live_column_absent_from_the_schema_is_left_unchanged()
    {
        List<string> recorded = [];
        TestableRestorerWithExtraColumn restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.ColumnsRestored);
        Assert.DoesNotContain(recorded, entry => entry.Contains("UserColumn", StringComparison.Ordinal));
    }

    /// <summary>A protected target refuses and writes nothing (ADR-0008 D4).</summary>
    [Fact]
    public void A_protected_target_refuses_without_writing()
    {
        List<string> recorded = [];
        TestableRestorer restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.SheetProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.TargetProtected, outcome.Refusal);
        Assert.Empty(recorded);
    }

    /// <summary>A missing table refuses with its own typed reason.</summary>
    [Fact]
    public void A_missing_table_refuses_without_writing()
    {
        List<string> recorded = [];
        TestableRestorer restorer = new(
            Application().Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            table: null,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.TableMissing, outcome.Refusal);
        Assert.Empty(recorded);
    }

    /// <summary>A missing workbook refuses with the no-workbook reason.</summary>
    [Fact]
    public void A_missing_workbook_refuses_without_writing()
    {
        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);
        List<string> recorded = [];
        TestableRestorer restorer = new(
            application.Object,
            Guard(ProtectionGuardOutcome.NotProtected).Object,
            new Mock<Excel.ListObject>().Object,
            Worksheet(),
            recorded);

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.NoActiveWorkbook, outcome.Refusal);
        Assert.Empty(recorded);
    }

    /// <summary>A foreign application object degrades rather than throwing.</summary>
    [Fact]
    public void A_foreign_application_object_degrades_to_no_workbook()
    {
        ExcelColumnPresentationRestorer restorer = new(new object());

        ColumnPresentationOutcome outcome = restorer.EnsureClassification();

        Assert.False(outcome.Succeeded);
        Assert.Equal(ColumnPresentationRefusalReason.NoActiveWorkbook, outcome.Refusal);
    }
}