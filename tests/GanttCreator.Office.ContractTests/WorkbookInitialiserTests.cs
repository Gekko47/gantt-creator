using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelWorkbookInitialiser"/>. They need no live
/// Office: the application object is absent, foreign, or a Moq proxy of the
/// Excel PIA interfaces. Every refusal validator has a positive test that
/// constructs the bad input and asserts the typed refusal fires with no
/// mutation (work item R2.2; AGENTS.md validator rule).
/// </summary>
/// <remarks>
/// Excel COM parameterised properties (indexers) cannot appear in Moq
/// expression trees (CS0855), so the three <c>internal virtual</c> seams on
/// <see cref="ExcelWorkbookInitialiser"/> are substituted by
/// <see cref="TestableInitialiser"/>; the real indexer behaviour is exercised
/// by the tagged live-Office integration test.
/// </remarks>
public class WorkbookInitialiserTests
{
    /// <summary>
    /// The expected plot-anchor <c>refersTo</c>, derived from the schema rather
    /// than pinned to a letter. The anchor sits one column right of the table's
    /// last column, so R4.7A's <c>SiblingOrder</c> correctly moved it from
    /// <c>$O$1</c> to <c>$P$1</c>; a pinned letter would have failed without
    /// saying why, and would have failed again on the next column.
    /// </summary>
    private static string ExpectedPlotAnchorRefersTo()
    {
        int anchorColumnIndex = GanttTableSchema.Default.Columns.Count + 1;
        return $"='{GanttWorkbookContract.GanttSheetLabel}'!${ColumnLetter(anchorColumnIndex)}$1";
    }

    /// <summary>Converts a one-based Excel column index to its letters.</summary>
    private static string ColumnLetter(int oneBasedIndex)
    {
        var letters = string.Empty;
        var remaining = oneBasedIndex;
        while (remaining > 0)
        {
            int index = (remaining - 1) % 26;
            letters = (char)('A' + index) + letters;
            remaining = (remaining - 1) / 26;
        }

        return letters;
    }

    /// <summary>
    /// One mocked worksheet: its name (getter backed by a field the setter
    /// updates, modelling Excel's read-back of a renamed sheet), its header
    /// range (whose <c>Value2</c> writes are captured), its list objects and
    /// the <c>tblGanttData</c> list object, and its sheet-scoped names.
    /// </summary>
    private sealed class WorksheetGraph
    {
        public Mock<Excel.Worksheet> Worksheet { get; } = new();
        public Mock<Excel.Range> HeaderRange { get; } = new();
        public Mock<Excel.Range> UsedRange { get; } = new();
        public Mock<Excel.ListObjects> ListObjects { get; } = new();
        public Mock<Excel.ListObject> Table { get; } = new();
        public Mock<Excel.ListColumns> ListColumns { get; } = new();

        /// <summary>
        /// One mocked list column per schema column, each with its own range so
        /// the hidden and locked writes are attributable to a specific column
        /// rather than collapsing into one shared proxy.
        /// </summary>
        public List<Mock<Excel.ListColumn>> Columns { get; } = [];

        /// <summary>The <c>Locked</c> value written to each column, in column order.</summary>
        public List<bool> LockedWrites { get; } = [];

        /// <summary>The <c>Hidden</c> value written to each column, in column order.</summary>
        public List<bool> HiddenWrites { get; } = [];

        /// <summary>
        /// Pre-sets each column's current <c>Hidden</c> state, so a test can start from
        /// a sheet that already matches (or deliberately contradicts) the schema.
        /// </summary>
        /// <param name="states">The schema index and the state that column already carries.</param>
        public void SetExistingHiddenState(IEnumerable<(int Index, bool Hidden)> states)
        {
            foreach ((int index, bool hidden) in states)
            {
                _existingHidden[index] = hidden;
            }

            // The per-column mocks captured their starting value when the graph was
            // constructed, so a later call has to push the new value into each one.
            for (var index = 0; index < _presettableHidden.Count; index++)
            {
                _presettableHidden[index][0] = _existingHidden.TryGetValue(index, out bool set) && set;
            }
        }

        /// <summary>The one-element boxed state each column's <c>Hidden</c> mock reads.</summary>
        private readonly List<object[]> _presettableHidden = [];

        /// <summary>
        /// Unboxes an <c>object</c>-typed PIA property write. <c>IRange.Locked</c>
        /// and <c>IRange.Hidden</c> are declared <see cref="object"/> in this
        /// interop assembly (probed, not assumed), so the value arrives boxed and
        /// the callback must accept <see cref="object"/>. A non-boolean would mean
        /// the adapter wrote something other than a flag, which is worth failing on
        /// rather than coercing.
        /// </summary>
        /// <param name="raw">The boxed host value.</param>
        /// <returns>The boolean it carries.</returns>
        private static bool RequireBool(object raw) => raw is bool value
            ? value
            : throw new Xunit.Sdk.XunitException(
                $"Expected a boolean cell-format write but received '{raw?.GetType().Name ?? "null"}'.");
        public Mock<Excel.Names> Names { get; } = new();
        public Mock<Excel.Shapes> Shapes { get; } = new();
        public Mock<Excel.Comments> Comments { get; } = new();
        public Mock<Excel.CommentsThreaded> ThreadedComments { get; } = new();
        public Mock<Excel.PivotTables> PivotTables { get; } = new();
        public Mock<Excel.QueryTables> QueryTables { get; } = new();
        public Mock<Excel.Hyperlinks> Hyperlinks { get; } = new();
        public List<object> WrittenValues { get; } = new();
        public List<string> AssignedTableNames { get; } = new();
        public List<Excel.XlSheetVisibility> VisibleValues { get; } = new();
        public List<bool> AutoFilterSettings { get; } = new();
        public List<bool> RowStripeSettings { get; } = new();
        public List<bool> ColumnStripeSettings { get; } = new();

        private string _name;
        private string? _existingTableName;
        private bool _tableWasCreated;

        /// <summary>
        /// The <c>Hidden</c> state each column starts in, keyed by schema index. A
        /// column absent from this map reads as visible, which is what an untouched
        /// worksheet column reports.
        /// </summary>
        private readonly Dictionary<int, bool> _existingHidden = [];

        /// <summary>
        /// The final <c>Hidden</c> state of each column, in column order, read back
        /// from what the adapter wrote. This is the observable end state; a per-column
        /// write list cannot answer "is this column hidden?" once writes are skipped
        /// for columns that already match.
        /// </summary>
        public bool[] HiddenStates => [.. _finalHidden];

        private readonly List<bool> _finalHidden = [];

        public WorksheetGraph(string name)
        {
            _name = name;
            _ = Worksheet.SetupGet(w => w.Name).Returns(() => _name);
            _ = Worksheet.SetupSet(w => w.Name = It.IsAny<string>())
                .Callback<string>(value => _name = value);
            _ = Worksheet.SetupGet(w => w.UsedRange).Returns(UsedRange.Object);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = Worksheet.SetupGet(w => w.Names).Returns(Names.Object);
            _ = Worksheet.SetupGet(w => w.Shapes).Returns(Shapes.Object);
            _ = Worksheet.SetupGet(w => w.Comments).Returns(Comments.Object);
            _ = Worksheet.SetupGet(w => w.CommentsThreaded).Returns(ThreadedComments.Object);
            _ = Worksheet.SetupGet(w => w.QueryTables).Returns(QueryTables.Object);
            _ = Worksheet.SetupGet(w => w.Hyperlinks).Returns(Hyperlinks.Object);
            _ = Shapes.SetupGet(s => s.Count).Returns(() => ShapeCount);
            _ = Comments.SetupGet(c => c.Count).Returns(() => CommentCount);
            _ = ThreadedComments.SetupGet(c => c.Count).Returns(() => ThreadedCommentCount);
            _ = Names.SetupGet(n => n.Count).Returns(() => SheetNameCount);
            _ = ListObjects.SetupGet(l => l.Count).Returns(() => ListObjectCount);
            _ = PivotTables.SetupGet(p => p.Count).Returns(() => PivotTableCount);
            _ = QueryTables.SetupGet(q => q.Count).Returns(() => QueryTableCount);
            _ = Hyperlinks.SetupGet(h => h.Count).Returns(() => HyperlinkCount);

            _ = HeaderRange.SetupSet(r => r.Value2 = It.IsAny<object>())
                .Callback<object>(value => WrittenValues.Add(value));

            _ = Worksheet.SetupSet(w => w.Visible = It.IsAny<Excel.XlSheetVisibility>())
                .Callback<Excel.XlSheetVisibility>(value => VisibleValues.Add(value));

            _ = Table.SetupSet(t => t.Name = It.IsAny<string>())
                .Callback<string>(value =>
                {
                    AssignedTableNames.Add(value);
                    MarkTableCreated();
                });

            // R2.7 appearance settings (ADR-0007 D8): plain bool properties,
            // captured so the mutation set can assert them.
            _ = Table.SetupSet(t => t.ShowAutoFilter = It.IsAny<bool>())
                .Callback<bool>(value => AutoFilterSettings.Add(value));
            _ = Table.SetupSet(t => t.ShowTableStyleRowStripes = It.IsAny<bool>())
                .Callback<bool>(value => RowStripeSettings.Add(value));
            _ = Table.SetupSet(t => t.ShowTableStyleColumnStripes = It.IsAny<bool>())
                .Callback<bool>(value => ColumnStripeSettings.Add(value));
            _ = ListObjects.Setup(l => l.Add(
                    It.IsAny<Excel.XlListObjectSourceType>(),
                    It.IsAny<object>(),
                    It.IsAny<object>(),
                    It.IsAny<Excel.XlYesNoGuess>(),
                    It.IsAny<object>()))
                .Returns(Table.Object);

            // R4.7C column presentation: one range per schema column, so a
            // hidden/locked write can be attributed to a specific column. An
            // unconfigured interface property returns null under Moq, which is
            // why these are wired explicitly — otherwise the presentation step
            // would silently no-op and its tests would pass vacuously.
            _ = Table.SetupGet(t => t.ListColumns).Returns(ListColumns.Object);

            // `ListColumns.Count` is read by every column walk, including the
            // rollback's visibility restore. Unconfigured, Moq returns 0 and the
            // walk silently does nothing -- which would make the rollback test pass
            // for the wrong reason.
            _ = ListColumns.SetupGet(l => l.Count)
                .Returns(() => _tableWasCreated || _existingTableName is not null
                    ? Columns.Count
                    : 0);
            for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
            {
                // A `for` loop has one `index` variable, so the closures below would
                // all capture the same final value. Each column needs its own copy,
                // or every recorded write lands on the last column.
                int columnIndex = index;
                var columnRange = new Mock<Excel.Range>();
                var entireColumn = new Mock<Excel.Range>();
                var column = new Mock<Excel.ListColumn>();
                _ = column.SetupGet(c => c.Range).Returns(columnRange.Object);
                _ = columnRange.SetupGet(r => r.EntireColumn).Returns(entireColumn.Object);
                _ = columnRange.SetupSet(r => r.Locked = It.IsAny<object>())
                    .Callback<object>(value => LockedWrites.Add(RequireBool(value)));

                // The `Hidden` getter is wired to the recorded state, not left
                // unconfigured. `Range.Hidden` is `object`-typed, so an unconfigured
                // getter returns null, and the adapter cannot then tell "visible" from
                // "the host did not say" -- a test that wants to prove the adapter
                // SKIPS a write it does not need has to hand it a real state. A column
                // with no preset state starts as visible, the ordinary default for an
                // untouched worksheet column. The value is boxed because the PIA
                // delivers an `object`; the adapter under test does the unboxing.
                var hidden = new object[1];
                hidden[0] = _existingHidden.TryGetValue(columnIndex, out bool startHidden) && startHidden;
                _presettableHidden.Add(hidden);
                _finalHidden.Add(false);
                _ = entireColumn.SetupGet(r => r.Hidden).Returns(() => hidden[0]);
                _ = entireColumn.SetupSet(r => r.Hidden = It.IsAny<object>())
                    .Callback<object>(value =>
                    {
                        hidden[0] = value;
                        _finalHidden[columnIndex] = RequireBool(value);
                        HiddenWrites.Add(RequireBool(value));
                    });
                Columns.Add(column);
            }
        }

        /// <summary>The sheet name, as Excel would read it back (post-rename).</summary>
        public string Name => _name;

        /// <summary>Gets or sets the mocked <c>WorksheetFunction.CountA</c> value for this sheet.</summary>
        public double NonEmptyCellCount { get; set; }
        public int ShapeCount { get; set; }
        public int CommentCount { get; set; }
        public int ThreadedCommentCount { get; set; }
        public int SheetNameCount { get; set; }
        public int PivotTableCount { get; set; }
        public int QueryTableCount { get; set; }
        public int HyperlinkCount { get; set; }
        public string UsedRangeAddress { get; set; } = "$A$1";
        public int ListObjectCount { get; set; }

        /// <summary>Wires the list-object collection to contain exactly one named table, or nothing.</summary>
        /// <param name="tableName">The existing table name, or <see langword="null"/> for none.</param>
        public void WithExistingTable(string? tableName)
        {
            _existingTableName = tableName;
            ListObjectCount = tableName is null ? 0 : 1;
            if (tableName is not null)
            {
                _ = Table.SetupGet(t => t.Name).Returns(tableName);
            }
        }

        internal Excel.ListObject TableAt(int index)
        {
            Assert.Equal(1, index);
            return _existingTableName is null && !_tableWasCreated
                ? throw new InvalidOperationException("No table was configured for this sheet.")
                : Table.Object;
        }

        /// <summary>
        /// Records that the adapter created the add-in's own table on this sheet, so
        /// the rollback paths that look the table up again can find it.
        /// </summary>
        internal void MarkTableCreated()
        {
            _tableWasCreated = true;
            ListObjectCount = 1;
            _ = Table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
        }

        /// <summary>
        /// Verifies the <c>tblGanttData</c> list object was created the given
        /// number of times on this sheet. Asserted through <see cref="Mock{T}.Verify"/>
        /// rather than a callback: the PIA's optional parameters carry no
        /// declared defaults, so Moq cannot re-invoke argument callbacks with
        /// the <see cref="Type.Missing"/> values the COM binder inserts.
        /// </summary>
        /// <param name="times">The expected call count.</param>
        public void VerifyTableCreated(Times times)
            => ListObjects.Verify(
                l => l.Add(
                    It.IsAny<Excel.XlListObjectSourceType>(),
                    It.IsAny<object>(),
                    It.IsAny<object>(),
                    It.IsAny<Excel.XlYesNoGuess>(),
                    It.IsAny<object>()),
                times);

        /// <summary>
        /// Verifies the sheet-scoped plot-anchor defined-name write: when
        /// <paramref name="expectedRefersTo"/> is non-null it must have been
        /// written with exactly that <c>refersTo</c>, the given number of
        /// times; when null it must never have been written.
        /// </summary>
        /// <param name="expectedRefersTo">The expected <c>refersTo</c>, or <see langword="null"/> for never.</param>
        /// <param name="times">The expected call count.</param>
        public void VerifyAnchorName(string? expectedRefersTo, Times times)
            => Names.Verify(
                n => n.Add(
                    GanttWorkbookContract.PlotAnchorDefinedName,
                    It.Is<object>(v => expectedRefersTo == null || (string)v == expectedRefersTo)),
                times);
    }

    /// <summary>
    /// The testable initialiser: substitutes the three internal virtual
    /// indexer seams with delegates backed by the graphs, and exercises every
    /// other member through the Moq proxies.
    /// </summary>
    private sealed class StubTypeOptionsMaterialiser : ITypeOptionsMaterialiser
    {
        public TypeOptionsMaterialiseOutcome Materialise() => TypeOptionsMaterialiseOutcome.Ok();

        public TypeOptionsMaterialiseOutcome EnsureCurrent() => TypeOptionsMaterialiseOutcome.Ok();
    }

    private sealed class TestableInitialiser : ExcelWorkbookInitialiser
    {
        public TestableInitialiser(
            object? application,
            Func<Excel.Sheets, int, Excel.Worksheet> sheetAt,
            Func<Excel.ListObjects, int, Excel.ListObject> tableAt,
            Func<Excel.Worksheet, Excel.Range> headerRangeAt,
            Func<Excel.Range, string> usedRangeAddressAt,
            Func<Excel.Worksheet, int> pivotTableCountAt,
            IConfigCatalogueWriter catalogueWriter,
            IWorksheetProtectionGuard? protectionGuard = null)
            : base(application, catalogueWriter, protectionGuard, new StubTypeOptionsMaterialiser())
        {
            SheetAt = sheetAt;
            TableAt = tableAt;
            HeaderRangeAt = headerRangeAt;
            UsedRangeAddressAt = usedRangeAddressAt;
            PivotTableCountAt = pivotTableCountAt;
        }

        /// <summary>
        /// Wires the R4.7C column-presentation seams. The graph resolves the
        /// table it created back to the list-column collection and columns of the
        /// sheet that owns it.
        /// </summary>
        /// <param name="listColumnsAt">Resolves a table to its list-column collection.</param>
        /// <param name="columnAt">Resolves a list column by its one-based index.</param>
        public void WireColumnPresentation(
            Func<Excel.ListObject, Excel.ListColumns> listColumnsAt,
            Func<Excel.ListColumns, int, Excel.ListColumn> columnAt)
        {
            ListColumnsAt = listColumnsAt;
            ColumnAt = columnAt;
        }

        private Func<Excel.Worksheet, Excel.Range> HeaderRangeAt { get; }

        private Func<Excel.Range, string> UsedRangeAddressAt { get; }

        private Func<Excel.Worksheet, int> PivotTableCountAt { get; }

        private Func<Excel.Sheets, int, Excel.Worksheet> SheetAt { get; }

        private Func<Excel.ListObjects, int, Excel.ListObject> TableAt { get; }

        private Func<Excel.ListObject, Excel.ListColumns> ListColumnsAt { get; set; } =
            _ => throw new InvalidOperationException("ListColumns seam was not wired.");

        private Func<Excel.ListColumns, int, Excel.ListColumn> ColumnAt { get; set; } =
            (_, _) => throw new InvalidOperationException("GetColumnAt seam was not wired.");

        internal override object GetSheetAt(Excel.Sheets sheets, int index)
            => SheetAt(sheets, index);

        internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index)
            => TableAt(listObjects, index);

        internal override Excel.Range GetHeaderRange(Excel.Worksheet target, int columnCount)
            => HeaderRangeAt(target);

        internal override string GetUsedRangeAddress(Excel.Range usedRange)
            => UsedRangeAddressAt(usedRange);

        internal override int GetPivotTableCount(Excel.Worksheet worksheet)
            => PivotTableCountAt(worksheet);

        internal override Excel.ListColumns GetTableColumns(Excel.ListObject table)
            => ListColumnsAt(table);

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
            => ColumnAt(columns, index);

        internal override Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;
    }

    /// <summary>The workbook-level graph: application, workbook, sheets, worksheet function.</summary>
    private sealed class WorkbookGraph
    {
        public Mock<Excel.Application> Application { get; } = new();
        public Mock<Excel.Workbook> Workbook { get; } = new();
        public Mock<Excel.Sheets> Sheets { get; } = new();
        public Mock<Excel.WorksheetFunction> Functions { get; } = new();
        public Mock<Excel.Names> WorkbookNames { get; } = new();
        public int WorkbookNameCount { get; set; }
        public List<WorksheetGraph> Graphs { get; } = new();
        public Dictionary<int, Excel.Worksheet> SheetsByIndex { get; } = new();
        public Queue<Excel.Worksheet> SheetsAdded { get; } = new();

        public WorkbookGraph(params WorksheetGraph[] sheets)
        {
            for (var index = 0; index < sheets.Length; index++)
            {
                Graphs.Add(sheets[index]);
                SheetsByIndex[index + 1] = sheets[index].Worksheet.Object;
            }

            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            _ = Application.SetupGet(a => a.WorksheetFunction).Returns(Functions.Object);
            _ = Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            _ = Workbook.SetupGet(w => w.Names).Returns(WorkbookNames.Object);
            _ = Workbook.SetupGet(w => w.ActiveSheet).Returns(Graphs[0].Worksheet.Object);
            _ = WorkbookNames.SetupGet(n => n.Count).Returns(() => WorkbookNameCount);
            _ = Sheets.SetupGet(s => s.Count).Returns(() => SheetsByIndex.Count);
            _ = Sheets.Setup(s => s.Add(
                    It.IsAny<object>(), It.IsAny<object>(),
                    It.IsAny<object>(), It.IsAny<object>()))
                .Returns(() =>
                {
                    // Dequeuing and registering must be one atomic step:
                    // Moq executes a Returns factory before any Callback
                    // behaviour, so a separate Callback cannot Peek the
                    // sheet this call is about to consume.
                    var sheet = SheetsAdded.Dequeue();
                    SheetsByIndex[SheetsByIndex.Count + 1] = sheet;
                    return sheet;
                });
        }

        /// <summary>
        /// Verifies the number of sheet additions performed through the
        /// workbook. Asserted through <see cref="Mock{T}.Verify"/> rather than
        /// a callback (see <see cref="WorksheetGraph.VerifyTableCreated"/>).
        /// </summary>
        /// <param name="times">The expected call count.</param>
        public void VerifySheetsAdded(Times times)
            => Sheets.Verify(
                s => s.Add(
                    It.IsAny<object>(), It.IsAny<object>(),
                    It.IsAny<object>(), It.IsAny<object>()),
                times);

        /// <summary>
        /// Registers a sheet created during the test (the created Gantt sheet
        /// or the configuration sheet) so the seam delegates resolve its
        /// members, and queues it as the next <c>Sheets.Add</c> result. The
        /// sheet is also added to <c>SheetsByIndex</c> so iteration-based
        /// operations (e.g. <see cref="RollBackConfigurationSheet"/>) can
        /// find it by index.
        /// </summary>
        /// <param name="sheet">The graph of the sheet Excel will return.</param>
        public void EnqueueCreated(WorksheetGraph sheet)
        {
            Graphs.Add(sheet);
            SheetsAdded.Enqueue(sheet.Worksheet.Object);
            // The sheet is added to SheetsByIndex via the sheets.Add mock callback
            // when the sheet is actually created (sheets.Add is called by
            // Initialise). This keeps the timing accurate: the sheet is not visible
            // to iteration-based operations (like NameTakenByOtherSheet or
            // RollBackConfigurationSheet) until after it has been created.
        }

        /// <summary>
        /// Completes the graph: active sheet, blankness wiring, and the seam
        /// delegates. Returns the initialiser under test.
        /// </summary>
        /// <param name="activeSheet">The active worksheet graph, or <see langword="null"/> for a non-worksheet active object.</param>
        /// <returns>The initialiser over the mocked application.</returns>
        public TestableInitialiser Build(WorksheetGraph? activeSheet)
        {
            _ = Workbook.SetupGet(w => w.ActiveSheet)
                .Returns(activeSheet is null ? new object() : activeSheet.Worksheet.Object);
            if (activeSheet is not null)
            {
                var nonEmpty = activeSheet.NonEmptyCellCount;
                _ = Functions.Setup(f => f.CountA(It.IsAny<object>())).Returns(nonEmpty);
            }

            var graphs = Graphs;
            var byIndex = SheetsByIndex;
            var catalogueWriter = new Mock<IConfigCatalogueWriter>();
            _ = catalogueWriter.Setup(w => w.Write()).Returns(ConfigWriteOutcome.Ok());
            TestableInitialiser initialiser = new(
                Application.Object,
                sheetAt: (_, index) => byIndex[index],
                tableAt: (listObjects, index) =>
                    graphs.Single(g => ReferenceEquals(g.ListObjects.Object, listObjects))
                        .TableAt(index),
                headerRangeAt: target =>
                    graphs.Single(g => ReferenceEquals(g.Worksheet.Object, target))
                        .HeaderRange.Object,
                usedRangeAddressAt: usedRange =>
                    graphs.Single(g => ReferenceEquals(g.UsedRange.Object, usedRange))
                        .UsedRangeAddress,
                pivotTableCountAt: target =>
                    graphs.Single(g => ReferenceEquals(g.Worksheet.Object, target))
                        .PivotTableCount,
                catalogueWriter: catalogueWriter.Object);
            WireColumnPresentationFor(graphs, initialiser);
            return initialiser;
        }

        /// <summary>
        /// Wires the column-presentation seams of <paramref name="initialiser"/>
        /// to the sheet graphs, so the R4.7C hidden/locked writes are observable.
        /// </summary>
        /// <param name="graphs">The sheet graphs, which own the list columns.</param>
        /// <param name="initialiser">The initialiser to wire.</param>
        private static void WireColumnPresentationFor(
            IReadOnlyCollection<WorksheetGraph> graphs,
            TestableInitialiser initialiser) =>
            initialiser.WireColumnPresentation(
                table => graphs.Single(g => ReferenceEquals(g.Table.Object, table)).ListColumns.Object,
                (columns, index) => graphs
                    .Single(g => ReferenceEquals(g.ListColumns.Object, columns))
                    .Columns[index - 1].Object);
        /// <summary>
        /// Completes the graph with an explicit catalogue writer instead of
        /// the default no-op stub: the given writer's outcome drives the
        /// catalogue step of the mutation order.
        /// </summary>
        /// <param name="activeSheet">The active worksheet graph, or <see langword="null"/> for a non-worksheet active object.</param>
        /// <param name="catalogueWriter">The writer the initialiser under test uses.</param>
        /// <returns>The initialiser over the mocked application.</returns>
        public TestableInitialiser BuildWithWriter(
            WorksheetGraph? activeSheet,
            IConfigCatalogueWriter catalogueWriter,
            IWorksheetProtectionGuard? protectionGuard = null)
        {
            _ = Workbook.SetupGet(w => w.ActiveSheet)
                .Returns(activeSheet is null ? new object() : activeSheet.Worksheet.Object);
            if (activeSheet is not null)
            {
                var nonEmpty = activeSheet.NonEmptyCellCount;
                _ = Functions.Setup(f => f.CountA(It.IsAny<object>())).Returns(nonEmpty);
            }

            var graphs = Graphs;
            var byIndex = SheetsByIndex;
            TestableInitialiser initialiser = new(
                Application.Object,
                sheetAt: (_, index) => byIndex[index],
                tableAt: (listObjects, index) =>
                    graphs.Single(g => ReferenceEquals(g.ListObjects.Object, listObjects))
                        .TableAt(index),
                headerRangeAt: target =>
                    graphs.Single(g => ReferenceEquals(g.Worksheet.Object, target))
                        .HeaderRange.Object,
                usedRangeAddressAt: usedRange =>
                    graphs.Single(g => ReferenceEquals(g.UsedRange.Object, usedRange))
                        .UsedRangeAddress,
                pivotTableCountAt: target =>
                    graphs.Single(g => ReferenceEquals(g.Worksheet.Object, target))
                        .PivotTableCount,
                catalogueWriter: catalogueWriter,
                protectionGuard: protectionGuard);
            WireColumnPresentationFor(graphs, initialiser);
            return initialiser;
        }
    }

    // ---------------------------------------------------------------------
    // Adopt path (blank active worksheet)
    // ---------------------------------------------------------------------

    [Fact]
    public void Initialise_adopts_a_blank_active_worksheet_and_renames_it()
    {
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);

        var outcome = graph.Build(active).Initialise();

        Assert.True(outcome.Succeeded);
        Assert.Equal(WorkbookInitialisePath.Adopted, outcome.Path);
        Assert.Equal(GanttWorkbookContract.GanttSheetLabel, outcome.SheetName);
        Assert.Equal(GanttWorkbookContract.GanttSheetLabel, active.Name);
        active.VerifyTableCreated(Times.Once());
        Assert.Equal(new[] { GanttTableSchema.TableName }, active.AssignedTableNames);
    }

    [Fact]
    public void Initialise_sets_the_neutral_table_appearance_on_the_visible_table()
    {
        // ADR-0007 D8: no autofilter dropdowns and no banded rows; the three
        // flags are the complete visible-table appearance contract.
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);

        _ = graph.Build(active).Initialise();

        Assert.Equal(FalseSettings, active.AutoFilterSettings);
        Assert.Equal(FalseSettings, active.RowStripeSettings);
        Assert.Equal(FalseSettings, active.ColumnStripeSettings);
    }

    [Fact]
    public void Initialise_applies_the_column_presentation_contract()
    {
        // R4.7C D1/D2, ADR-0029 D7/D8. Every schema column is locked, and what it
        // carries is derived from the schema's classification rather than from a list
        // maintained beside it.
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);

        _ = graph.Build(active).Initialise();

        IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
        Assert.Equal(columns.Count, active.LockedWrites.Count);

        for (var index = 0; index < columns.Count; index++)
        {
            Assert.Equal(columns[index].IsLocked, active.LockedWrites[index]);
        }

        // Visibility is a write-if-different, so only the columns whose state
        // actually has to change appear. On a fresh worksheet every authoring column
        // already reads visible and no engine column does, so every recorded write is
        // `true` and there is exactly one per hidden column -- not one per column,
        // which is what the always-true guard used to produce.
        Assert.Equal(
            columns.Count(column => column.IsHidden),
            active.HiddenWrites.Count);
        Assert.All(active.HiddenWrites, written => Assert.True(written));
    }

    [Fact]
    public void Initialise_hides_every_engine_column_and_leaves_the_authoring_ones_visible()
    {
        // Stated column-by-column rather than derived, so a column moved into the
        // engine class fails here instead of passing because both sides moved.
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);

        _ = graph.Build(active).Initialise();

        bool[] hidden = active.HiddenStates;
        Assert.Equal(GanttTableSchema.Default.Columns.Count, hidden.Length);

        IReadOnlyList<string> authoring = ["Type", "Description", "Start", "Finish"];
        foreach (string name in authoring)
        {
            int index = IndexOf(name);
            Assert.False(hidden[index]);
            Assert.False(active.LockedWrites[index]);
        }

        // Duration is the one user-visible locked column (ADR-0029 D3).
        int duration = IndexOf("Duration");
        Assert.False(hidden[duration]);
        Assert.True(active.LockedWrites[duration]);

        foreach (string name in new[]
        {
            "Id", "ParentId", "SiblingOrder", "LaneId", "StackIndex", "StyleKey",
            "LabelPosition", "FillColour", "StrokeColour", "Visible", "SortOrder",
        })
        {
            int index = IndexOf(name);
            Assert.True(hidden[index]);
            Assert.True(active.LockedWrites[index]);
        }
    }

    /// <summary>
    /// The zero-based schema index of a column, failing loudly when the name is not
    /// in the schema — so a typo in a test above reads as a missing column rather
    /// than as an assertion about the wrong column.
    /// </summary>
    private static int IndexOf(string columnName)
    {
        for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            if (string.Equals(
                GanttTableSchema.Default.Columns[index].Name,
                columnName,
                StringComparison.Ordinal))
            {
                return index;
            }
        }

        throw new Xunit.Sdk.XunitException(
            $"Column '{columnName}' is not in the schema; the test's expectation is stale.");
    }

    [Fact]
    public void Initialise_materialises_the_catalogues_after_the_config_sheet()
    {
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);
        var writer = new Mock<IConfigCatalogueWriter>();
        writer.Setup(w => w.Write()).Returns(ConfigWriteOutcome.Ok());

        var outcome = graph.BuildWithWriter(active, writer.Object).Initialise();

        Assert.True(outcome.Succeeded);
        writer.Verify(w => w.Write(), Times.Once());
        active.VerifyAnchorName(ExpectedPlotAnchorRefersTo(), Times.Once());
    }

    [Theory]
    [InlineData(ProtectionGuardOutcome.SheetProtected)]
    [InlineData(ProtectionGuardOutcome.WorkbookStructureProtected)]
    public void Initialise_shared_guard_refusal_does_not_mutate(ProtectionGuardOutcome protection)
    {
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);
        var writer = new Mock<IConfigCatalogueWriter>();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(protection);

        var outcome = graph.BuildWithWriter(active, writer.Object, guard.Object).Initialise();

        Assert.Equal(WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected), outcome);
        guard.Verify(g => g.Query(), Times.Once);
        writer.Verify(w => w.Write(), Times.Never);
        active.VerifyTableCreated(Times.Never());
        graph.VerifySheetsAdded(Times.Never());
        Assert.Empty(active.WrittenValues);
    }

    /// <summary>
    /// A column already carrying the schema's visibility is not rewritten.
    /// </summary>
    /// <remarks>
    /// The guard that avoided this write compared an <c>object</c>-typed
    /// <c>Range.Hidden</c> against a <c>bool</c> with <c>!=</c>, which boxed the
    /// right-hand side and compared references. That is always true, so every column
    /// was written on every Initialise and an already-correct workbook was marked
    /// dirty for nothing. This asserts the skip, which only happens when the value is
    /// unwrapped to a real boolean first.
    /// <para>
    /// The end state is asserted rather than the write count, because the run ends in
    /// a rollback whose own visibility restore legitimately writes. What must hold is
    /// that the presentation step left every column where it found it.
    /// </para>
    /// </remarks>
    [Fact]
    public void Initialise_does_not_rewrite_a_column_that_already_has_the_schemas_visibility()
    {
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);

        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        active.SetExistingHiddenState(
            schema.Select((column, index) => (index, column.IsHidden)));

        _ = graph.Build(active).Initialise();

        // No column is written by the presentation step. The only writes that occur
        // are the rollback's unhide pass, and those are all `false` -- restoring
        // visibility, never hiding. The always-true guard wrote all 16 columns,
        // every one of them, whatever the host actually reported.
        Assert.DoesNotContain(active.HiddenWrites, written => written);
    }

    /// <summary>
    /// A refused Initialise leaves no column hidden.
    /// </summary>
    /// <remarks>
    /// Hiding a column is a <em>worksheet</em> change that outlives the table, so
    /// rolling back the table and the header while leaving the engine columns hidden
    /// would leave the user's sheet altered after a refusal the message calls a
    /// zero-mutation outcome. This asserts the visibility is restored too.
    /// </remarks>
    [Fact]
    public void A_refused_initialise_restores_the_visibility_it_applied()
    {
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);
        var writer = new Mock<IConfigCatalogueWriter>();
        _ = writer.Setup(w => w.Write())
            .Returns(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.TargetProtected));

        var outcome = graph.BuildWithWriter(active, writer.Object).Initialise();

        Assert.Equal(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected),
            outcome);

        // Every engine column was hidden by the presentation step, and the rollback
        // must have unhidden every one of them again.
        int hiddenColumns = GanttTableSchema.Default.Columns.Count(column => column.IsHidden);
        Assert.True(hiddenColumns > 0, "The schema must hide at least one column for this to mean anything.");
        Assert.Equal(
            hiddenColumns,
            active.HiddenWrites.Count(written => !written));
    }

    [Fact]
    public void Initialise_rolls_back_everything_when_the_catalogue_write_refuses()
    {
        // The writer's typed refusal (e.g. a protection race after the
        // read-only checks) leaves no header, no table, and no config
        // sheet: the refusal is transactional (AGENTS.md; R2.2 D4).
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);
        var writer = new Mock<IConfigCatalogueWriter>();
        writer.Setup(w => w.Write()).Returns(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.TargetProtected));

        var outcome = graph.BuildWithWriter(active, writer.Object).Initialise();

        Assert.Equal(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected),
            outcome);
        active.VerifyAnchorName(null, Times.Never());
    }

    [Fact]
    public void Initialise_writes_the_headers_in_the_exact_schema_order()
    {
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);

        _ = graph.Build(active).Initialise();

        var written = Assert.Single(active.WrittenValues);
        var values = Assert.IsType<object[,]>(written);
        Assert.Equal(1, values.GetLength(0));
        Assert.Equal(GanttTableSchema.Default.Columns.Count, values.GetLength(1));
        for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            Assert.Equal(GanttTableSchema.Default.Columns[index].Name, values[0, index]);
        }
    }

    [Fact]
    public void Initialise_suffixes_the_label_when_the_name_is_taken()
    {
        var active = BlankActiveSheet();
        var other = new WorksheetGraph(GanttWorkbookContract.GanttSheetLabel);
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active, other);
        graph.EnqueueCreated(config);

        var outcome = graph.Build(active).Initialise();

        Assert.True(outcome.Succeeded);
        Assert.Equal("Gantt Data (2)", outcome.SheetName);
        Assert.Equal("Gantt Data (2)", active.Name);
        Assert.Equal(GanttWorkbookContract.GanttSheetLabel, other.Name);
    }

    [Fact]
    public void Initialise_writes_the_sheet_scoped_plot_anchor_name()
    {
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);

        _ = graph.Build(active).Initialise();

        // 14 schema columns → the anchor is column 15 = "O"; the name is
        // sheet-scoped and refers to the post-rename label.
        active.VerifyAnchorName(ExpectedPlotAnchorRefersTo(), Times.Once());
    }

    [Fact]
    public void Initialise_creates_the_configuration_sheet_very_hidden()
    {
        var active = BlankActiveSheet();
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(config);

        _ = graph.Build(active).Initialise();

        Assert.Equal(GanttWorkbookContract.ConfigSheetName, config.Name);
        var visibility = Assert.Single(config.VisibleValues);
        Assert.Equal(Excel.XlSheetVisibility.xlSheetVeryHidden, visibility);
        Assert.Empty(config.WrittenValues);
        config.VerifyAnchorName(null, Times.Never());
        config.VerifyTableCreated(Times.Never());
        graph.VerifySheetsAdded(Times.Once());
    }

    /// <summary>A non-pristine active worksheet must be preserved on create.</summary>
    private enum NonPristineArtefact
    {
        /// <summary>Used range extends beyond A1.</summary>
        UsedRange,

        /// <summary>Cell content exists.</summary>
        CellContent,

        /// <summary>A drawing shape exists.</summary>
        Shape,

        /// <summary>A legacy note exists.</summary>
        Comment,

        /// <summary>A threaded comment exists.</summary>
        ThreadedComment,

        /// <summary>A worksheet-scoped name exists.</summary>
        SheetName,

        /// <summary>A workbook-scoped name exists.</summary>
        WorkbookName,

        /// <summary>A non-Gantt table exists.</summary>
        Table,

        /// <summary>A pivot table exists.</summary>
        PivotTable,

        /// <summary>A query table exists.</summary>
        QueryTable,

        /// <summary>A hyperlink exists.</summary>
        Hyperlink,
    }

    // ---------------------------------------------------------------------
    // Create path (non-blank active worksheet)
    // ---------------------------------------------------------------------

    [Fact]
    public void Initialise_creates_a_new_labelled_sheet_when_the_active_worksheet_has_content()
    {
        var active = new WorksheetGraph("Sheet1");
        active.NonEmptyCellCount = 3;
        active.WithExistingTable(null);
        var created = new WorksheetGraph("Sheet2");
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(created);
        graph.EnqueueCreated(config);

        var outcome = graph.Build(active).Initialise();

        Assert.True(outcome.Succeeded);
        Assert.Equal(WorkbookInitialisePath.CreatedNew, outcome.Path);
        Assert.Equal(GanttWorkbookContract.GanttSheetLabel, outcome.SheetName);
        Assert.Equal(GanttWorkbookContract.GanttSheetLabel, created.Name);
        created.VerifyTableCreated(Times.Once());
        Assert.Equal(new[] { GanttTableSchema.TableName }, created.AssignedTableNames);
        created.VerifyAnchorName(ExpectedPlotAnchorRefersTo(), Times.Once());
    }

    [Theory]
    [InlineData("UsedRange")]
    [InlineData("CellContent")]
    [InlineData("Shape")]
    [InlineData("Comment")]
    [InlineData("ThreadedComment")]
    [InlineData("SheetName")]
    [InlineData("WorkbookName")]
    [InlineData("Table")]
    [InlineData("PivotTable")]
    [InlineData("QueryTable")]
    [InlineData("Hyperlink")]
    public void Initialise_preserves_a_cell_empty_sheet_with_non_cell_state(
        string artefactName)
    {
        var artefact = Enum.Parse<NonPristineArtefact>(artefactName);
        var active = BlankActiveSheet();
        switch (artefact)
        {
            case NonPristineArtefact.UsedRange:
                active.UsedRangeAddress = "$B$1";
                break;
            case NonPristineArtefact.CellContent:
                active.NonEmptyCellCount = 1;
                break;
            case NonPristineArtefact.Shape:
                active.ShapeCount = 1;
                break;
            case NonPristineArtefact.Comment:
                active.CommentCount = 1;
                break;
            case NonPristineArtefact.ThreadedComment:
                active.ThreadedCommentCount = 1;
                break;
            case NonPristineArtefact.SheetName:
                active.SheetNameCount = 1;
                break;
            case NonPristineArtefact.WorkbookName:
                break;
            case NonPristineArtefact.Table:
                active.WithExistingTable("ForeignTable");
                break;
            case NonPristineArtefact.PivotTable:
                active.PivotTableCount = 1;
                break;
            case NonPristineArtefact.QueryTable:
                active.QueryTableCount = 1;
                break;
            case NonPristineArtefact.Hyperlink:
                active.HyperlinkCount = 1;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(artefactName));
        }

        var created = new WorksheetGraph("Sheet2");
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        if (artefact == NonPristineArtefact.WorkbookName)
        {
            graph.WorkbookNameCount = 1;
        }

        graph.EnqueueCreated(created);
        graph.EnqueueCreated(config);

        var outcome = graph.Build(active).Initialise();

        Assert.True(outcome.Succeeded);
        Assert.Equal(WorkbookInitialisePath.CreatedNew, outcome.Path);
        Assert.Equal("Sheet1", active.Name);
        Assert.Empty(active.WrittenValues);
        active.VerifyTableCreated(Times.Never());
        created.VerifyTableCreated(Times.Once());
        graph.VerifySheetsAdded(Times.Exactly(2));
    }

    [Fact]
    public void Initialise_leaves_the_non_blank_active_worksheet_untouched()
    {
        var active = new WorksheetGraph("Sheet1");
        active.NonEmptyCellCount = 3;
        active.WithExistingTable(null);
        var created = new WorksheetGraph("Sheet2");
        var config = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName);
        var graph = new WorkbookGraph(active);
        graph.EnqueueCreated(created);
        graph.EnqueueCreated(config);

        _ = graph.Build(active).Initialise();

        Assert.Equal("Sheet1", active.Name);
        Assert.Empty(active.WrittenValues);
        active.VerifyAnchorName(null, Times.Never());
        Assert.Empty(active.VisibleValues);
        Assert.Empty(active.AssignedTableNames);
        active.VerifyTableCreated(Times.Never());

        // Exactly two structural additions: the new Gantt sheet and the
        // config sheet.
        graph.VerifySheetsAdded(Times.Exactly(2));
    }

    // ---------------------------------------------------------------------
    // Refusal positives: each validator's bad input, typed refusal, no mutation
    // ---------------------------------------------------------------------

    [Fact]
    public void Initialise_reports_no_active_workbook_without_an_application_object()
        => Assert.Equal(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook),
            new ExcelWorkbookInitialiser(null).Initialise());

    [Fact]
    public void Initialise_degrades_to_no_active_workbook_for_a_foreign_object()
        => Assert.Equal(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook),
            new ExcelWorkbookInitialiser("not an Excel application").Initialise());

    [Fact]
    public void Initialise_reports_no_active_workbook_when_excel_has_no_active_workbook()
    {
        var application = new Mock<Excel.Application>();
        application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        Assert.Equal(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook),
            new ExcelWorkbookInitialiser(application.Object).Initialise());
    }

    [Fact]
    public void Initialise_refuses_when_the_configuration_sheet_already_exists()
    {
        // The active worksheet already carries the configuration sheet's name,
        // which is exactly how a second helper sheet would be created — the
        // refusal must fire before any mutation.
        var active = new WorksheetGraph(GanttWorkbookContract.ConfigSheetName)
        {
            NonEmptyCellCount = 0,
        };
        active.WithExistingTable(null);
        var graph = new WorkbookGraph(active);

        var outcome = graph.Build(active).Initialise();

        Assert.Equal(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.ConfigSheetExists),
            outcome);
        Assert.Equal(GanttWorkbookContract.ConfigSheetName, active.Name);
        Assert.Empty(active.WrittenValues);
        graph.VerifySheetsAdded(Times.Never());
        active.VerifyTableCreated(Times.Never());
    }

    [Fact]
    public void Initialise_refuses_when_an_otherwise_blank_target_already_contains_the_table()
    {
        var active = BlankActiveSheet();
        active.WithExistingTable(GanttTableSchema.TableName);
        var graph = new WorkbookGraph(active);

        var outcome = graph.Build(active).Initialise();

        Assert.Equal(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TableExists),
            outcome);
        Assert.Equal("Sheet1", active.Name);
        Assert.Empty(active.WrittenValues);
        graph.VerifySheetsAdded(Times.Never());
    }

    [Fact]
    public void Initialise_refuses_when_the_target_rename_is_blocked()
    {
        var active = BlankActiveSheet();
        // CA2201: the production catch clause is COMException-typed, so the
        // fault injection must construct exactly that type. This is a test
        // input, not a runtime failure signal.
#pragma warning disable CA2201
        active.Worksheet.SetupSet(w => w.Name = It.IsAny<string>())
            .Throws(new COMException("The sheet is protected."));
#pragma warning restore CA2201
        var graph = new WorkbookGraph(active);

        var outcome = graph.Build(active).Initialise();

        Assert.Equal(
            WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected),
            outcome);
        Assert.Equal("Sheet1", active.Name);
        Assert.Empty(active.WrittenValues);
        graph.VerifySheetsAdded(Times.Never());
        active.VerifyTableCreated(Times.Never());
    }

    /// <summary>A single <c>false</c> setting (the neutral-appearance contract).</summary>
    private static readonly bool[] FalseSettings = [false];

    /// <summary>A blank active worksheet named <c>Sheet1</c> with no existing table.</summary>
    /// <returns>The configured worksheet graph.</returns>
    private static WorksheetGraph BlankActiveSheet()
    {
        var active = new WorksheetGraph("Sheet1")
        {
            NonEmptyCellCount = 0,
        };
        active.WithExistingTable(null);
        return active;
    }
}
