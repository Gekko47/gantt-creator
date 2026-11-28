using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Shared Moq graph over the <c>_GanttCreatorConfig</c> worksheet for the
/// catalogue contract tests. Models the config sheet as named mock ranges:
/// <list type="bullet">
/// <item>an anchor map (address → range mock capturing <c>Value2</c>
/// writes and answering resizes);</item>
/// <item>a table registry (name → <see cref="ConfigTableFake"/> holding
/// headers and an in-memory body matrix);</item>
/// <item>protection flags and the sheet name.</item>
/// </list>
/// Writer calls register new tables (emulating <c>ListObjects.Add</c>) and
/// reader calls replay the captures, so the round-trip
/// write → read → equals-Core-catalogues test runs through the real
/// production <c>Write()</c>/<c>Read()</c> code with only the COM seams
/// substituted. A table the test builds with a live matrix behaves like an
/// existing table (the D4 regeneration path); the emitted header row is
/// prepended to the body matrix to emulate the real extent payload.
/// </summary>
internal sealed class ConfigSheetFake
{
    /// <summary>
    /// One registered table: headers plus an in-memory body matrix, and the
    /// mocks the writer/reader see through the seams.
    /// </summary>
    internal sealed class TableFake
    {
        /// <summary>
        /// Backing variable for the mock's <c>Name</c> getter; updated by the
        /// setter callback so the writer's <c>table.Name = tableName</c>
        /// assignment is visible to the reader's <c>FindTable</c>.
        /// </summary>
        private string _name;

        public TableFake(string name, string[] headers, List<object?[]> body, List<TableFake> owner)
        {
            _name = name;
            Name = name;
            Headers = headers;
            Body = body;
            ColumnsMock = new Mock<Excel.ListColumns>();
            for (var index = 0; index < headers.Length; index++)
            {
                var column = new Mock<Excel.ListColumn>();
                var columnIndex = index;
                // Live read: a damaged-header fixture mutates Headers[i] after
                // construction, and the reader must observe the new name.
                _ = column.SetupGet(c => c.Name).Returns(() => Headers[columnIndex]);
                Columns.Add(column);
            }

            _ = ColumnsMock.SetupGet(c => c.Count).Returns(() => headers.Length);
            _ = BodyMock.SetupGet(r => r.Value2).Returns(() => ToBodyMatrix(Body));
            _ = TableMock.SetupGet(t => t.Name).Returns(() => _name);
            _ = TableMock.SetupSet(t => t.Name = It.IsAny<string>())
                .Callback<string>(value => _name = value);
            _ = TableMock.SetupGet(t => t.ListColumns).Returns(ColumnsMock.Object);
            _ = TableMock.SetupGet(t => t.DataBodyRange).Returns(BodyMock.Object);
            // ListObject.Delete unregisters the table, so the catalogue write's
            // delete-then-recreate regeneration (and the settings fallback)
            // does not accumulate tables. The in-place settings path never
            // reaches this callback, which is what its contract test asserts.
            _ = TableMock.Setup(t => t.Delete()).Callback(() => owner.Remove(this));
        }

        /// <summary>
        /// The registered table name (the public face; the mock getter returns
        /// the same value via <see cref="_name"/>).
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// The table's current name, following any <c>Name</c> assignment the
        /// production code made. <see cref="Name"/> is the name the fake was
        /// constructed with, which is only the contract name after the writer
        /// has assigned it — so a test that inspects the sheet after a rewrite
        /// must read this instead.
        /// </summary>
        public string CurrentName => _name;

        public string[] Headers { get; }

        public List<object?[]> Body { get; }

        /// <summary>
        /// The table's top-left origin, one-based. Set when the writer's
        /// <c>AddTable</c> seam registers the table through the extent the
        /// resize seam produced, so the anchor check reads the same address
        /// the production call used. A table built directly by a fixture keeps
        /// the default <c>(0, 0)</c>, which is off every contract anchor.
        /// </summary>
        public (int Row, int Column) Origin { get; private set; }

        /// <summary>Moves the table's recorded origin, for the off-anchor case.</summary>
        /// <param name="row">The one-based row.</param>
        /// <param name="column">The one-based column.</param>
        public void MoveTo(int row, int column) => Origin = (row, column);

        /// <summary>
        /// Replaces the body rows from the rows the writer handed the seam. Each
        /// row is copied, so a fixture that keeps its own row arrays cannot
        /// alias the table's stored state.
        /// </summary>
        /// <param name="rows">The body rows, in contract column order.</param>
        public void ReplaceRows(IReadOnlyList<object?[]> rows)
        {
            Body.Clear();
            foreach (object?[] row in rows)
            {
                Body.Add([.. row]);
            }
        }

        public List<Mock<Excel.ListColumn>> Columns { get; } = [];

        public Mock<Excel.ListObject> TableMock { get; } = new();

        public Mock<Excel.ListColumns> ColumnsMock { get; }

        public Mock<Excel.Range> BodyMock { get; } = new();

        public Mock<Excel.ListColumn> ColumnMockAt(int index) => Columns[index - 1];

        public Excel.ListColumn ColumnAt(int index) => Columns[index - 1].Object;
    }

    private readonly List<TableFake> _tables = [];

    private readonly Mock<Excel.Worksheet> _worksheet = new();

    private readonly Mock<Excel.ListObjects> _listObjects = new();

    private readonly Dictionary<string, Mock<Excel.Range>> _cellRanges = new(StringComparer.Ordinal);

    private int _failAddAt = -1;

    private int _addCallCount;

    private bool _failBodyWrite;

    private readonly List<(Excel.Range Range, string Address)> _cellAddresses = [];

    private readonly List<(Excel.Range Range, string Address)> _extentAddresses = [];

    private readonly List<Excel.Range> _clearedRanges = [];

    /// <summary>
    /// Initialises the fake with the given protection state and config
    /// sheet identity.
    /// </summary>
    /// <param name="protectContents">The mocked <c>Worksheet.ProtectContents</c>.</param>
    /// <param name="sheetName">The mocked worksheet name.</param>
    public ConfigSheetFake(bool protectContents = false, string sheetName = "_GanttCreatorConfig")
    {
        _ = _worksheet.SetupGet(w => w.Name).Returns(sheetName);
        _ = _worksheet.SetupGet(w => w.ProtectContents).Returns(protectContents);
        _ = _worksheet.SetupGet(w => w.ListObjects).Returns(_listObjects.Object);
        _ = _listObjects.SetupGet(l => l.Count).Returns(() => _tables.Count);
    }

    /// <summary>
    /// Makes the Nth <c>ListObjects.Add</c> call fail, so a test can drive a
    /// failure partway through the writer's five-table sequence.
    /// </summary>
    /// <param name="oneBasedIndex">
    /// The one-based add-call index to fail, counted from this call: arming it
    /// resets the add counter so a fixture that already ran a successful write
    /// is not silently skipped.
    /// </param>
    public void FailAddAt(int oneBasedIndex)
    {
        _failAddAt = oneBasedIndex;
        _addCallCount = 0;
    }

    /// <summary>
    /// Makes the next in-place settings body write mutate the body and then
    /// fail, so a test can prove the rollback restores the prior rows. The
    /// mutation before the throw models a host that refused the call after
    /// applying it, which is the case the pre-write snapshot exists for.
    /// </summary>
    public void FailBodyWrite() => _failBodyWrite = true;

    /// <summary>
    /// The cell ranges whose contents were cleared through the writer's
    /// rollback seam, in call order.
    /// </summary>
    public List<Excel.Range> ClearedRanges => _clearedRanges;

    /// <summary>
    /// The names of the columns whose body the writer formatted as Text, in
    /// call order. The settings table's Value column is the only one: its
    /// free text (the plot dates) must survive Excel's locale coercion.
    /// </summary>
    public List<string> TextFormattedColumns { get; } = [];

    /// <summary>
    /// The ordered evidence of the settings write path: one
    /// <c>format:<column></c> entry per Text-format call, and
    /// <c>body:write</c> when the body payload is applied. A test asserts the
    /// format precedes the write, which is the ordering that keeps free-text
    /// dates such as <c>10/01/2025</c> verbatim.
    /// </summary>
    public List<string> Operations { get; } = [];

    /// <summary>How many times a table was created through the add seam.</summary>
    public int AddCallCount => _addCallCount;

    /// <summary>Gets the mocked configuration worksheet.</summary>
    public Excel.Worksheet Worksheet => _worksheet.Object;

    /// <summary>Gets the mocked list-objects collection.</summary>
    public Excel.ListObjects ListObjects => _listObjects.Object;

    /// <summary>Gets the registered tables in registration order.</summary>
    public List<ConfigSheetFake.TableFake> Tables => _tables;

    /// <summary>
    /// Emulates <c>ListObjects.Add</c>: parses the extent payload, reads
    /// the header row, registers the table, and returns it. Registration
    /// order is creation order.
    /// </summary>
    /// <param name="extent">The extent range carrying the payload.</param>
    /// <returns>The registered table object.</returns>
    public Excel.ListObject Add(Excel.Range extent)
    {
        _addCallCount++;
        if (_addCallCount == _failAddAt)
        {
            // An InvalidOperationException rather than a COMException, because
            // that is how a refused COM call is reported once it crosses the
            // interop boundary into managed code, and because CA2201 reserves
            // the runtime's own exception types. The writer catches broadly for
            // the same reason: the host chooses which of the two it raises.
            throw new InvalidOperationException("The host refused the table write.");
        }

        var payload = (object[,])extent.Value2;
        var columns = payload.GetLength(1);
        var headers = new string[columns];
        for (var column = 0; column < columns; column++)
        {
            headers[column] = Convert.ToString(payload[0, column], System.Globalization.CultureInfo.InvariantCulture)
                ?? string.Empty;
        }

        var body = new List<object?[]>();
        for (var row = 1; row < payload.GetLength(0); row++)
        {
            var cells = new object?[columns];
            for (var column = 0; column < columns; column++)
            {
                cells[column] = payload[row, column];
            }

            body.Add(cells);
        }

        var table = new TableFake($"pending-{_tables.Count + 1}", headers, body, _tables);
        if (TryParseAddress(AddressOf(extent), out var originRow, out var originColumn))
        {
            table.MoveTo(originRow, originColumn);
        }

        _tables.Add(table);
        return table.TableMock.Object;
    }

    /// <summary>
    /// Writes the settings body in place, through the writer's
    /// <c>SetBodyValues</c> seam. Mutates the owning table's body rows so the
    /// next read sees the write, exactly as a live <c>Value2</c> assignment
    /// would. When <see cref="FailBodyWrite"/> is armed it applies the write
    /// and then refuses, which is what the rollback must recover from.
    /// </summary>
    /// <param name="body">The table's data body range.</param>
    /// <param name="rows">The body rows the writer assigned.</param>
    public void WriteBody(Excel.Range body, IReadOnlyList<object?[]> rows)
    {
        Operations.Add("body:write");
        TableFake match = _tables.Single(candidate => ReferenceEquals(candidate.BodyMock.Object, body));
        match.ReplaceRows(rows);
        if (_failBodyWrite)
        {
            // Same exception shape as the add seam: a refused COM call
            // crossing the interop boundary, never an ArgumentException.
            throw new InvalidOperationException("The host refused the in-place settings write.");
        }
    }

    /// <summary>
    /// Returns a table's recorded origin, one-based.
    /// </summary>
    /// <param name="table">The table object.</param>
    /// <returns>The one-based row and column.</returns>
    public (int Row, int Column) OriginOf(Excel.ListObject table) =>
        _tables.Single(candidate => ReferenceEquals(candidate.TableMock.Object, table)).Origin;

    /// <summary>
    /// One shared cell-range mock per anchor address. The cell mock carries
    /// no behaviour: resizing goes through the
    /// <c>GetResizedRange</c>/<c>Resize</c> seam (the COM <c>Resize</c>
    /// indexer cannot appear in a Moq expression tree, CS0855).
    /// </summary>
    /// <param name="address">The anchor address (e.g. <c>"A1"</c>).</param>
    /// <returns>The range mock.</returns>
    public Excel.Range CellAt(string address)
    {
        if (!_cellRanges.TryGetValue(address, out var range))
        {
            range = new Mock<Excel.Range>();
            _cellRanges[address] = range;
            _cellAddresses.Add((range.Object, address));
        }

        return range.Object;
    }

    /// <summary>
    /// Creates the extent the writer writes into and records the anchor
    /// address it was resized from, so the table registered by
    /// <see cref="Add"/> carries the origin the production call used. The
    /// address is what the writer's anchor check reads back, because the
    /// COM <c>Resize</c> indexer cannot appear in a Moq expression tree.
    /// </summary>
    /// <param name="source">The anchor cell range.</param>
    /// <param name="rows">The row count.</param>
    /// <param name="columns">The column count.</param>
    /// <returns>The extent mock.</returns>
    public Excel.Range ResizeFrom(Excel.Range source, int rows, int columns)
    {
        string address = _cellAddresses
            .Where(entry => ReferenceEquals(entry.Range, source))
            .Select(entry => entry.Address)
            .FirstOrDefault() ?? string.Empty;
        Excel.Range extent = CreateExtent(rows, columns);
        if (address.Length > 0)
        {
            _extentAddresses.Add((extent, address));
        }

        return extent;
    }

    /// <summary>The anchor address a resize produced the given extent from.</summary>
    /// <param name="extent">The extent mock.</param>
    /// <returns>The address, or an empty string when the extent has none.</returns>
    private string AddressOf(Excel.Range extent) =>
        _extentAddresses
            .Where(entry => ReferenceEquals(entry.Range, extent))
            .Select(entry => entry.Address)
            .FirstOrDefault() ?? string.Empty;

    /// <summary>Parses an A1 address into a one-based row and column.</summary>
    /// <param name="address">The address (e.g. <c>"A70"</c>).</param>
    /// <param name="row">The one-based row.</param>
    /// <param name="column">The one-based column.</param>
    /// <returns><see langword="true"/> when the address parses.</returns>
    private static bool TryParseAddress(string address, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (address.Length == 0)
        {
            return false;
        }

        int index = 0;
        int parsedColumn = 0;
        while (index < address.Length && char.IsLetter(address[index]))
        {
            parsedColumn = (parsedColumn * 26) + (char.ToUpperInvariant(address[index]) - 'A' + 1);
            index++;
        }

        if (index == 0 || index == address.Length)
        {
            return false;
        }

        if (!int.TryParse(address[index..], NumberStyles.None, CultureInfo.InvariantCulture, out int parsedRow)
            || parsedRow <= 0
            || parsedColumn <= 0)
        {
            return false;
        }

        row = parsedRow;
        column = parsedColumn;
        return true;
    }

    /// <summary>
    /// One extent mock per (anchor, rows, columns) shape; the
    /// <c>Value2</c> setter stores the payload, the getter replays it, so
    /// the writer's own write is immediately readable.
    /// </summary>
    /// <param name="anchor">The anchor address.</param>
    /// <param name="rows">The row count.</param>
    /// <param name="columns">The column count.</param>
    /// <returns>The extent mock.</returns>
    public static Excel.Range Extent(string anchor, int rows, int columns)
    {
        var extent = new Mock<Excel.Range>();
        object? payload = null;
        _ = extent.SetupSet(r => r.Value2 = It.IsAny<object>()).Callback<object>(value => payload = value);
        _ = extent.SetupGet(r => r.Value2).Returns(() => payload!);
        return extent.Object;
    }

    /// <summary>
    /// Returns the registered table at the one-based index (the
    /// <c>GetTableAt</c> seam target).
    /// </summary>
    /// <param name="index">The one-based table index.</param>
    /// <returns>The table object.</returns>
    public Excel.ListObject TableAt(int index) => _tables[index - 1].TableMock.Object;

    /// <summary>
    /// Returns the list column at the one-based index of the given
    /// columns collection (the <c>GetColumnAt</c> seam target).
    /// </summary>
    /// <param name="columns">The columns collection.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The column object.</returns>
    public Excel.ListColumn ColumnAt(Excel.ListColumns columns, int index)
    {
        var match = _tables.Single(candidate => ReferenceEquals(candidate.ColumnsMock.Object, columns));
        return match.ColumnAt(index);
    }

    /// <summary>
    /// Returns the body range of the table the given object belongs to.
    /// </summary>
    /// <param name="table">The table object.</param>
    /// <returns>The body range, or <see langword="null"/> for an empty body.</returns>
    public Excel.Range? BodyOf(Excel.ListObject table)
    {
        var match = _tables.Single(candidate => ReferenceEquals(candidate.TableMock.Object, table));
        return match.Body.Count == 0 ? null : match.BodyMock.Object;
    }

    /// <summary>
    /// Returns the bulk <c>Value2</c> payload of the given range.
    /// </summary>
    /// <param name="body">The range.</param>
    /// <returns>The payload.</returns>
    public static object? ValuesOf(Excel.Range body) => body.Value2;

    /// <summary>
    /// Creates a payload-carrying extent mock for the writer's resize seam:
    /// the <c>Value2</c> setter stores the payload, the getter replays it.
    /// </summary>
    /// <param name="rows">The row count.</param>
    /// <param name="columns">The column count.</param>
    /// <returns>The extent mock.</returns>
    public static Excel.Range CreateExtent(int rows, int columns)
    {
        var extent = new Mock<Excel.Range>();
        object? payload = null;
        _ = extent.SetupSet(r => r.Value2 = It.IsAny<object>()).Callback<object>(value => payload = value);
        _ = extent.SetupGet(r => r.Value2).Returns(() => payload!);
        return extent.Object;
    }

    /// <summary>
    /// Converts body rows into the <c>Value2</c> payload Excel returns for a
    /// table's <c>DataBodyRange</c> (body rows only, no header row): a
    /// rectangular SAFEARRAY with one-based lower bounds, matching live
    /// Excel.
    /// </summary>
    /// <param name="body">The body rows.</param>
    /// <returns>The payload.</returns>
    public static object ToBodyMatrix(List<object?[]> body)
    {
        var rows = body.Count;
        var columns = body.Count > 0 ? body[0].Length : 0;
        if (rows == 0 || columns == 0)
        {
            return Array.CreateInstance(typeof(object), [0, 0], [1, 1]);
        }

        // Live DataBodyRange.Value2 is indexed from 1 (the R2.7 live office
        // gate crashed on matrix[0, ...] in a 0-based-only reader loop, and
        // ExcelGanttTableReader's contract fake already pins the same
        // one-based shape). Serving one-based matrices here keeps that whole
        // bug class a contract-test failure instead of a live-only failure.
        var matrix = (object[,])Array.CreateInstance(typeof(object), [rows, columns], [1, 1]);
        for (var row = 1; row <= rows; row++)
        {
            for (var column = 1; column <= columns; column++)
            {
                matrix[row, column] = body[row - 1][column - 1] ?? string.Empty;
            }
        }

        return matrix;
    }

    /// <summary>
    /// Converts headers plus body rows into the one-based <c>Value2</c>
    /// payload Excel returns for a table body (header row first).
    /// </summary>
    /// <param name="headers">The header row.</param>
    /// <param name="body">The body rows.</param>
    /// <returns>The payload.</returns>
    public static object ToMatrix(string[] headers, List<object?[]> body)
    {
        var rows = body.Count + 1;
        var columns = headers.Length;
#pragma warning disable CA1814 // COM interop requires a rectangular SAFEARRAY; jagged arrays do not marshal.
        var matrix = new object[rows, columns];
#pragma warning restore CA1814
        for (var column = 0; column < columns; column++)
        {
            matrix[0, column] = headers[column];
        }

        for (var row = 0; row < body.Count; row++)
        {
            for (var column = 0; column < columns; column++)
            {
                matrix[row + 1, column] = body[row][column] ?? string.Empty;
            }
        }

        return matrix;
    }
}
