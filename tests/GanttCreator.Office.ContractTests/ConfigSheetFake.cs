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
            // ListObject.Delete unregisters the table, so the writer's
            // delete-then-recreate regeneration does not accumulate tables.
            _ = TableMock.Setup(t => t.Delete()).Callback(() => owner.Remove(this));
        }

        /// <summary>
        /// The registered table name (the public face; the mock getter returns
        /// the same value via <see cref="_name"/>).
        /// </summary>
        public string Name { get; }

        public string[] Headers { get; }

        public List<object?[]> Body { get; }

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
        _tables.Add(table);
        return table.TableMock.Object;
    }

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
        }

        return range.Object;
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
    /// table's <c>DataBodyRange</c> (body rows only, no header row).
    /// </summary>
    /// <param name="body">The body rows.</param>
    /// <returns>The payload.</returns>
    public static object ToBodyMatrix(List<object?[]> body)
    {
        var rows = body.Count;
        var columns = body.Count > 0 ? body[0].Length : 0;
        if (rows == 0 || columns == 0)
        {
#pragma warning disable CA1814 // COM interop requires a rectangular SAFEARRAY; jagged arrays do not marshal.
            return new object[0, 0];
#pragma warning restore CA1814
        }

#pragma warning disable CA1814 // COM interop requires a rectangular SAFEARRAY; jagged arrays do not marshal.
        var matrix = new object[rows, columns];
#pragma warning restore CA1814
        for (var row = 0; row < rows; row++)
        {
            for (var column = 0; column < columns; column++)
            {
                matrix[row, column] = body[row][column] ?? string.Empty;
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
