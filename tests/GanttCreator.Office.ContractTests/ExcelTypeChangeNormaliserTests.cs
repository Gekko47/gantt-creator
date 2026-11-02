using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelTypeChangeNormaliser"/>.
/// </summary>
/// <remarks>
/// The Excel graph is mocked and every COM parameterised property the adapter
/// uses is behind an <c>internal virtual</c> seam (<c>GetSheetAt</c>,
/// <c>GetTableAt</c>, <c>GetColumnAt</c>, <c>GetTableBody</c>,
/// <c>GetBodyValues</c>, <c>GetCell</c>), so the whole normalisation rule set runs
/// on the test host. Type facts are read from
/// <see cref="EntityTypeCatalog"/> rather than hard-coded, so a catalogue change
/// that invalidates a test's premise fails the premise assertion instead of
/// silently testing nothing.
/// </remarks>
public sealed class ExcelTypeChangeNormaliserTests
{
    private static readonly string[] _allColumns =
    [
        "Type",
        "StyleKey",
        "LabelPosition",
        "DateLabelPosition",
        "FillColour",
        "StrokeColour",
        "ParentId",
        "LaneId",
        "StackIndex",
        "SiblingOrder",
    ];

    [Fact]
    public void A_row_whose_values_are_all_legal_for_its_type_is_left_alone()
    {
        GanttEntityType type = FirstType(definition => definition.Kind == EntityKind.Span);
        EntityTypeDefinition definition = Definition(type);
        Fixture fixture = new(_allColumns);
        fixture.Row(
            ("Type", definition.DisplayName),
            ("StyleKey", "UserStyle"),
            ("LabelPosition", definition.AllowedLabelPositions.First().ToString()),
            ("FillColour", HasFill(definition) ? "#FFFFFF" : null),
            ("StrokeColour", HasStroke(definition) ? "#000000" : null),
            ("ParentId", EntityHierarchyCatalog.MayBeChild(type) ? "PARENT" : null),
            ("LaneId", UsesLane(definition) ? "LANE" : null),
            ("StackIndex", UsesLane(definition) ? "1" : null),
            ("SiblingOrder", EntityHierarchyCatalog.MayBeChild(type) ? "1" : null));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void A_type_that_can_use_none_of_the_optional_values_has_all_of_them_cleared()
    {
        GanttEntityType type = FirstType(candidate =>
            candidate.ColourCapability == EntityColourCapability.None &&
            !EntityHierarchyCatalog.MayBeChild(candidate.Type) &&
            candidate.Kind is not (EntityKind.Span or EntityKind.Milestone));
        EntityTypeDefinition definition = Definition(type);

        // The premise: this type really does refuse every optional column.
        Assert.False(definition.RequiresStyleKey);
        Assert.False(HasFill(definition));
        Assert.False(HasStroke(definition));
        Assert.False(UsesLane(definition));

        Fixture fixture = new(_allColumns);
        fixture.Row(
            ("Type", definition.DisplayName),
            ("LabelPosition", DisallowedLabelPosition(type)),
            ("DateLabelPosition", DisallowedDateLabelPosition(type)),
            ("FillColour", "#FFFFFF"),
            ("StrokeColour", "#000000"),
            ("ParentId", "PARENT"),
            ("LaneId", "LANE"),
            ("StackIndex", "1"),
            ("SiblingOrder", "1"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(8, outcome.CellsWritten);
        foreach (string column in new[]
                 {
                     "LabelPosition", "DateLabelPosition", "FillColour", "StrokeColour",
                     "ParentId", "LaneId", "StackIndex", "SiblingOrder",
                 })
        {
            Assert.Equal(string.Empty, fixture.WriteTo(1, column));
        }
    }

    [Fact]
    public void A_type_that_requires_a_style_key_receives_its_default_when_the_cell_is_blank()
    {
        GanttEntityType type = FirstType(candidate =>
            candidate.RequiresStyleKey && candidate.DefaultStyleKey.Length > 0);
        EntityTypeDefinition definition = Definition(type);
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", definition.DisplayName), ("StyleKey", " "));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(1, outcome.CellsWritten);
        Assert.Equal(definition.DefaultStyleKey, fixture.WriteTo(1, "StyleKey"));
    }

    [Fact]
    public void A_type_that_requires_a_style_key_keeps_the_style_key_the_user_chose()
    {
        GanttEntityType type = FirstType(candidate => candidate.RequiresStyleKey);
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", Definition(type).DisplayName), ("StyleKey", "Chosen"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void An_unknown_type_is_skipped()
    {
        Fixture fixture = new(_allColumns);
        fixture.Row(
            ("Type", "Not A Type"),
            ("FillColour", "#FFFFFF"),
            ("ParentId", "PARENT"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void A_blank_type_is_skipped()
    {
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", "  "), ("FillColour", "#FFFFFF"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void A_table_without_a_type_column_is_never_normalised()
    {
        Fixture fixture = new("FillColour", "ParentId");
        fixture.Row(("FillColour", "#FFFFFF"), ("ParentId", "PARENT"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void A_missing_optional_column_is_simply_not_written()
    {
        GanttEntityType type = FirstType(candidate =>
            !candidate.RequiresStyleKey &&
            (candidate.ColourCapability & EntityColourCapability.Fill) == 0);
        Fixture fixture = new("Type", "FillColour");
        fixture.Row(("Type", Definition(type).DisplayName), ("FillColour", "#FFFFFF"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(1, outcome.CellsWritten);
        Assert.Equal(string.Empty, fixture.WriteTo(1, "FillColour"));
    }

    [Fact]
    public void An_unparseable_label_position_is_left_alone()
    {
        GanttEntityType type = FirstType(candidate => candidate.AllowedLabelPositions.Count > 0);
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", Definition(type).DisplayName), ("LabelPosition", "NotAPosition"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void An_undefined_label_position_number_is_left_alone()
    {
        GanttEntityType type = FirstType(candidate => candidate.AllowedLabelPositions.Count > 0);
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", Definition(type).DisplayName), ("LabelPosition", "9999"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void An_unparseable_date_label_position_is_left_alone()
    {
        GanttEntityType type = FirstType(_ => true);
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", Definition(type).DisplayName), ("DateLabelPosition", "NotAPosition"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void Every_row_is_normalised_against_its_own_type()
    {
        GanttEntityType cleared = FirstType(candidate =>
            candidate.ColourCapability == EntityColourCapability.None &&
            !EntityHierarchyCatalog.MayBeChild(candidate.Type) &&
            candidate.Kind is not (EntityKind.Span or EntityKind.Milestone));
        GanttEntityType lane = FirstType(candidate =>
            candidate.Kind == EntityKind.Span &&
            (candidate.ColourCapability & EntityColourCapability.Fill) != 0);

        Fixture fixture = new(_allColumns);
        fixture.Row(
            ("Type", Definition(cleared).DisplayName),
            ("FillColour", "#FFFFFF"),
            ("LaneId", "LANE"));
        fixture.Row(
            ("Type", Definition(lane).DisplayName),
            ("FillColour", "#FFFFFF"),
            ("LaneId", "LANE"),
            ("StackIndex", "2"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        // Row 1 clears FillColour and LaneId; row 2 keeps both because its type
        // supports a fill and a lane.
        Assert.Equal(2, outcome.CellsWritten);
        Assert.Equal(string.Empty, fixture.WriteTo(1, "FillColour"));
        Assert.Equal(string.Empty, fixture.WriteTo(1, "LaneId"));
        Assert.Null(fixture.WriteTo(2, "FillColour"));
        Assert.Null(fixture.WriteTo(2, "LaneId"));
    }

    [Fact]
    public void No_application_is_a_no_op()
    {
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", "Task"), ("FillColour", "#FFFFFF"));

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise(applicationPresent: false);

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void No_active_workbook_is_a_no_op()
    {
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", "Task"), ("FillColour", "#FFFFFF"));
        fixture.ForgetWorkbook();

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void A_protected_target_sheet_is_a_no_op()
    {
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", "Task"), ("FillColour", "#FFFFFF"));
        fixture.ProtectTarget();

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void A_workbook_without_the_data_table_is_a_no_op()
    {
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", "Task"), ("FillColour", "#FFFFFF"));
        fixture.RemoveTables();

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void A_table_without_a_body_is_a_no_op()
    {
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", "Task"), ("FillColour", "#FFFFFF"));
        fixture.RemoveBody();

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    [Fact]
    public void A_body_payload_that_is_not_a_matrix_produces_no_rows()
    {
        Fixture fixture = new(_allColumns);
        fixture.Row(("Type", "Task"), ("FillColour", "#FFFFFF"));
        fixture.BodyPayloadOverride = "not a matrix";

        GanttTypeChangeNormaliseOutcome outcome = fixture.Normalise();

        Assert.Equal(0, outcome.CellsWritten);
        Assert.Empty(fixture.Writes);
    }

    private static EntityTypeDefinition Definition(GanttEntityType type)
    {
        EntityTypeDefinition? definition = EntityTypeCatalog.GetDefinition(type);
        Assert.NotNull(definition);
        return definition;
    }

    private static bool HasFill(EntityTypeDefinition definition) =>
        (definition.ColourCapability & EntityColourCapability.Fill) != 0;

    private static bool HasStroke(EntityTypeDefinition definition) =>
        (definition.ColourCapability & EntityColourCapability.Stroke) != 0;

    private static bool UsesLane(EntityTypeDefinition definition) =>
        definition.Kind is EntityKind.Span or EntityKind.Milestone;

    private static GanttEntityType FirstType(Func<EntityTypeDefinition, bool> predicate) =>
        Enum.GetValues<GanttEntityType>()
            .First(type => predicate(Definition(type)));

    private static string DisallowedLabelPosition(GanttEntityType type)
    {
        EntityTypeDefinition definition = Definition(type);
        return Enum.GetValues<GanttLabelPosition>()
            .First(position => !definition.AllowedLabelPositions.Contains(position))
            .ToString();
    }

    private static string DisallowedDateLabelPosition(GanttEntityType type)
    {
        IReadOnlyList<GanttDateLabelPosition> allowed =
            EntityTypeCatalog.GetAllowedDateLabelPositions(type);
        return Enum.GetValues<GanttDateLabelPosition>()
            .First(position => !allowed.Contains(position))
            .ToString();
    }

    /// <summary>
    /// The adapter with all six COM parameterised-property seams substituted.
    /// </summary>
    private sealed class TestableNormaliser(
        object? application,
        IWorksheetProtectionGuard guard,
        Fixture fixture)
        : ExcelTypeChangeNormaliser(application, guard)
    {
        internal override Excel.Worksheet GetSheetAt(Excel.Sheets sheets, int index) =>
            fixture.WorksheetObject;

        internal override Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) =>
            fixture.TableObject;

        internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) =>
            fixture.ColumnAt(index);

        internal override Excel.Range? GetTableBody(Excel.ListObject table) =>
            fixture.BodyObject;

        internal override object? GetBodyValues(Excel.Range body) => fixture.BodyPayload;

        internal override Excel.Range GetCell(Excel.Range body, int row, int column) =>
            fixture.CellAt(row, column);
    }

    /// <summary>
    /// The mocked Excel graph, the column schema and the write recorder.
    /// </summary>
    private sealed class Fixture
    {
        private readonly string[] _columns;
        private readonly List<object?[]> _rows = [];
        private readonly Dictionary<(int Row, int Column), Excel.Range> _cells = [];

        public Fixture(params string[] columns)
        {
            _columns = columns;
            _ = Columns.SetupGet(c => c.Count).Returns(columns.Length);
            _ = Table.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = Table.SetupGet(t => t.ListColumns).Returns(Columns.Object);
            _ = Table.SetupGet(t => t.DataBodyRange).Returns(Body.Object);
            _ = ListObjects.SetupGet(l => l.Count).Returns(1);
            _ = Worksheet.SetupGet(w => w.ListObjects).Returns(ListObjects.Object);
            _ = Sheets.SetupGet(s => s.Count).Returns(1);
            _ = Workbook.SetupGet(w => w.Sheets).Returns(Sheets.Object);
            _ = Application.SetupGet(a => a.ActiveWorkbook).Returns(Workbook.Object);
            _ = Guard.Setup(g => g.QueryTarget(It.IsAny<object?>()))
                .Returns(ProtectionGuardOutcome.NotProtected);
        }

        /// <summary>The colours and clears the adapter wrote, keyed by one-based cell.</summary>
        public Dictionary<(int Row, int Column), object?> Writes { get; } = [];

        /// <summary>An explicit payload, used to prove a non-matrix payload yields no rows.</summary>
        public object? BodyPayloadOverride { get; set; }

        /// <summary>
        /// The payload <c>Range.Value2</c> reports: the override when set, otherwise
        /// the matrix built from the added rows.
        /// </summary>
        public object? BodyPayload => BodyPayloadOverride ?? Matrix();

        /// <summary>The mocked worksheet.</summary>
        public Excel.Worksheet WorksheetObject => Worksheet.Object;

        /// <summary>The mocked data table.</summary>
        public Excel.ListObject TableObject => Table.Object;

        /// <summary>The mocked table body.</summary>
        public Excel.Range? BodyObject => _bodyPresent ? Body.Object : null;

        private Mock<Excel.Application> Application { get; } = new();

        private Mock<Excel.Workbook> Workbook { get; } = new();

        private Mock<Excel.Sheets> Sheets { get; } = new();

        private Mock<Excel.Worksheet> Worksheet { get; } = new();

        private Mock<Excel.ListObjects> ListObjects { get; } = new();

        private Mock<Excel.ListObject> Table { get; } = new();

        private Mock<Excel.ListColumns> Columns { get; } = new();

        private Mock<Excel.Range> Body { get; } = new();

        private Mock<IWorksheetProtectionGuard> Guard { get; } = new();

        private bool _bodyPresent = true;

        /// <summary>Adds a row, keyed by column name.</summary>
        /// <param name="cells">The values present in the row.</param>
        public void Row(params (string Column, object? Value)[] cells)
        {
            object?[] row = new object?[_columns.Length];
            foreach ((string column, object? value) in cells)
            {
                int index = Array.IndexOf(_columns, column);
                Assert.True(index >= 0, $"Column '{column}' is not in this fixture.");
                row[index] = value;
            }

            _rows.Add(row);
        }

        /// <summary>Runs the normaliser.</summary>
        /// <param name="applicationPresent">Whether the host supplied an application.</param>
        /// <returns>The outcome.</returns>
        public GanttTypeChangeNormaliseOutcome Normalise(bool applicationPresent = true) =>
            new TestableNormaliser(
                applicationPresent ? Application.Object : null,
                Guard.Object,
                this).Normalise();

        /// <summary>The value written to a one-based cell, or <see langword="null"/> when untouched.</summary>
        /// <param name="row">The one-based row.</param>
        /// <param name="column">The column name.</param>
        /// <returns>The written value.</returns>
        public object? WriteTo(int row, string column) =>
            Writes.TryGetValue((row, Array.IndexOf(_columns, column) + 1), out object? value)
                ? value
                : null;

        /// <summary>Substitutes the column-index seam.</summary>
        /// <param name="index">The one-based column index.</param>
        /// <returns>The mocked list column.</returns>
        public Excel.ListColumn ColumnAt(int index)
        {
            Mock<Excel.ListColumn> column = new();
            _ = column.SetupGet(c => c.Name).Returns(_columns[index - 1]);
            return column.Object;
        }

        /// <summary>Returns the mocked cell for a one-based address, recording writes.</summary>
        /// <param name="row">The one-based row.</param>
        /// <param name="column">The one-based column.</param>
        /// <returns>The mocked cell.</returns>
        public Excel.Range CellAt(int row, int column)
        {
            if (!_cells.TryGetValue((row, column), out Excel.Range? cell))
            {
                Mock<Excel.Range> mock = new();
                _ = mock
                    .SetupSet(r => r.Value2 = It.IsAny<object>())
                    .Callback<object?>(value => Writes[(row, column)] = value);
                cell = mock.Object;
                _cells[(row, column)] = cell;
            }

            return cell;
        }

        /// <summary>Removes the active workbook.</summary>
        public void ForgetWorkbook() => _ = Application
            .SetupGet(a => a.ActiveWorkbook)
            .Returns((Excel.Workbook)null!);

        /// <summary>Makes the worksheet report no tables.</summary>
        public void RemoveTables() => _ = ListObjects.SetupGet(l => l.Count).Returns(0);

        /// <summary>Makes the table report no body.</summary>
        public void RemoveBody()
        {
            _bodyPresent = false;
            _ = Table.SetupGet(t => t.DataBodyRange).Returns((Excel.Range)null!);
        }

        /// <summary>Makes the target-sheet query report protection.</summary>
        public void ProtectTarget() => _ = Guard
            .Setup(g => g.QueryTarget(It.IsAny<object?>()))
            .Returns(ProtectionGuardOutcome.SheetProtected);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1814:MultiDimArrays", Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
        private object?[,] Matrix()
        {
            object?[,] matrix = new object?[_rows.Count, _columns.Length];
            for (int row = 0; row < _rows.Count; row++)
            {
                for (int column = 0; column < _columns.Length; column++)
                {
                    matrix[row, column] = _rows[row][column];
                }
            }

            return matrix;
        }
    }
}
