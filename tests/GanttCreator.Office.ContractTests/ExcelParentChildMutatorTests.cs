using GanttCreator.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R5.3 parent/child mutation adapter.
/// </summary>
/// <remarks>
/// The tests drive the adapter through the same fake-COM-graph
/// seam discipline as the R2.8 inserter tests: every COM
/// parameterised property is an <c>internal virtual</c> seam
/// the testable subclass overrides, so the full guard-first
/// pipeline runs without a host. The Core planner runs for
/// real -- these tests pin the ADAPTER's Excel.Application of the
/// plan, not the planner's arithmetic (that is the Core
/// tests' job).
/// </remarks>
public class ExcelParentChildMutatorTests
{
	private sealed class Graph
	{
		private readonly Dictionary<int, Mock<Excel.Range>> _rowRanges = new Dictionary<int, Mock<Excel.Range>>();

		private readonly Dictionary<(int Row, int Column), Mock<Excel.Range>> _cellRanges = new Dictionary<(int, int), Mock<Excel.Range>>();

		private readonly Dictionary<Excel.Range, int> _rowIndexes = new Dictionary<Excel.Range, int>();

		private readonly Dictionary<Excel.Range, (int Row, int Column)> _cellIndexes = new Dictionary<Excel.Range, (int, int)>();

		private readonly Dictionary<int, Mock<Excel.Range>> _worksheetRows = new Dictionary<int, Mock<Excel.Range>>();

		private readonly Dictionary<Excel.Range, int> _worksheetRowIndexes = new Dictionary<Excel.Range, int>();

		public Mock<Excel.Application> Application { get; } = new Mock<Excel.Application>();

		public Mock<Excel.Workbook> Workbook { get; } = new Mock<Excel.Workbook>();

		public Mock<Excel.Sheets> Sheets { get; } = new Mock<Excel.Sheets>();

		public Mock<Excel.Worksheet> Worksheet { get; } = new Mock<Excel.Worksheet>();

		public Mock<Excel.ListObjects> ListObjects { get; } = new Mock<Excel.ListObjects>();

		public Mock<Excel.ListObject> Table { get; } = new Mock<Excel.ListObject>();

		public Mock<Excel.ListColumns> ListColumns { get; } = new Mock<Excel.ListColumns>();

		public Mock<Excel.ListColumn>[] Columns { get; } = (from _ in Enumerable.Range(0, GanttTableSchema.Default.Columns.Count)
			select new Mock<Excel.ListColumn>()).ToArray();

		public Mock<Excel.ListRows> ListRows { get; } = new Mock<Excel.ListRows>();

		public Mock<Excel.Range> TableRange { get; } = new Mock<Excel.Range>();

		public Mock<Excel.Range> DataBodyRange { get; } = new Mock<Excel.Range>();

		public Mock<Excel.Range> BodyRows { get; } = new Mock<Excel.Range>();

		public Mock<Excel.Range> ActiveCell { get; } = new Mock<Excel.Range>();

[System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1814:MultiDimArrays", Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
		public object?[,]? BodyMatrix { get; set; }

		public int RowCount { get; set; }

		public int ActiveRow { get; set; }

		public bool TableIsActive { get; set; } = true;

		public bool TablePresent { get; set; } = true;

		public bool ThrowOnWrite { get; set; }

		public bool ThrowOnDelete { get; set; }

		public bool ThrowOnHiddenWrite { get; set; }

		public List<(int Row, object? Payload)> WrittenRows { get; } = new List<(int Row, object? Payload)>();

		public List<(int Row, int Column, object? Value)> WrittenCells { get; } = new List<(int Row, int Column, object? Value)>();

		public List<int> DeletedRows { get; } = new List<int>();

		public HashSet<int> HiddenRows { get; } = new HashSet<int>();

		public List<(int Row, bool Hidden)> HiddenWrites { get; } = new List<(int Row, bool Hidden)>();

		/// <summary>
		/// Every worksheet row the hidden-flag rotation addressed, in
		/// address order (R5.4 D13).
		/// </summary>
		public List<int> WorksheetRowReads { get; } = new List<int>();

		public Graph()
		{
			Application.SetupGet((Excel.Application application) => application.ActiveWorkbook).Returns(Workbook.Object);
			Workbook.SetupGet((Excel.Workbook workbook) => workbook.Sheets).Returns(Sheets.Object);
			Sheets.SetupGet((Excel.Sheets sheets) => sheets.Count).Returns(1);
			Worksheet.SetupGet((Excel.Worksheet worksheet) => worksheet.ListObjects).Returns(ListObjects.Object);
			ListObjects.SetupGet((Excel.ListObjects listObjects) => listObjects.Count).Returns(() => TablePresent ? 1 : 0);
			Table.SetupGet((Excel.ListObject table) => table.Name).Returns("tblGanttData");
			Table.SetupGet((Excel.ListObject table) => table.ListColumns).Returns(ListColumns.Object);
			ListColumns.SetupGet((Excel.ListColumns columns) => columns.Count).Returns(GanttTableSchema.Default.Columns.Count);
			for (int num = 0; num < Columns.Length; num++)
			{
				Columns[num].SetupGet((Excel.ListColumn column) => column.Name).Returns(GanttTableSchema.Default.Columns[num].Name);
			}
			Table.SetupGet((Excel.ListObject table) => table.Range).Returns(TableRange.Object);
			TableRange.SetupGet((Excel.Range range) => range.Row).Returns(1);
			Table.SetupGet((Excel.ListObject table) => table.DataBodyRange).Returns(DataBodyRange.Object);
		DataBodyRange.SetupGet((Excel.Range range) => range.Row).Returns(2);
			ActiveCell.SetupGet((Excel.Range range) => range.Row).Returns(() => ActiveRow);
		}

		public Excel.Range RowRange(int row)
		{
			if (!_rowRanges.TryGetValue(row, out Mock<Excel.Range>? value))
			{
				value = new Mock<Excel.Range>();
				_rowRanges[row] = value;
				_rowIndexes[value.Object] = row;
			}
			return value.Object;
		}

		public Excel.Range CellRange(int row, int column)
		{
			if (!_cellRanges.TryGetValue((row, column), out Mock<Excel.Range>? value))
			{
				value = new Mock<Excel.Range>();
				_cellRanges[(row, column)] = value;
				_cellIndexes[value.Object] = (row, column);
			}
			return value.Object;
		}

		/// <summary>
		/// Returns the cached whole-row range for a one-based worksheet
		/// row. The range's <c>Hidden</c> reads back the Graph's hidden
		/// state, so a hidden-flag rotation round-trips through the
		/// same instance it addressed (R5.4 D13).
		/// </summary>
		/// <param name="row">The one-based worksheet row.</param>
		/// <returns>The cached row range.</returns>
		public Excel.Range WorksheetRow(int row)
		{
			if (!_worksheetRows.TryGetValue(row, out Mock<Excel.Range>? value))
			{
				value = new Mock<Excel.Range>();
				value.SetupGet((Excel.Range range) => range.Hidden).Returns(() => HiddenRows.Contains(row));
				_worksheetRows[row] = value;
				_worksheetRowIndexes[value.Object] = row;
			}
			return value.Object;
		}

		public void RecordWrite(Excel.Range range, object? payload)
		{
			if (ThrowOnWrite)
			{
				throw RefusedComException();
			}
			if (_rowIndexes.TryGetValue(range, out var value))
			{
				WrittenRows.Add((value, payload));
				return;
			}
			if (_cellIndexes.TryGetValue(range, out (int, int) value2))
			{
				WrittenCells.Add((value2.Item1, value2.Item2, payload));
				return;
			}
			throw new InvalidOperationException("A write reached an unmapped range.");
		}

		public static Dictionary<string, object?> ValuesOf(object? payload)
		{
			object[,] array = Assert.IsType<object[,]>(payload);
			Dictionary<string, object?> dictionary = new Dictionary<string, object?>();
			for (int i = 0; i < GanttTableSchema.Default.Columns.Count; i++)
			{
				dictionary[GanttTableSchema.Default.Columns[i].Name] = array[0, i];
			}
			return dictionary;
		}

		public TestableMutator Build(IWorksheetProtectionGuard guard, PromotePlan? craftedPromotePlan = null, DemotePlan? craftedDemotePlan = null, MovePlan? craftedMovePlan = null, Action<Excel.ListRows, int>? deleteListRowAt = null, IOutlineGroupPort? outlineGroups = null)
		{
			return new TestableMutator(Application.Object, guard, outlineGroups, (Excel.Sheets _, int _) => Worksheet.Object, delegate(Excel.ListObjects listObjects, int _)
			{
				Assert.Same(ListObjects.Object, listObjects);
				return Table.Object;
			}, (Excel.ListColumns _, int index) => Columns[index - 1].Object, delegate(Excel.ListObject table)
			{
				Assert.Same(Table.Object, table);
				return ListRows.Object;
			}, delegate(Excel.ListRows rows)
			{
				Assert.Same(ListRows.Object, rows);
				return RowCount;
			}, delegate(Excel.ListObject table)
			{
				Assert.Same(Table.Object, table);
				return TableRange.Object;
			}, delegate(Excel.ListObject table)
			{
				Assert.Same(Table.Object, table);
				return DataBodyRange.Object;
			}, delegate(Excel.Range range)
			{
				Assert.Same(DataBodyRange.Object, range);
				return BodyRows.Object;
			}, delegate(Excel.Range rows, int index)
			{
				Assert.Same(BodyRows.Object, rows);
				return RowRange(index);
			}, delegate(Excel.Range row, int columnIndex)
			{
				if (!_rowIndexes.TryGetValue(row, out var value))
				{
					throw new InvalidOperationException("A cell write reached an unmapped row range.");
				}
				return CellRange(value, columnIndex);
			}, (Excel.Range range) => (range == DataBodyRange.Object) ? BodyMatrix : null, RecordWrite, (Excel.Application _) => ActiveCell.Object, (Excel.ListObject table) => TableIsActive, (Excel.Range range) => range.Row, craftedPromotePlan, craftedDemotePlan, craftedMovePlan, deleteListRowAt ?? ((Action<Excel.ListRows, int>)delegate(Excel.ListRows rows, int index)
			{
				Assert.Same(ListRows.Object, rows);
				if (ThrowOnDelete)
				{
					throw RefusedComException();
				}
				DeletedRows.Add(index);
			}), readHiddenFlagAt: (Excel.Range range) => _worksheetRowIndexes.TryGetValue(range, out var hiddenRow) && HiddenRows.Contains(hiddenRow), worksheetRowAt: (Excel.Worksheet worksheet, int rowIndex) =>
			{
				WorksheetRowReads.Add(rowIndex);
				return WorksheetRow(rowIndex);
			}, setRowHiddenAt: (Excel.Range row, bool hidden) =>
			{
				if (ThrowOnHiddenWrite)
				{
					throw RefusedComException();
				}
				if (!_worksheetRowIndexes.TryGetValue(row, out var rowIndex))
				{
					throw new InvalidOperationException("A hidden-flag write reached an unmapped worksheet row.");
				}
				if (hidden)
				{
					HiddenRows.Add(rowIndex);
				}
				else
				{
					HiddenRows.Remove(rowIndex);
				}
				HiddenWrites.Add((rowIndex, hidden));
			});
		}
	}

	private sealed class TestableMutator(Excel.Application? application, IWorksheetProtectionGuard? guard, IOutlineGroupPort? outlineGroups, Func<Excel.Sheets, int, object> sheetAt, Func<Excel.ListObjects, int, Excel.ListObject> tableAt, Func<Excel.ListColumns, int, Excel.ListColumn> columnAt, Func<Excel.ListObject, Excel.ListRows> listRowsAt, Func<Excel.ListRows, int> listRowCountAt, Func<Excel.ListObject, Excel.Range> tableRangeAt, Func<Excel.ListObject, Excel.Range?> dataBodyRangeAt, Func<Excel.Range, Excel.Range> rangeRowsAt, Func<Excel.Range, int, Excel.Range> rangeAt, Func<Excel.Range, int, Excel.Range> cellAt, Func<Excel.Range, object?> rangeValue2At, Action<Excel.Range, object> setRangeValues, Func<Excel.Application, Excel.Range?> activeCellAt, Func<Excel.ListObject, bool> tableActiveAt, Func<Excel.Range, int> rangeRowAt, PromotePlan? craftedPromotePlan = null, DemotePlan? craftedDemotePlan = null, MovePlan? craftedMovePlan = null, Action<Excel.ListRows, int>? deleteListRowAt = null, Func<Excel.Range, bool>? readHiddenFlagAt = null, Func<Excel.Worksheet, int, Excel.Range?>? worksheetRowAt = null, Action<Excel.Range, bool>? setRowHiddenAt = null) : ExcelParentChildMutator(application, guard, outlineGroups)
	{
		private readonly Func<Excel.Sheets, int, object> _sheetAt = sheetAt;

		private readonly Func<Excel.ListObjects, int, Excel.ListObject> _tableAt = tableAt;

		private readonly Func<Excel.ListColumns, int, Excel.ListColumn> _columnAt = columnAt;

		private readonly Func<Excel.ListObject, Excel.ListRows> _listRowsAt = listRowsAt;

		private readonly Func<Excel.ListRows, int> _listRowCountAt = listRowCountAt;

		private readonly Func<Excel.ListObject, Excel.Range> _tableRangeAt = tableRangeAt;

		private readonly Func<Excel.ListObject, Excel.Range?> _dataBodyRangeAt = dataBodyRangeAt;

		private readonly Func<Excel.Range, Excel.Range> _rangeRowsAt = rangeRowsAt;

		private readonly Func<Excel.Range, int, Excel.Range> _rangeAt = rangeAt;

		private readonly Func<Excel.Range, int, Excel.Range> _cellAt = cellAt;

		private readonly Func<Excel.Range, object?> _rangeValue2At = rangeValue2At;

		private readonly Action<Excel.Range, object> _setRangeValues = setRangeValues;

		private readonly Func<Excel.Application, Excel.Range?> _activeCellAt = activeCellAt;

		private readonly Func<Excel.ListObject, bool> _tableActiveAt = tableActiveAt;

		private readonly Func<Excel.Range, int> _rangeRowAt = rangeRowAt;

		private readonly PromotePlan? _craftedPromotePlan = craftedPromotePlan;

		private readonly DemotePlan? _craftedDemotePlan = craftedDemotePlan;

		private readonly MovePlan? _craftedMovePlan = craftedMovePlan;

		private readonly Action<Excel.ListRows, int>? _deleteListRowAt = deleteListRowAt;

		private readonly Func<Excel.Range, bool>? _readHiddenFlagAt = readHiddenFlagAt;

		private readonly Func<Excel.Worksheet, int, Excel.Range?>? _worksheetRowAt = worksheetRowAt;

		private readonly Action<Excel.Range, bool>? _setRowHiddenAt = setRowHiddenAt;

		internal override object GetSheetAt(Excel.Sheets sheets, int index)
		{
			return _sheetAt(sheets, index);
		}

		internal override Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index)
		{
			return _tableAt(listObjects, index);
		}

		internal override Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index)
		{
			return _columnAt(columns, index);
		}

		internal override Excel.ListRows GetListRows(Excel.ListObject table)
		{
			return _listRowsAt(table);
		}

		internal override int GetListRowCount(Excel.ListRows rows)
		{
			return _listRowCountAt(rows);
		}

		internal override Excel.Range GetTableRange(Excel.ListObject table)
		{
			return _tableRangeAt(table);
		}

		internal override Excel.Range? GetDataBodyRange(Excel.ListObject table)
		{
			return _dataBodyRangeAt(table);
		}

		internal override Excel.Range GetRangeRows(Excel.Range range)
		{
			return _rangeRowsAt(range);
		}

		internal override Excel.Range GetRangeAt(Excel.Range rows, int index)
		{
			return _rangeAt(rows, index);
		}

		internal override Excel.Range GetCellAt(Excel.Range row, int columnIndex)
		{
			return _cellAt(row, columnIndex);
		}

		internal override object? GetRangeValue2(Excel.Range range)
		{
			return _rangeValue2At(range);
		}

		internal override void SetRangeValues(Excel.Range range, object payload)
		{
			_setRangeValues(range, payload);
		}

		internal override Excel.Range? GetActiveCell(Excel.Application application)
		{
			return _activeCellAt(application);
		}

		internal override bool GetTableActive(Excel.ListObject table)
		{
			return _tableActiveAt(table);
		}

		internal override int GetRangeRow(Excel.Range range)
		{
			return _rangeRowAt(range);
		}

		internal override PromotePlan CreatePromotePlan(IReadOnlyList<GanttEvent> rows, GanttRowId selectedId)
		{
			return _craftedPromotePlan ?? base.CreatePromotePlan(rows, selectedId);
		}

		internal override DemotePlan CreateDemotePlan(IReadOnlyList<GanttEvent> rows, GanttRowId selectedId)
		{
			return _craftedDemotePlan ?? base.CreateDemotePlan(rows, selectedId);
		}

		internal override MovePlan CreateMovePlan(IReadOnlyList<GanttEvent> rows, GanttRowId selectedId, bool moveUp)
		{
			return _craftedMovePlan ?? base.CreateMovePlan(rows, selectedId, moveUp);
		}

		internal override void DeleteListRow(Excel.ListRows rows, int index)
		{
			if (_deleteListRowAt == null)
			{
				base.DeleteListRow(rows, index);
			}
			else
			{
				_deleteListRowAt(rows, index);
			}
		}

		internal override bool ReadHiddenFlag(Excel.Range range)
		{
			return _readHiddenFlagAt is null ? base.ReadHiddenFlag(range) : _readHiddenFlagAt(range);
		}

		internal override Excel.Range? GetWorksheetRow(Excel.Worksheet worksheet, int rowIndex)
		{
			return _worksheetRowAt is null ? base.GetWorksheetRow(worksheet, rowIndex) : _worksheetRowAt(worksheet, rowIndex);
		}

		internal override void SetRowHidden(Excel.Range row, bool hidden)
		{
			if (_setRowHiddenAt is null)
			{
				base.SetRowHidden(row, hidden);
			}
			else
			{
				_setRowHiddenAt(row, hidden);
			}
		}
	}

	private const string ParentId = "G-0000000000000000000000000000000a";

	private const string ChildOneId = "G-0000000000000000000000000000000b";

	private const string ChildTwoId = "G-0000000000000000000000000000000c";

	private const string ChildThreeId = "G-0000000000000000000000000000000d";

	private const string GroupOneId = "G-0000000000000000000000000000000e";

	private const string GroupTwoId = "G-0000000000000000000000000000000f";

	private const string TargetId = "G-00000000000000000000000000000010";

	private static int TableIndexOf(string columnName)
	{
		for (int i = 0; i < GanttTableSchema.Default.Columns.Count; i++)
		{
			if (GanttTableSchema.Default.Columns[i].Name == columnName)
			{
				return i + 1;
			}
		}

		throw new InvalidOperationException("'" + columnName + "' is not a schema column.");
	}
	/// <summary>
	/// Builds one raw body row carrying the fields the hierarchy
	/// planners read, with the remaining schema columns blank.
	/// </summary>
	/// <param name="id">The row stable identifier.</param>
	/// <param name="type">The row type display name.</param>
	/// <param name="parentId">The row parent, or null.</param>
	/// <param name="siblingOrder">The row SiblingOrder, or null.</param>
	/// <param name="description">The row description.</param>
	/// <returns>The raw row, one value per table column.</returns>
	private static object?[] BodyRow(
		string id,
		string type,
		string? parentId = null,
		int? siblingOrder = null,
		string description = "Row")
	{
		IReadOnlyList<GanttTableColumn> columns = GanttTableSchema.Default.Columns;
		object?[] row = new object?[columns.Count];
		for (var index = 0; index < columns.Count; index++)
		{
			row[index] = columns[index].Name switch
			{
				"Id" => id,
				"Type" => type,
				"Description" => description,
				"ParentId" => parentId,
				"SiblingOrder" => siblingOrder,
				"Visible" => true,
				_ => null,
			};
		}

		return row;
	}

	/// <summary>
	/// Builds the rectangular Value2 payload of the table body.
	/// </summary>
	/// <param name="rows">The body rows.</param>
	/// <returns>The rectangular payload.</returns>
	[System.Diagnostics.CodeAnalysis.SuppressMessage(
		"Performance",
		"CA1814:MultiDimArrays",
		Justification = "Excel Value2 is a SAFEARRAY shaped as a multidimensional array; a jagged array cannot express it.")]
	private static object?[,] Matrix(params object?[][] rows)
	{
		object?[,] array = new object?[rows.Length, GanttTableSchema.Default.Columns.Count];
		for (int i = 0; i < rows.Length; i++)
		{
			for (int j = 0; j < rows[i].Length; j++)
			{
				array[i, j] = rows[i][j];
			}
		}
		return array;
	}

	private static System.Runtime.InteropServices.COMException RefusedComException()
	{
		var exception = (System.Runtime.InteropServices.COMException?)System.Runtime.InteropServices.Marshal.GetExceptionForHR(unchecked((int)0x800A03EC));
		return exception ?? throw new InvalidOperationException(
            "The runtime did not produce a COMException for HRESULT 0x800A03EC.");
	}

	private static string TypeOf(GanttEntityType type)
	{
		return EntityTypeCatalog.Format(type);
	}

[System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1814:MultiDimArrays", Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
	private static Graph SeededGraph(object?[,] matrix, int activeRow)
	{
		return new Graph
		{
			BodyMatrix = matrix,
			RowCount = matrix.GetLength(0),
			ActiveRow = activeRow
		};
	}

[System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1814:MultiDimArrays", Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
	private static object?[,] ThreeRowRun()
	{
		return Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 0, "First child"), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 1, "Second child"));
	}

	private static IWorksheetProtectionGuard NotProtected()
	{
		Mock<IWorksheetProtectionGuard> mock = new Mock<IWorksheetProtectionGuard>();
		mock.Setup((IWorksheetProtectionGuard g) => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
		mock.Setup((IWorksheetProtectionGuard g) => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
		return mock.Object;
	}

	[Fact]
	public void Promote_refuses_when_the_guard_reports_no_active_workbook()
	{
		Graph graph = new Graph();
		Mock<IWorksheetProtectionGuard> mock = new Mock<IWorksheetProtectionGuard>();
		mock.Setup((IWorksheetProtectionGuard g) => g.Query()).Returns(ProtectionGuardOutcome.NoActiveWorkbook);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(mock.Object).Promote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildMutationRefusalReason.NoActiveWorkbook, parentChildMutationOutcome.Refusal);
		Assert.Null(parentChildMutationOutcome.PlanRefusal);
		Assert.Null(parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Demote_refuses_when_the_active_sheet_is_protected()
	{
		Graph graph = new Graph();
		Mock<IWorksheetProtectionGuard> mock = new Mock<IWorksheetProtectionGuard>();
		mock.Setup((IWorksheetProtectionGuard g) => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(mock.Object).Demote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, parentChildMutationOutcome.Refusal);
		graph.Application.VerifyGet((Excel.Application application) => application.ActiveWorkbook, Times.Never);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Mutation_refuses_when_the_target_worksheet_is_protected()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 3);
		Mock<IWorksheetProtectionGuard> mock = new Mock<IWorksheetProtectionGuard>();
		mock.Setup((IWorksheetProtectionGuard g) => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
		mock.Setup((IWorksheetProtectionGuard g) => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.SheetProtected);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(mock.Object).Promote();
		ParentChildMutationOutcome parentChildMutationOutcome2 = graph.Build(mock.Object).Demote();
		Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, parentChildMutationOutcome.Refusal);
		Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, parentChildMutationOutcome2.Refusal);
		mock.Verify((IWorksheetProtectionGuard g) => g.QueryTarget(graph.Worksheet.Object), Times.Exactly(2));
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Mutation_refuses_when_the_table_is_missing()
	{
		Graph graph = new Graph
		{
			TablePresent = false
		};
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		ParentChildMutationOutcome parentChildMutationOutcome2 = graph.Build(NotProtected()).Demote();
		Assert.Equal(ParentChildMutationRefusalReason.TableMissing, parentChildMutationOutcome.Refusal);
		Assert.Equal(ParentChildMutationRefusalReason.TableMissing, parentChildMutationOutcome2.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Mutation_refuses_when_the_selection_is_the_header_row()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 1);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, parentChildMutationOutcome.Refusal);
		Assert.Null(parentChildMutationOutcome.SelectedRowNumber);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Mutation_refuses_when_the_selection_is_below_the_body()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 5);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Demote();
		Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Mutation_refuses_when_the_table_does_not_own_the_selection()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 2);
		graph.TableIsActive = false;
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Mutation_refuses_when_the_row_count_falls_short_of_the_selection()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId)), 4);
		graph.RowCount = 3;
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Mutation_refuses_when_the_selected_row_has_an_unparseable_id()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow("not-an-id", TypeOf(GanttEntityType.AsPlannedActivity), ParentId)), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.Equal(ParentChildMutationRefusalReason.SelectionUnknown, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Mutation_refuses_when_any_other_row_has_an_unparseable_id()
	{
		Graph graph = SeededGraph(Matrix(BodyRow("not-an-id", TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Demote();
		Assert.Equal(ParentChildMutationRefusalReason.TableInvalid, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

[System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1814:MultiDimArrays", Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
	private static object?[,] TwoChildRun()
	{
		return Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 7, "First child"), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 1, "Second child"));
	}

	[Fact]
	public void Promote_moves_the_selected_child_to_the_run_bottom()
	{
		Graph graph = SeededGraph(TwoChildRun(), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Null(parentChildMutationOutcome.Refusal);
		Assert.Null(parentChildMutationOutcome.PlanRefusal);
		Assert.Equal(2, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(2 * GanttTableSchema.Default.Columns.Count + 1, parentChildMutationOutcome.CellsWritten);
		Assert.Single(graph.WrittenCells);
		Assert.Equal(2, graph.WrittenRows.Count);
		Assert.Equal(2, graph.WrittenRows[0].Row);
		Dictionary<string, object?> dictionary = Graph.ValuesOf(graph.WrittenRows[0].Payload);
		Assert.Equal(ChildTwoId, dictionary["Id"]);
		Assert.Equal("Second child", dictionary["Description"]);
		Assert.Equal(ParentId, dictionary["ParentId"]);
		Assert.Equal(0, dictionary["SiblingOrder"]);
		Assert.Equal(3, graph.WrittenRows[1].Row);
		Dictionary<string, object?> dictionary2 = Graph.ValuesOf(graph.WrittenRows[1].Payload);
		Assert.Equal(ChildOneId, dictionary2["Id"]);
		Assert.Equal("First child", dictionary2["Description"]);
		Assert.Null(dictionary2["ParentId"]);
		Assert.Equal(1, dictionary2["SiblingOrder"]);
		Assert.Equal(1, graph.WrittenCells[0].Row);
		Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[0].Column);
		Assert.Equal(0, graph.WrittenCells[0].Value);
	}

	[Fact]
	public void Promote_of_a_middle_child_resequences_the_run_member_above_the_window()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 0, "First child"), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 1, "Second child"), BodyRow(ChildThreeId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 2, "Third child")), 4);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(3, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(2 * GanttTableSchema.Default.Columns.Count + 2, parentChildMutationOutcome.CellsWritten);
		Assert.Equal(2, graph.WrittenRows.Count);
		Assert.Equal(2, graph.WrittenCells.Count);
		Assert.Equal(3, graph.WrittenRows[0].Row);
		Dictionary<string, object?> dictionary = Graph.ValuesOf(graph.WrittenRows[0].Payload);
		Assert.Equal(ChildThreeId, dictionary["Id"]);
		Assert.Equal(1, dictionary["SiblingOrder"]);
		Assert.Equal(4, graph.WrittenRows[1].Row);
		Dictionary<string, object?> dictionary2 = Graph.ValuesOf(graph.WrittenRows[1].Payload);
		Assert.Equal(ChildTwoId, dictionary2["Id"]);
		Assert.Null(dictionary2["ParentId"]);
		Assert.Equal(1, dictionary2["SiblingOrder"]);
		Assert.Equal(2, graph.WrittenCells.Count);
		Assert.Equal(2, graph.WrittenCells[0].Row);
		Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[0].Column);
		Assert.Equal(0, graph.WrittenCells[0].Value);
		Assert.Equal(1, graph.WrittenCells[1].Row);
		Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[1].Column);
		Assert.Equal(0, graph.WrittenCells[1].Value);
	}

	[Fact]
	public void Promote_of_the_last_child_writes_the_row_and_the_run_above()
	{
		Graph graph = SeededGraph(TwoChildRun(), 4);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(3, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(GanttTableSchema.Default.Columns.Count + 2, parentChildMutationOutcome.CellsWritten);
		Assert.Single(graph.WrittenRows);
		Assert.Equal(2, graph.WrittenCells.Count);
		Assert.Equal(3, graph.WrittenRows[0].Row);
		Dictionary<string, object?> dictionary = Graph.ValuesOf(graph.WrittenRows[0].Payload);
		Assert.Equal(ChildTwoId, dictionary["Id"]);
		Assert.Null(dictionary["ParentId"]);
		Assert.Equal(1, dictionary["SiblingOrder"]);
		Assert.Equal(2, graph.WrittenCells.Count);
		Assert.Equal(2, graph.WrittenCells[0].Row);
		Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[0].Column);
		Assert.Equal(0, graph.WrittenCells[0].Value);
		Assert.Equal(1, graph.WrittenCells[1].Row);
		Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[1].Column);
		Assert.Equal(0, graph.WrittenCells[1].Value);
	}

	[Fact]
	public void Demote_assigns_the_nearest_eligible_parent_and_moves_beneath_its_run()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Demote();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(2, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(GanttTableSchema.Default.Columns.Count + 1, parentChildMutationOutcome.CellsWritten);
		Assert.Single(graph.WrittenRows);
		Assert.Single(graph.WrittenCells);
		Assert.Equal(2, graph.WrittenRows[0].Row);
		Dictionary<string, object?> dictionary = Graph.ValuesOf(graph.WrittenRows[0].Payload);
		Assert.Equal(TargetId, dictionary["Id"]);
		Assert.Equal(ParentId, dictionary["ParentId"]);
		Assert.Equal(0, dictionary["SiblingOrder"]);
		Assert.Equal(1, graph.WrittenCells[0].Row);
		Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[0].Column);
		Assert.Equal(0, graph.WrittenCells[0].Value);
	}

	[Fact]
	public void Demote_across_an_unrelated_row_moves_beneath_the_target_run()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 0, "First child"), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedMilestone), null, 1, "Milestone"), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, 2, "Target")), 5);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Demote();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(4, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(2 * GanttTableSchema.Default.Columns.Count + 1, parentChildMutationOutcome.CellsWritten);
		Assert.Equal(2, graph.WrittenRows.Count);
		Assert.Equal(3, graph.WrittenRows[0].Row);
		Dictionary<string, object?> dictionary = Graph.ValuesOf(graph.WrittenRows[0].Payload);
		Assert.Equal(TargetId, dictionary["Id"]);
		Assert.Equal("Target", dictionary["Description"]);
		Assert.Equal(ParentId, dictionary["ParentId"]);
		Assert.Equal(1, dictionary["SiblingOrder"]);
		Assert.Equal(4, graph.WrittenRows[1].Row);
		Dictionary<string, object?> dictionary2 = Graph.ValuesOf(graph.WrittenRows[1].Payload);
		Assert.Equal(ChildTwoId, dictionary2["Id"]);
		Assert.Equal("Milestone", dictionary2["Description"]);
		Assert.Null(dictionary2["ParentId"]);
		Assert.Single(graph.WrittenCells);
		Assert.Equal(1, graph.WrittenCells[0].Row);
		Assert.Equal(TableIndexOf("SiblingOrder"), graph.WrittenCells[0].Column);
		Assert.Equal(0, graph.WrittenCells[0].Value);
	}

	[Fact]
	public void Demote_of_a_parent_rehomes_its_children_to_the_new_parent()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, 1, "Target"), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), TargetId, 0, "Rehomed child")), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Demote();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(2, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Single(graph.WrittenRows);
		Assert.Equal(2, graph.WrittenRows[0].Row);
		Dictionary<string, object?> dictionary = Graph.ValuesOf(graph.WrittenRows[0].Payload);
		Assert.Equal(TargetId, dictionary["Id"]);
		Assert.Equal(ParentId, dictionary["ParentId"]);
		Assert.Equal(0, dictionary["SiblingOrder"]);
		Assert.Contains(graph.WrittenCells, ((int Row, int Column, object? Value) cell) => cell.Row == 3 && cell.Column == TableIndexOf("ParentId") && object.Equals(cell.Value, ParentId));
	}

	[Fact]
	public void Demoting_adjacent_to_a_full_group_is_refused_as_parent_unavailable()
	{
		List<object?[]> list = new List<object?[]>
		{
			BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, 0, "Group one"),
			BodyRow(GroupTwoId, TypeOf(GanttEntityType.AsPlannedActivity), null, 1, "Group two")
		};
		for (int i = 0; i < 7; i++)
		{
			list.Add(BodyRow($"G-{1000 + i:x32}", TypeOf(GanttEntityType.AsPlannedActivity), GroupTwoId, i, $"Group two child {i}"));
		}
		list.Add(BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, 2, "Target"));
		Graph graph = SeededGraph(Matrix(list.ToArray()), 11);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Demote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Null(parentChildMutationOutcome.Refusal);
		Assert.Equal(ParentChildRefusalReason.ParentUnavailable, parentChildMutationOutcome.PlanRefusal);
		Assert.Equal(10, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Promote_of_a_top_level_row_is_refused_as_not_a_child()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 2);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Null(parentChildMutationOutcome.Refusal);
		Assert.Equal(ParentChildRefusalReason.NotAChild, parentChildMutationOutcome.PlanRefusal);
		Assert.Equal(1, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Promote_of_a_critical_interval_is_refused()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(ChildOneId, TypeOf(GanttEntityType.CriticalInterval), ParentId)), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildRefusalReason.PromotionForbiddenForType, parentChildMutationOutcome.PlanRefusal);
		Assert.Equal(2, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Promote_of_a_row_that_owns_children_is_refused()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedActivity), ChildOneId)), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildRefusalReason.HasChildren, parentChildMutationOutcome.PlanRefusal);
		Assert.Equal(2, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Demote_with_no_eligible_parent_above_is_refused()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))), 2);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Demote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildRefusalReason.NoEligibleParentAbove, parentChildMutationOutcome.PlanRefusal);
		Assert.Equal(1, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Demote_of_a_row_that_is_already_a_child_is_refused()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId)), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Demote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildRefusalReason.AlreadyAChild, parentChildMutationOutcome.PlanRefusal);
		Assert.Equal(2, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void A_promote_plan_that_moves_upward_is_refused_rather_than_rotated()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 3);
		PromotePlan craftedPromotePlan = new PromotePlan(GanttRowId.Parse(ChildOneId), GanttRowId.Parse(ParentId), 1, Array.Empty<SiblingOrderWrite>(), null);
		IParentChildMutator parentChildMutator = graph.Build(NotProtected(), craftedPromotePlan);
		ParentChildMutationOutcome parentChildMutationOutcome = parentChildMutator.Promote();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildMutationRefusalReason.TableInvalid, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void A_host_write_refusal_is_reported_as_failed()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 3);
		graph.ThrowOnWrite = true;
		Graph graph2 = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))), 3);
		graph2.ThrowOnWrite = true;
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Promote();
		ParentChildMutationOutcome parentChildMutationOutcome2 = graph2.Build(NotProtected()).Demote();
		Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, parentChildMutationOutcome.Refusal);
		Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, parentChildMutationOutcome2.Refusal);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Equal(0, parentChildMutationOutcome2.CellsWritten);
	}

	private static GanttRowId ParsedId(string text)
	{
		Assert.True(GanttRowId.TryParse(text, out GanttRowId? id) && id is not null);
		return id;
	}

	[Fact]
	public void Moving_a_group_up_moves_the_moved_row_and_resequences_outside_cells()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "A"), BodyRow(GroupTwoId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "B"), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "C")), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).MoveUp();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(1, parentChildMutationOutcome.MovedToRowNumber);
		int count = GanttTableSchema.Default.Columns.Count;
		Assert.Empty(graph.DeletedRows);
		Assert.Equal(2, graph.WrittenRows.Count);
		Assert.Equal(1, graph.WrittenRows[0].Row);
		Assert.Equal(2, graph.WrittenRows[1].Row);
		Assert.Equal(2 * count + 1, parentChildMutationOutcome.CellsWritten);
		Assert.Equal(GroupTwoId, Graph.ValuesOf(graph.WrittenRows[0].Payload)["Id"]);
		Assert.Equal(GroupOneId, Graph.ValuesOf(graph.WrittenRows[1].Payload)["Id"]);
		int item = TableIndexOf("SiblingOrder");
		Assert.Equal((3, item, 2), Assert.Single(graph.WrittenCells));
	}

	[Fact]
	public void Moving_a_group_down_moves_the_neighbour_row_above_the_group()
	{
		Graph graph = SeededGraph(ThreeRowTopLevelAndChild(), 2);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).MoveDown();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(2, parentChildMutationOutcome.MovedToRowNumber);
		int count = GanttTableSchema.Default.Columns.Count;
		Assert.Empty(graph.DeletedRows);
		Assert.Equal(3, graph.WrittenRows.Count);
		Assert.Equal(3 * count, parentChildMutationOutcome.CellsWritten);
		Assert.Equal(TargetId, Graph.ValuesOf(graph.WrittenRows[0].Payload)["Id"]);
		Assert.Equal(GroupOneId, Graph.ValuesOf(graph.WrittenRows[1].Payload)["Id"]);
		Assert.Equal(ChildOneId, Graph.ValuesOf(graph.WrittenRows[2].Payload)["Id"]);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Moving_a_child_up_swaps_the_sibling_rows_and_resequences_the_run()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 4);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).MoveUp();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(2, parentChildMutationOutcome.MovedToRowNumber);
		int count = GanttTableSchema.Default.Columns.Count;
		Assert.Empty(graph.DeletedRows);
		Assert.Equal(2, graph.WrittenRows.Count);
		Assert.Equal(2 * count, parentChildMutationOutcome.CellsWritten);
		Dictionary<string, object?> dictionary = Graph.ValuesOf(graph.WrittenRows[0].Payload);
		Assert.Equal(ChildTwoId, dictionary["Id"]);
		Assert.Equal(0, dictionary["SiblingOrder"]);
		Dictionary<string, object?> dictionary2 = Graph.ValuesOf(graph.WrittenRows[1].Payload);
		Assert.Equal(ChildOneId, dictionary2["Id"]);
		Assert.Equal(1, dictionary2["SiblingOrder"]);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Moving_up_at_the_top_of_the_table_writes_nothing()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))), 2);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).MoveUp();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Null(parentChildMutationOutcome.MovedToRowNumber);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Moving_down_at_the_bottom_of_the_table_writes_nothing()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))), 3);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).MoveDown();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Null(parentChildMutationOutcome.MovedToRowNumber);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void A_crafted_move_plan_with_overlapping_windows_is_refused()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 2);
		MovePlan craftedMovePlan = new MovePlan(ParsedId(ParentId), 0, 2, 1, 2, Array.Empty<SiblingOrderWrite>(), null);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, craftedMovePlan).MoveUp();
		Assert.Equal(ParentChildMutationRefusalReason.TableInvalid, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void A_crafted_move_plan_with_an_out_of_range_window_is_refused()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 2);
		MovePlan craftedMovePlan = new MovePlan(ParsedId(ParentId), 2, 5, 0, 1, Array.Empty<SiblingOrderWrite>(), null);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, craftedMovePlan).MoveDown();
		Assert.Equal(ParentChildMutationRefusalReason.TableInvalid, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void A_crafted_move_plan_refusal_names_the_selected_row()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 3);
		MovePlan craftedMovePlan = MovePlan.Refused(ParentChildRefusalReason.SelectionUnknown);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, craftedMovePlan).MoveUp();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildRefusalReason.SelectionUnknown, parentChildMutationOutcome.PlanRefusal);
		Assert.Equal(2, parentChildMutationOutcome.SelectedRowNumber);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Moving_a_group_down_moves_the_moved_row_below_a_longer_neighbour()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "A"), BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "B"), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "B's child")), 2);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).MoveDown();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(3, parentChildMutationOutcome.MovedToRowNumber);
		int count = GanttTableSchema.Default.Columns.Count;
		Assert.Empty(graph.DeletedRows);
		Assert.Equal(3, graph.WrittenRows.Count);
		Assert.Equal(3 * count, parentChildMutationOutcome.CellsWritten);
		Assert.Equal(GroupOneId, Graph.ValuesOf(graph.WrittenRows[0].Payload)["Id"]);
		Assert.Equal(ChildOneId, Graph.ValuesOf(graph.WrittenRows[1].Payload)["Id"]);
		Assert.Equal(GroupOneId, Graph.ValuesOf(graph.WrittenRows[1].Payload)["ParentId"]);
		Assert.Equal(TargetId, Graph.ValuesOf(graph.WrittenRows[2].Payload)["Id"]);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Moving_between_gap_rows_rotates_the_content()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "X"), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "G"), BodyRow(GroupTwoId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "Y")), 4);
		MovePlan craftedMovePlan = new MovePlan(ParsedId(GroupTwoId), 2, 1, 0, 1, Array.Empty<SiblingOrderWrite>(), null);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, craftedMovePlan).MoveUp();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(1, parentChildMutationOutcome.MovedToRowNumber);
		int count = GanttTableSchema.Default.Columns.Count;
		Assert.Equal(3 * count, parentChildMutationOutcome.CellsWritten);
		Assert.Equal(3, graph.WrittenRows.Count);
		Assert.Equal(1, graph.WrittenRows[0].Row);
		Assert.Equal(2, graph.WrittenRows[1].Row);
		Assert.Equal(3, graph.WrittenRows[2].Row);
		Assert.Equal(GroupTwoId, Graph.ValuesOf(graph.WrittenRows[0].Payload)["Id"]);
		Assert.Equal(TargetId, Graph.ValuesOf(graph.WrittenRows[1].Payload)["Id"]);
		Assert.Equal(GroupOneId, Graph.ValuesOf(graph.WrittenRows[2].Payload)["Id"]);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void A_move_whose_write_is_refused_reports_failure()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "A"), BodyRow(GroupTwoId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "B"), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "C")), 3);
		graph.ThrowOnWrite = true;
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).MoveUp();
		Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, parentChildMutationOutcome.Refusal);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
		Assert.Empty(graph.DeletedRows);
	}

	[Fact]
	public void Move_refuses_when_the_target_sheet_is_protected()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 3);
		Mock<IWorksheetProtectionGuard> mock = new Mock<IWorksheetProtectionGuard>();
		mock.Setup((IWorksheetProtectionGuard g) => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
		mock.Setup((IWorksheetProtectionGuard g) => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.SheetProtected);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(mock.Object).MoveUp();
		ParentChildMutationOutcome parentChildMutationOutcome2 = graph.Build(mock.Object).MoveDown();
		Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, parentChildMutationOutcome.Refusal);
		Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, parentChildMutationOutcome2.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Remove_refuses_when_the_guard_reports_no_active_workbook()
	{
		Graph graph = new Graph();
		Mock<IWorksheetProtectionGuard> mock = new Mock<IWorksheetProtectionGuard>();
		mock.Setup((IWorksheetProtectionGuard g) => g.Query()).Returns(ProtectionGuardOutcome.NoActiveWorkbook);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(mock.Object).Remove();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildMutationRefusalReason.NoActiveWorkbook, parentChildMutationOutcome.Refusal);
		Assert.Null(parentChildMutationOutcome.PlanRefusal);
		Assert.Null(parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
		Assert.Empty(graph.DeletedRows);
	}

	[Fact]
	public void Remove_refuses_when_the_target_sheet_is_protected()
	{
		Graph graph = SeededGraph(ThreeRowTopLevelAndChild(), 2);
		Mock<IWorksheetProtectionGuard> mock = new Mock<IWorksheetProtectionGuard>();
		mock.Setup((IWorksheetProtectionGuard g) => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
		mock.Setup((IWorksheetProtectionGuard g) => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.SheetProtected);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(mock.Object).Remove();
		Assert.Equal(ParentChildMutationRefusalReason.TargetProtected, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
		Assert.Empty(graph.DeletedRows);
	}

	[Fact]
	public void Remove_refuses_when_the_selection_is_the_header_row()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 1);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Remove();
		Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
		Assert.Empty(graph.DeletedRows);
	}

	[Fact]
	public void Remove_of_the_last_row_deletes_it_and_resequences_the_surviving_run()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, 0, "A"), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "A child"), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, 1, "B")), 4);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Remove();
		Assert.True(parentChildMutationOutcome.Succeeded);
		int num = 1;
		List<int> list = new List<int>(num);
		System.Runtime.InteropServices.CollectionsMarshal.SetCount(list, num);
		System.Runtime.InteropServices.CollectionsMarshal.AsSpan(list)[0] = 3;
		Assert.Equal(list, graph.DeletedRows);
		Assert.Empty(graph.WrittenRows);
		int item = TableIndexOf("SiblingOrder");
		Assert.Equal((1, item, 0), Assert.Single(graph.WrittenCells));
		Assert.Equal(1, parentChildMutationOutcome.CellsWritten);
	}

	[Fact]
	public void Remove_of_a_lone_row_deletes_it_and_writes_nothing()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, 0)), 2);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Remove();
		Assert.True(parentChildMutationOutcome.Succeeded);
		int num = 1;
		List<int> list = new List<int>(num);
		System.Runtime.InteropServices.CollectionsMarshal.SetCount(list, num);
		System.Runtime.InteropServices.CollectionsMarshal.AsSpan(list)[0] = 1;
		Assert.Equal(list, graph.DeletedRows);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
	}

	[Fact]
	public void Remove_promotes_the_deleted_row_s_children_to_its_parent_and_resequences_the_run()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, 0, "A"), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "A child one"), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 1, "A child two"), BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 2, "D"), BodyRow(ChildThreeId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 0, "G1"), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), ParentId, 1, "G2")), 5);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Remove();
		Assert.True(parentChildMutationOutcome.Succeeded);
		int num = 1;
		List<int> list = new List<int>(num);
		System.Runtime.InteropServices.CollectionsMarshal.SetCount(list, num);
		System.Runtime.InteropServices.CollectionsMarshal.AsSpan(list)[0] = 4;
		Assert.Equal(list, graph.DeletedRows);
		Assert.Empty(graph.WrittenRows);
		int item = TableIndexOf("ParentId");
		int item2 = TableIndexOf("SiblingOrder");
		string item3 = GroupOneId;
		Assert.Equal<List<(int Row, int Column, object? Value)>>(new List<(int Row, int Column, object? Value)>
		{
			(2, item2, 0),
			(3, item2, 1),
			(5, item, item3),
			(5, item2, 2),
			(6, item, item3),
			(6, item2, 3)
		}, graph.WrittenCells);
		Assert.Equal(6, parentChildMutationOutcome.CellsWritten);
	}

	[Fact]
	public void Remove_of_a_top_level_row_clears_its_children_s_parent_cells()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, 0, "A"), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "A child one"), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 1, "A child two"), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, 1, "B")), 2);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Remove();
		Assert.True(parentChildMutationOutcome.Succeeded);
		int num = 1;
		List<int> list = new List<int>(num);
		System.Runtime.InteropServices.CollectionsMarshal.SetCount(list, num);
		System.Runtime.InteropServices.CollectionsMarshal.AsSpan(list)[0] = 1;
		Assert.Equal(list, graph.DeletedRows);
		Assert.Empty(graph.WrittenRows);
		int item = TableIndexOf("ParentId");
		int item2 = TableIndexOf("SiblingOrder");
		Assert.Equal<List<(int Row, int Column, object? Value)>>(new List<(int Row, int Column, object? Value)>
		{
			(2, item, null),
			(2, item2, 0),
			(3, item, null),
			(3, item2, 1),
			(4, item2, 2)
		}, graph.WrittenCells);
		Assert.Equal(5, parentChildMutationOutcome.CellsWritten);
	}

	[Fact]
	public void Remove_translates_a_refused_cell_write_to_write_failed_without_deleting()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, 0, "A"), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "First child"), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 1, "Second child")), 3);
		graph.ThrowOnWrite = true;
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Remove();
		Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.DeletedRows);
	}

	[Fact]
	public void Remove_translates_a_refused_row_delete_to_write_failed()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, 0, "A"), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "First child"), BodyRow(ChildTwoId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 1, "Second child")), 3);
		graph.ThrowOnDelete = true;
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Remove();
		Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, parentChildMutationOutcome.Refusal);
		Assert.Empty(graph.DeletedRows);
		int item = TableIndexOf("SiblingOrder");
		Assert.Equal((3, item, 0), Assert.Single(graph.WrittenCells));
	}

	[Fact]
	public void DeleteListRow_deletes_the_row_s_entire_worksheet_row()
	{
		Mock<Excel.Range> mock = new Mock<Excel.Range>();
		Mock<Excel.Range> mock2 = new Mock<Excel.Range>();
		mock2.SetupGet((Excel.Range r) => r.EntireRow).Returns(mock.Object);
		Mock<Excel.ListRow> mock3 = new Mock<Excel.ListRow>();
		mock3.SetupGet((Excel.ListRow r) => r.Range).Returns(mock2.Object);
		Mock<Excel.ListRows> mock4 = new Mock<Excel.ListRows>();
		mock4.Setup((Excel.ListRows r) => r[2]).Returns(mock3.Object);
		ExcelParentChildMutator excelParentChildMutator = new ExcelParentChildMutator(null);
		excelParentChildMutator.DeleteListRow(mock4.Object, 2);
		mock.Verify((Excel.Range r) => r.Delete(Type.Missing));
		mock3.Verify((Excel.ListRow r) => r.Delete(), Times.Never());
	}

[System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1814:MultiDimArrays", Justification = "Excel's Value2 is a SAFEARRAY and is shaped as a multidimensional array; a jagged array cannot express it.")]
	private static object?[,] ThreeRowTopLevelAndChild()
	{
		return Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "A"), BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "A child"), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "B"));
	}

	[Fact]
	public void Remove_updates_no_cells_when_the_selection_is_outside_the_table()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 5);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).Remove();
		Assert.False(parentChildMutationOutcome.Succeeded);
		Assert.Equal(ParentChildMutationRefusalReason.SelectionOutsideTable, parentChildMutationOutcome.Refusal);
		Assert.Null(parentChildMutationOutcome.PlanRefusal);
		Assert.Null(parentChildMutationOutcome.SelectedRowNumber);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Assert.Empty(graph.WrittenRows);
		Assert.Empty(graph.WrittenCells);
	}

	[Fact]
	public void Promote_regroups_the_outline_from_the_post_mutation_table()
	{
		Graph graph = SeededGraph(TwoChildRun(), 3);
		Mock<IOutlineGroupPort> mock = new Mock<IOutlineGroupPort>();
		mock.Setup((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>())).Returns(OutlineGroupOutcome.Ok(1, 1));
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, null, null, mock.Object).Promote();
		Assert.True(parentChildMutationOutcome.Succeeded);
		mock.Verify((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>()), Times.Once);
	}

	[Fact]
	public void Demote_regroups_the_outline_from_the_post_mutation_table()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(ParentId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))), 3);
		Mock<IOutlineGroupPort> mock = new Mock<IOutlineGroupPort>();
		mock.Setup((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>())).Returns(OutlineGroupOutcome.Ok(1, 0));
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, null, null, mock.Object).Demote();
		Assert.True(parentChildMutationOutcome.Succeeded);
		mock.Verify((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>()), Times.Once);
	}

	[Fact]
	public void Move_regroups_the_outline_from_the_post_mutation_table()
	{
		Graph graph = SeededGraph(ThreeRowTopLevelAndChild(), 4);
		Mock<IOutlineGroupPort> mock = new Mock<IOutlineGroupPort>();
		mock.Setup((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>())).Returns(OutlineGroupOutcome.Ok(1, 0));
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, null, null, mock.Object).MoveUp();
		Assert.True(parentChildMutationOutcome.Succeeded);
		mock.Verify((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>()), Times.Once);
	}

	[Fact]
	public void Moving_a_collapsed_group_rotates_its_row_hidden_flags_with_the_content()
	{
		#region
		// An asymmetric collapse: GroupOne (a collapsed parent) with its
		// hidden child, followed by the visible top-level Target. Worksheet
		// rows are 2..4, so the body starts at worksheet row 2 and the
		// hidden child is worksheet row 3.
		#endregion
		Graph graph = SeededGraph(
			Matrix(
				BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "A"),
				BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "A child"),
				BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "B")),
			activeRow: 2);

		// GroupOne is collapsed, so its child row (worksheet row 3) is hidden.
		graph.HiddenRows.Add(3);

		// Move GroupOne's group down past Target: the content becomes
		// Target, GroupOne, ChildOne over worksheet rows 2..4.
		ParentChildMutationOutcome outcome = graph.Build(NotProtected()).MoveDown();

		Assert.True(outcome.Succeeded);

		// The whole union span was addressed.
		Assert.Contains(2, graph.WorksheetRowReads);
		Assert.Contains(3, graph.WorksheetRowReads);
		Assert.Contains(4, graph.WorksheetRowReads);

		// The hidden flag travelled with its row (R5.4 D13): Target is
		// visible on row 2, GroupOne is visible on row 3, and ChildOne
		// is hidden on row 4. Without the rotation GroupOne would land
		// on the still-hidden row 3, and Refresh would refuse with
		// DegenerateRowHeight.
		Assert.DoesNotContain(2, graph.HiddenRows);
		Assert.DoesNotContain(3, graph.HiddenRows);
		Assert.Contains(4, graph.HiddenRows);

		// Two flags changed: row 3 un-hid for GroupOne and row 4
		// hid for ChildOne. Without the rotation neither would be
		// written, because every flag would already match its
		// stale position.
		Assert.Equal((3, false), graph.HiddenWrites[0]);
		Assert.Equal((4, true), graph.HiddenWrites[1]);
	}

	[Fact]

	public void A_hidden_flag_write_refusal_is_reported_as_failed()

	{

		// A collapsed group whose move rotates hidden flags: the

		// rotation writes flags, so a refused hidden-flag write is

		// the D13 WriteFailed path, distinct from the content-write

		// refusal the ThrowOnWrite tests already pin.

		Graph graph = SeededGraph(

			Matrix(

				BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "A"),

				BodyRow(ChildOneId, TypeOf(GanttEntityType.AsPlannedActivity), GroupOneId, 0, "A child"),

				BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity), null, null, "B")),

			activeRow: 2);

		graph.HiddenRows.Add(3);

		graph.ThrowOnHiddenWrite = true;

		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected()).MoveDown();

		Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, parentChildMutationOutcome.Refusal);

		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);

		// The hidden-flag write is the first host write, so a

		// refusal there leaves the content unwritten: the failure

		// is the D13 path, not the content-write path.

		Assert.Empty(graph.WrittenRows);

		Assert.Empty(graph.WrittenCells);

	}



	[Fact]

	public void WriteHiddenFlags_refuses_when_the_first_worksheet_row_is_unresolved()

	{

		// firstWorksheetRow <= 0 means the body row could not be

		// resolved, so no row is addressed and the rotation is

		// refused rather than silently skipped: the content landed

		// but its presentation state did not.

		Mock<Excel.Worksheet> worksheet = new Mock<Excel.Worksheet>();

		ExcelParentChildMutator excelParentChildMutator = new ExcelParentChildMutator(null);

		ParentChildMutationOutcome outcome = excelParentChildMutator.WriteHiddenFlags(worksheet.Object, 0, new List<bool> { false }, 0);

		Assert.Equal(ParentChildMutationRefusalReason.WriteFailed, outcome.Refusal);

	}
	[Fact]
	public void Remove_regroups_the_outline_from_the_post_mutation_table()
	{
		Graph graph = SeededGraph(ThreeRowRun(), 2);
		Mock<IOutlineGroupPort> mock = new Mock<IOutlineGroupPort>();
		mock.Setup((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>())).Returns(OutlineGroupOutcome.Ok(0, 1));
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, null, null, mock.Object).Remove();
		Assert.True(parentChildMutationOutcome.Succeeded);
		mock.Verify((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>()), Times.Once);
	}

	[Fact]
	public void Mutation_skips_the_regroup_on_a_boundary_no_op_and_survives_an_outline_refusal()
	{
		Graph graph = SeededGraph(Matrix(BodyRow(GroupOneId, TypeOf(GanttEntityType.AsPlannedActivity)), BodyRow(TargetId, TypeOf(GanttEntityType.AsPlannedActivity))), 2);
		Mock<IOutlineGroupPort> mock = new Mock<IOutlineGroupPort>(MockBehavior.Strict);
		ParentChildMutationOutcome parentChildMutationOutcome = graph.Build(NotProtected(), null, null, null, null, mock.Object).MoveUp();
		Assert.True(parentChildMutationOutcome.Succeeded);
		Assert.Equal(0, parentChildMutationOutcome.CellsWritten);
		Graph graph2 = SeededGraph(TwoChildRun(), 3);
		Mock<IOutlineGroupPort> mock2 = new Mock<IOutlineGroupPort>();
		mock2.Setup((IOutlineGroupPort o) => o.Apply(It.IsAny<IReadOnlyList<GanttEvent>>())).Returns(OutlineGroupOutcome.Refused(OutlineGroupRefusalReason.TargetProtected));
		ParentChildMutationOutcome parentChildMutationOutcome2 = graph2.Build(NotProtected(), null, null, null, null, mock2.Object).Promote();
		Assert.True(parentChildMutationOutcome2.Succeeded);
	}
}
