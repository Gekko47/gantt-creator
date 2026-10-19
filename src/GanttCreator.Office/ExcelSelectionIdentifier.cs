using GanttCreator.Core;
// Namespace alias rather than a wholesale using: Microsoft.Office.Interop.Excel
// declares its own Action type, which would collide with System.Action under a
// plain `using Microsoft.Office.Interop.Excel;`.
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="ISelectionIdentifier"/> over the Excel application object
/// supplied by the host at add-in load. Resolves the active cell against
/// <c>tblGanttData</c> on demand (reusing the row inserter/mutator active-cell
/// seam shape) and pushes Excel's native
/// <c>Application.SheetSelectionChange</c> event.
/// </summary>
/// <remarks>
/// <para>
/// COM ownership: the <c>Application</c> object and every object reachable
/// from it are Excel-owned shared roots. This adapter takes no ownership of
/// them, never calls <c>FinalReleaseComObject</c>, and force-releases nothing.
/// The one resource it owns is its event connection, detached exactly once.
/// </para>
/// <para>
/// Every member degrades instead of throwing: the adapter is reached from
/// Ribbon state callbacks and from Excel application events, where a
/// propagating exception surfaces as a host error dialog.
/// </para>
/// </remarks>
public class ExcelSelectionIdentifier : ISelectionIdentifier
{
    private readonly Excel.Application? _application;
    private readonly Excel.AppEvents_Event? _events;

    /// <summary>
    /// Creates a selection identifier over the Excel application object.
    /// </summary>
    /// <param name="application">
    /// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
    /// <see langword="null"/> when the host supplied none (unit tests, non-Excel
    /// host). A foreign object fails the interface casts and degrades to
    /// <see cref="SelectionFacts.None"/> with an inert subscription.
    /// </param>
    public ExcelSelectionIdentifier(object? application)
    {
        _application = application as Excel.Application;

        // Cast through _application rather than from `application` directly:
        // `Excel.Application` inherits from `Excel.AppEvents_Event`, so if the
        // application object was successfully cast, the event interface will also
        // be available. Casting directly from object? to AppEvents_Event can fail
        // on the Excel-DNA RCW even when the Application cast succeeds, which
        // would leave SheetSelectionChange unsubscribed and ribbon dropdowns stale.
        _events = _application;
    }

    /// <inheritdoc />
    public SelectionFacts Identify()
    {
        Excel.Application? application = _application;
        if (application is null)
        {
            return SelectionFacts.None;
        }

        // CA1031: runs on the Ribbon state path; a failed probe reports
        // "no selection" so the state service keeps its previous snapshot.
#pragma warning disable CA1031
        try
        {
            Excel.Workbook? workbook = application.ActiveWorkbook;
            if (workbook is null)
            {
                return SelectionFacts.None;
            }

            Excel.Sheets sheets = workbook.Sheets;
            if (!TryFindTable(sheets, out Excel.Worksheet? _, out Excel.ListObject? table) || table is null)
            {
                return SelectionFacts.None;
            }

            Excel.Range? selection = GetSelection(application) as Excel.Range;
            Excel.Range? active = GetActiveCell(application);
            Excel.Range? range = selection ?? active;
            return range is null || !TryResolveToBody(table, range, out int? bodyIndex, out int rowCount, out bool isHeader)
                ? SelectionFacts.None
                : new SelectionFacts(
                    true,
                    !isHeader && rowCount == 1 && bodyIndex.HasValue,
                    bodyIndex,
                    isHeader,
                    rowCount);
        }
        catch
        {
            return SelectionFacts.None;
        }
#pragma warning restore CA1031
    }

    /// <inheritdoc />
    public IDisposable SubscribeSelectionChanged(Action handler)
    {
        ArgumentNullException.ThrowIfNull(handler);

        Excel.AppEvents_Event? events = _events;
        if (events is null)
        {
            return InertSelectionSubscription.Instance;
        }

        SelectionSubscription subscription = new SelectionSubscription(events, handler);
        subscription.Attach();
        return subscription;
    }

    /// <summary>
    /// Returns the active cell. Test seam over the COM
    /// <c>Application.ActiveCell</c> property.
    /// </summary>
    /// <param name="application">The Excel application.</param>
    /// <returns>The active cell, or null.</returns>
    internal virtual Excel.Range? GetActiveCell(Excel.Application application) =>
        application.ActiveCell;

    /// <summary>
    /// Returns the current selection. Test seam over the COM
    /// <c>Application.Selection</c> property. The selection — not the active
    /// cell — carries the selection's EXTENT, so a multi-row selection is
    /// distinguishable from a single cell (R5.5 ruling 1: bespoke per row).
    /// A non-range selection (shape, chart) casts to null.
    /// </summary>
    /// <param name="application">The Excel application.</param>
    /// <returns>The selection range, or null.</returns>
    internal virtual object? GetSelection(Excel.Application application) =>
        application.Selection;

    /// <summary>
    /// Locates <c>tblGanttData</c>. Instance method over the same
    /// count/index seams the row mutators use, so contract tests substitute
    /// the graph without COM.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="worksheet">The worksheet holding the table, or null.</param>
    /// <param name="table">The table, or null.</param>
    /// <returns>True when the table was found.</returns>
    internal bool TryFindTable(
        Excel.Sheets sheets,
        out Excel.Worksheet? worksheet,
        out Excel.ListObject? table)
    {
        worksheet = null;
        table = null;

        // CA1031: a failed table probe reports "no selection".
#pragma warning disable CA1031
        try
        {
            int sheetCount = GetSheetCount(sheets);
            for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
            {
                if (GetSheetAt(sheets, sheetIndex) is not Excel.Worksheet candidate)
                {
                    continue;
                }

                Excel.ListObjects listObjects = candidate.ListObjects;
                int tableCount = listObjects.Count;
                for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    Excel.ListObject located = GetTableAt(listObjects, tableIndex);
                    if (string.Equals(
                        located.Name,
                        GanttTableSchema.TableName,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = candidate;
                        table = located;
                        return true;
                    }
                }
            }
        }
        catch
        {
            worksheet = null;
            table = null;
            return false;
        }
#pragma warning restore CA1031

        return false;
    }

    /// <summary>Test seam over the COM <c>Sheets.Count</c> property.</summary>
    internal virtual int GetSheetCount(Excel.Sheets sheets) => sheets.Count;

    /// <summary>Test seam over the COM parameterised <c>Sheets.Item</c> property.</summary>
    internal virtual object GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    /// <summary>Test seam over the COM parameterised <c>ListObjects.Item</c> property.</summary>
    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects listObjects, int index) =>
        listObjects[index];

    /// <summary>
    /// Resolves a range to a one-based table body index plus the selected row
    /// span. Mirrors the row inserter/mutator active-cell seam: the header row
    /// resolves as header, a body row as its one-based index, anything else as
    /// unresolvable.
    /// </summary>
    private bool TryResolveToBody(
        Excel.ListObject table,
        Excel.Range active,
        out int? bodyIndex,
        out int rowCount,
        out bool isHeader)
    {
        bodyIndex = null;
        rowCount = 0;
        isHeader = false;

        Excel.Range? headerRow = GetHeaderRow(table);
        Excel.Range? body = GetDataBodyRange(table);
        if (headerRow is null)
        {
            return false;
        }

        if (Intersects(headerRow, active))
        {
            isHeader = true;
            rowCount = GetRowCount(active);
            return true;
        }

        if (body is null)
        {
            return false;
        }

        Excel.Range? intersection = Intersect(body, active);
        if (intersection is null)
        {
            return false;
        }

        rowCount = GetRowCount(active);
        int firstWorksheetRow = GetRangeRow(body);
        int activeWorksheetRow = GetRangeRow(intersection);
        if (firstWorksheetRow <= 0 || activeWorksheetRow < firstWorksheetRow)
        {
            return false;
        }

        bodyIndex = activeWorksheetRow - firstWorksheetRow + 1;
        return true;
    }

    /// <summary>Test seam over the COM <c>ListObject.HeaderRowRange</c> property.</summary>
    internal virtual Excel.Range? GetHeaderRow(Excel.ListObject table) => table.HeaderRowRange;

    /// <summary>Test seam over the COM <c>ListObject.DataBodyRange</c> property.</summary>
    internal virtual Excel.Range? GetDataBodyRange(Excel.ListObject table) => table.DataBodyRange;

    /// <summary>Test seam over the COM <c>Application.Intersect</c> method.</summary>
    internal virtual Excel.Range? Intersect(Excel.Range first, Excel.Range second) =>
        _application?.Intersect(first, second, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

    /// <summary>Test seam over the COM <c>Range.Row</c> property.</summary>
    internal virtual int GetRangeRow(Excel.Range range) => range.Row;

    /// <summary>Test seam over the COM <c>Range.Rows.Count</c> property.</summary>
    internal virtual int GetRowCount(Excel.Range range) =>
        range.Rows is Excel.Range rows ? rows.Count : 1;

    private bool Intersects(Excel.Range first, Excel.Range second) => Intersect(first, second) is not null;

    /// <summary>
    /// The one owned resource: the event connection. It adds exactly one
    /// handler to <c>SheetSelectionChange</c> and removes exactly that
    /// instance (the <c>WorkbookStateSubscription</c> pattern, one handler).
    /// </summary>
    private sealed class SelectionSubscription(Excel.AppEvents_Event events, Action handler) : IDisposable
    {
        private readonly Excel.AppEvents_Event _events = events;
        private readonly Action _handler = handler;
        private Excel.AppEvents_SheetSelectionChangeEventHandler? _onSelectionChange;
        private EventHandlerState _state;
        private bool _disposed;

        internal void Attach()
        {
            _onSelectionChange = new Excel.AppEvents_SheetSelectionChangeEventHandler(OnSelectionChange);
            _events.SheetSelectionChange += _onSelectionChange;
            _state = EventHandlerState.Attached;
        }

        private void OnSelectionChange(object sh, Excel.Range target)
        {
            _ = sh;
            _ = target;

            // CA1031: a throw here escapes into Excel's event dispatch.
#pragma warning disable CA1031
            try
            {
                _handler();
            }
            catch
            {
                // Intentionally empty: the subscription must never fault Excel.
            }
#pragma warning restore CA1031
        }

        /// <inheritdoc />
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            // CA1031: teardown must never propagate into Excel.
#pragma warning disable CA1031
            try
            {
                if (_state == EventHandlerState.Attached && _onSelectionChange is not null)
                {
                    _state = EventHandlerState.Detaching;
                    _events.SheetSelectionChange -= _onSelectionChange;
                    _state = EventHandlerState.Detached;
                    _onSelectionChange = null;
                }
            }
            catch
            {
                _state = EventHandlerState.DetachFailed;
            }
#pragma warning restore CA1031
        }
    }

    /// <summary>
    /// The subscription returned when no application object is available. It is
    /// already inert, so callers need no special case.
    /// </summary>
    private sealed class InertSelectionSubscription : IDisposable
    {
        private InertSelectionSubscription()
        {
        }

        /// <summary>Gets the shared instance; the type holds no state.</summary>
        internal static InertSelectionSubscription Instance { get; } = new();

        /// <inheritdoc />
        public void Dispose()
        {
            // Intentionally empty: nothing was attached.
        }
    }
}
