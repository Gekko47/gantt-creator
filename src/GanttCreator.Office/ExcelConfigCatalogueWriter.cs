using System.Globalization;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IConfigCatalogueWriter"/> over the Excel application
/// object. Materialises the five code-owned catalogue tables (ADR-0007 D2)
/// on the <c>_GanttCreatorConfig</c> worksheet; preserves user-authored
/// style rows and present setting values on regeneration (ADR-0007 D4).
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <remarks>
/// <para>
/// COM ownership: every proxy is held in a local and used without chained
/// member expressions, matching <see cref="ExcelWorkbookInitialiser"/>. The
/// <c>internal virtual</c> accessors isolate the COM parameterised
/// properties (indexers and <c>ListObjects.Add</c>), so contract tests can
/// substitute them (CS0855); the real behaviour is exercised by the tagged
/// live-Office integration test.
/// </para>
/// <para>
/// Mutation order (ADR-0007 D4): all read-only checks and the preservation
/// capture first, then per-table delete-and-recreate. <c>tblGanttTypes</c>,
/// <c>tblGanttMetrics</c>, and <c>tblGanttConfig</c> hold no user content
/// and regenerate outright; <c>tblGanttStyles</c> preserves non-built-in
/// rows and <c>tblGanttSettings</c> preserves present values, appending
/// missing defaults. The port writes only on the configuration worksheet —
/// never schedule data, never the visible sheet, never a second helper
/// sheet.
/// </para>
/// </remarks>
public class ExcelConfigCatalogueWriter(object? application) : IConfigCatalogueWriter
{
    private readonly Application? _application = application as Application;

    /// <summary>
    /// The preservation data captured before any mutation: existing setting
    /// values, user-authored style rows, and the workbook ID.
    /// </summary>
    /// <param name="Settings">Existing key/value pairs from <c>tblGanttSettings</c>.</param>
    /// <param name="UserStyleRows">
    /// Full rows of <c>tblGanttStyles</c> whose <c>StyleKey</c> is not a
    /// built-in preset key.
    /// </param>
    /// <param name="WorkbookId">The stored workbook ID, when present.</param>
    private sealed record CataloguePreservation(
        IReadOnlyDictionary<string, string> Settings,
        IReadOnlyList<object?[]> UserStyleRows,
        string? WorkbookId);

    /// <inheritdoc />
    public ConfigWriteOutcome Write()
    {
        Application? application = _application;
        if (application is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook);
        }

        // One proxy per local: no chained member expressions
        // (docs/02-ARCHITECTURE.md COM ownership).
        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;

        // Read-only check 1: the configuration worksheet must exist —
        // catalogues are materialised during initialise (initialise itself
        // refuses when the sheet already exists, so a missing sheet here is
        // a call-order error, not a repair trigger).
        Worksheet? config = FindConfigSheet(sheets);
        if (config is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.ConfigSheetMissing);
        }

        // Read-only check 2: worksheet or workbook-structure protection
        // would block the table writes; refuse before any mutation.
        if (config.ProtectContents || workbook.ProtectStructure)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.TargetProtected);
        }

        // Read-only phase: capture the user content the write must preserve
        // (ADR-0007 D4) before any table is touched.
        CataloguePreservation preservation = ReadPreservation(config);

        // Mutation phase: per-table delete-and-recreate under one
        // DisplayAlerts save/restore (table deletion prompts).
        var original = _application!.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;
            WriteOrReplaceTable(
                config,
                GanttCatalogues.TypesAnchor,
                GanttCatalogues.TypesTableName,
                GanttCatalogues.TypesHeaders,
                BuildTypeRows());
            WriteOrReplaceTable(
                config,
                GanttCatalogues.StylesAnchor,
                GanttCatalogues.StylesTableName,
                GanttCatalogues.StylesHeaders,
                BuildStyleRows(preservation.UserStyleRows));
            WriteOrReplaceTable(
                config,
                GanttCatalogues.MetricsAnchor,
                GanttCatalogues.MetricsTableName,
                GanttCatalogues.MetricsHeaders,
                BuildMetricRows());
            WriteOrReplaceTable(
                config,
                GanttCatalogues.SettingsAnchor,
                GanttCatalogues.SettingsTableName,
                GanttCatalogues.SettingsHeaders,
                BuildSettingRows(preservation.Settings));
            WriteOrReplaceTable(
                config,
                GanttCatalogues.ConfigAnchor,
                GanttCatalogues.ConfigTableName,
                GanttCatalogues.ConfigHeaders,
                BuildConfigRows(preservation.WorkbookId));
        }
        finally
        {
            _application.DisplayAlerts = original;
        }

        return ConfigWriteOutcome.Ok();
    }

    /// <summary>
    /// Finds <c>_GanttCreatorConfig</c> by exact name. Chart sheets carry no
    /// worksheets semantics and are skipped via the worksheet cast.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <returns>The configuration worksheet, or <see langword="null"/>.</returns>
    private Worksheet? FindConfigSheet(Sheets sheets)
    {
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            var sheet = GetSheetAt(sheets, index);
            if (sheet is Worksheet worksheet
                && string.Equals(
                    worksheet.Name,
                    GanttWorkbookContract.ConfigSheetName,
                    StringComparison.Ordinal))
            {
                return worksheet;
            }
        }

        return null;
    }

    /// <summary>
    /// Captures the user content the write must preserve (ADR-0007 D4):
    /// present setting values, non-built-in style rows, and the stored
    /// workbook ID. Read-only; runs before any mutation.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <returns>The preservation snapshot; absent tables contribute nothing.</returns>
    private CataloguePreservation ReadPreservation(Worksheet config)
    {
        Dictionary<string, string> settings = ReadKeyValues(config, GanttCatalogues.SettingsTableName);
        List<object?[]> userStyles = ReadUserStyleRows(config);
        var workbookId = ReadConfigValue(config, GanttCatalogues.ConfigWorkbookIdKey);
        return new CataloguePreservation(settings, userStyles, workbookId);
    }

    /// <summary>
    /// Reads an existing key/value table on the configuration worksheet as a
    /// key→value map; an absent table or body contributes an empty map.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="tableName">The table to read.</param>
    /// <returns>The key→value map.</returns>
    private Dictionary<string, string> ReadKeyValues(Worksheet config, string tableName)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        ListObject? table = FindTable(config, tableName);
        if (table is null)
        {
            return map;
        }

        List<object?[]> rows = ReadBodyRows(table);
        foreach (var row in rows)
        {
            map[ToText(row.ElementAtOrDefault(0))] = ToText(row.ElementAtOrDefault(1));
        }

        return map;
    }

    /// <summary>
    /// Reads the rows of <c>tblGanttStyles</c> whose <c>StyleKey</c> is not
    /// a built-in preset key. Read-only.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <returns>The user rows in body order.</returns>
    private List<object?[]> ReadUserStyleRows(Worksheet config)
    {
        List<object?[]> userRows = [];
        ListObject? table = FindTable(config, GanttCatalogues.StylesTableName);
        if (table is null)
        {
            return userRows;
        }

        var builtInKeys = GanttCatalogues.StylePresets
            .Select(preset => preset.StyleKey)
            .ToHashSet(StringComparer.Ordinal);
        foreach (var row in ReadBodyRows(table))
        {
            var styleKey = ToText(row.ElementAtOrDefault(0));
            if (styleKey.Length > 0 && !builtInKeys.Contains(styleKey))
            {
                userRows.Add(row);
            }
        }

        return userRows;
    }

    /// <summary>
    /// Reads one <c>tblGanttConfig</c> value by key; <see langword="null"/>
    /// when the table or key is absent.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="key">The config key.</param>
    /// <returns>The value text, or <see langword="null"/>.</returns>
    private string? ReadConfigValue(Worksheet config, string key)
    {
        Dictionary<string, string> values = ReadKeyValues(config, GanttCatalogues.ConfigTableName);
        return values.TryGetValue(key, out var value) ? value : null;
    }

    /// <summary>
    /// Bulk-reads a table body through the seams into 0-based row arrays.
    /// </summary>
    /// <param name="table">The table to read.</param>
    /// <returns>The body rows; empty when the table has no body.</returns>
    private List<object?[]> ReadBodyRows(ListObject table)
    {
        List<object?[]> rows = [];
        Excel.Range? body = GetTableBody(table);
        if (body is null)
        {
            return rows;
        }

        var raw = GetBodyValues(body);
        foreach (var row in ConvertMatrix(raw))
        {
            rows.Add(row);
        }

        return rows;
    }

    /// <summary>
    /// Converts a <c>Value2</c> payload into 0-based row arrays. Handles the
    /// 2D SAFEARRAY (the normal case, one-based from live Excel — iterated
    /// by lower bound like <see cref="ExcelConfigCatalogueReader"/>) and
    /// degrades other shapes to empty.
    /// </summary>
    /// <param name="raw">The <c>Value2</c> payload.</param>
    /// <returns>The rows.</returns>
    private static List<object?[]> ConvertMatrix(object? raw)
    {
        List<object?[]> rows = [];
        if (raw is not object[,] matrix)
        {
            return rows;
        }

        var rowLower = matrix.GetLowerBound(0);
        var rowUpper = matrix.GetUpperBound(0);
        var columnLower = matrix.GetLowerBound(1);
        var columnCount = matrix.GetLength(1);
        for (var row = rowLower; row <= rowUpper; row++)
        {
            var cells = new object?[columnCount];
            for (var column = 0; column < columnCount; column++)
            {
                cells[column] = matrix[row, columnLower + column];
            }

            rows.Add(cells);
        }

        return rows;
    }

    /// <summary>
    /// Deletes the named table when it exists, then writes the header row
    /// plus the data rows as one array assignment and creates the table over
    /// the extent (one Add call). Mutates only the configuration worksheet.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="anchor">The anchor cell (e.g. <c>"A1"</c>).</param>
    /// <param name="tableName">The contract table name.</param>
    /// <param name="headers">The contract header row.</param>
    /// <param name="dataRows">The data rows (already merged with user content).</param>
    private void WriteOrReplaceTable(
        Worksheet config,
        string anchor,
        string tableName,
        string[] headers,
        List<object?[]> dataRows)
    {
        ListObject? existing = FindTable(config, tableName);
        existing?.Delete();


        (var anchorRow, var anchorColumn) = ParseAnchor(anchor);
        var rowCount = dataRows.Count + 1;
        var columnCount = headers.Length;
        Excel.Range anchorCell = GetCellRange(config, anchorRow, anchorColumn);
        Excel.Range extent = GetResizedRange(anchorCell, rowCount, columnCount);

        // CA1814: Excel's Range.Value2 accepts only a rectangular object
        // array (a COM SAFEARRAY of VARIANT); the multidimensional form is
        // the requirement, not a style choice.
#pragma warning disable CA1814
        var matrix = new object[rowCount, columnCount];
#pragma warning restore CA1814
        for (var column = 0; column < columnCount; column++)
        {
            matrix[0, column] = headers[column];
        }

        for (var row = 0; row < dataRows.Count; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                matrix[row + 1, column] = dataRows[row][column] ?? string.Empty;
            }
        }

        extent.Value2 = matrix;
        ListObjects listObjects = GetListObjects(config);
        ListObject table = AddTable(listObjects, extent);
        table.Name = tableName;
    }

    /// <summary>
    /// Finds a table by exact name on the configuration worksheet.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="tableName">The contract table name.</param>
    /// <returns>The table, or <see langword="null"/>.</returns>
    private ListObject? FindTable(Worksheet config, string tableName)
    {
        ListObjects listObjects = GetListObjects(config);
        var count = listObjects.Count;
        for (var index = 1; index <= count; index++)
        {
            ListObject table = GetTableAt(listObjects, index);
            if (string.Equals(table.Name, tableName, StringComparison.Ordinal))
            {
                return table;
            }
        }

        return null;
    }

    /// <summary>
    /// Builds the <c>tblGanttTypes</c> data rows from the code-owned type
    /// catalogue. No user content exists on this table by contract.
    /// </summary>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildTypeRows() =>
    [
        .. GanttCatalogues.TypeRows
            .Select(row => new object?[]
            {
                row.TypeName,
                row.DisplayName,
                row.Kind,
                row.DateMode,
                row.DefaultStyleKey,
                row.ColourCapability,
                row.RequiresStyleKey,
                row.AllowedLabelPositions,
            }),
    ];

    /// <summary>
    /// Builds the <c>tblGanttStyles</c> data rows: the built-in presets plus
    /// the preserved user rows (ADR-0007 D4).
    /// </summary>
    /// <param name="userRows">The preserved non-built-in rows.</param>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildStyleRows(IReadOnlyList<object?[]> userRows)
    {
        List<object?[]> rows =
        [
            .. GanttCatalogues.StylePresets
                .Select(preset => new object?[]
                {
                    preset.StyleKey,
                    preset.DisplayName,
                    preset.FillColour,
                    preset.StrokeColour,
                    preset.HatchPattern.ToString(),
                    preset.HatchPitchPt,
                    preset.HatchLinePt,
                    preset.TextColour,
                    preset.StandardOutlinePt,
                    preset.ActivityHeightPt,
                    preset.MilestoneSizePt,
                    preset.DefaultLabelPosition.ToString(),
                    string.Join(
                        " ",
                        preset.AllowedLabelPositions
                            .Select(position => position.ToString())
                            .OrderBy(name => name, StringComparer.Ordinal)),
                    preset.ColourCapability.ToString(),
                }),
        ];
        rows.AddRange(userRows);
        return rows;
    }

    /// <summary>
    /// Builds the <c>tblGanttMetrics</c> data rows. No user content exists
    /// on this table by contract.
    /// </summary>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildMetricRows() =>
    [
        .. GanttCatalogues.Metrics
            .Select(metric => new object?[]
            {
                metric.Name,
                metric.DefaultValue,
                metric.Minimum,
                metric.Maximum,
            }),
    ];

    /// <summary>
    /// Builds the <c>tblGanttSettings</c> data rows: present values win and
    /// absent keys get the code default, so the table always holds exactly
    /// the approved key set (ADR-0007 D2/D3). A key that is not in the
    /// approved set is not carried forward: the reader validates the exact
    /// key set, so preserving an unapproved key would leave the workbook
    /// permanently unreadable after every regeneration. Dialog-authored
    /// keys are a future decision owned by the settings-dialog work item,
    /// which must bump the schema version and catalogue hash first
    /// (ADR-0007 D5/D6).
    /// </summary>
    /// <param name="existing">The preserved key/value map.</param>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildSettingRows(IReadOnlyDictionary<string, string> existing)
    {
        List<object?[]> rows = [];
        foreach (GanttSettingDefinition setting in GanttCatalogues.Settings)
        {
            rows.Add(
            [
                setting.Key,
                existing.TryGetValue(setting.Key, out var value) ? value : setting.DefaultValue,
            ]);
        }

        return rows;
    }

    /// <summary>
    /// Builds the <c>tblGanttConfig</c> data rows: the schema version, the
    /// catalogue hash, the preserved-or-new workbook ID, and the add-in
    /// version last used.
    /// </summary>
    /// <param name="existingWorkbookId">The preserved workbook ID, or <see langword="null"/>.</param>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildConfigRows(string? existingWorkbookId)
    {
        var workbookId = existingWorkbookId;
        if (string.IsNullOrWhiteSpace(workbookId))
        {
            workbookId = Guid.NewGuid().ToString("D", CultureInfo.InvariantCulture);
        }

        return
        [
            [GanttCatalogues.ConfigSchemaVersionKey, GanttSchemaVersion.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture)],
            [GanttCatalogues.ConfigCatalogueHashKey, GanttCatalogues.ComputeCatalogueHash()],
            [GanttCatalogues.ConfigWorkbookIdKey, workbookId],
            [GanttCatalogues.ConfigAddInVersionKey, VersionInfo.SemanticVersion],
        ];
    }

    /// <summary>
    /// Converts cell text: <see langword="null"/> and Excel error values
    /// become empty; everything else is invariant text.
    /// </summary>
    /// <param name="value">The cell value.</param>
    /// <returns>The text.</returns>
    private static string ToText(object? value) =>
        value switch
        {
            null => string.Empty,
            string text => text,
            double number when double.IsNaN(number) || double.IsInfinity(number) => string.Empty,
            bool flag => flag ? "TRUE" : "FALSE",
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty,
        };

    /// <summary>
    /// Parses an A1-style anchor (e.g. <c>"A1"</c>, <c>"AB20"</c>) into
    /// one-based row and column indices, culture-invariant.
    /// </summary>
    /// <param name="anchor">The anchor address.</param>
    /// <returns>The (row, column) pair.</returns>
    private static (int Row, int Column) ParseAnchor(string anchor)
    {
        var letterCount = 0;
        while (letterCount < anchor.Length && char.IsAsciiLetterUpper(anchor[letterCount]))
        {
            letterCount++;
        }

        var column = 0;
        for (var index = 0; index < letterCount; index++)
        {
            column = (column * 26) + anchor[index] - 'A' + 1;
        }

        var row = int.Parse(
            anchor[letterCount..],
            NumberStyles.None,
            CultureInfo.InvariantCulture);
        return (row, column);
    }

    /// <summary>
    /// Returns the sheet at the one-based index. Test seam over the COM
    /// parameterised <c>Sheets.Item</c> property.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="index">The one-based sheet index.</param>
    /// <returns>The sheet at the index.</returns>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>
    /// Returns a worksheet's list objects. Test seam over the COM property.
    /// </summary>
    /// <param name="worksheet">The configuration worksheet.</param>
    /// <returns>The list objects.</returns>
    internal virtual ListObjects GetListObjects(Worksheet worksheet) => worksheet.ListObjects;

    /// <summary>
    /// Returns the list object at the one-based index. Test seam over the
    /// COM parameterised <c>ListObjects.Item</c> property.
    /// </summary>
    /// <param name="listObjects">The worksheet's list objects.</param>
    /// <param name="index">The one-based table index.</param>
    /// <returns>The list object at the index.</returns>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) => listObjects[index];

    /// <summary>
    /// Adds a table over the source range with header-name behaviour. Test
    /// seam over the COM <c>ListObjects.Add</c> method.
    /// </summary>
    /// <param name="listObjects">The worksheet's list objects.</param>
    /// <param name="source">The extent range (header row plus data rows).</param>
    /// <returns>The created list object.</returns>
    internal virtual ListObject AddTable(ListObjects listObjects, object source) =>
        listObjects.Add(
            XlListObjectSourceType.xlSrcRange,
            source,
            Type.Missing,
            XlYesNoGuess.xlYes,
            Type.Missing);

    /// <summary>
    /// Returns the cell at the one-based row and column. Test seam over the
    /// COM parameterised <c>Range.Item</c> property.
    /// </summary>
    /// <param name="worksheet">The configuration worksheet.</param>
    /// <param name="row">The one-based row.</param>
    /// <param name="column">The one-based column.</param>
    /// <returns>The cell range.</returns>
    internal virtual Excel.Range GetCellRange(Worksheet worksheet, int row, int column) =>
        worksheet.Cells[row, column];

    /// <summary>
    /// Resizes a range. Test seam over the COM parameterised
    /// <c>Range.Resize</c> property.
    /// </summary>
    /// <param name="range">The anchor range.</param>
    /// <param name="rows">The row count.</param>
    /// <param name="columns">The column count.</param>
    /// <returns>The resized range.</returns>
    internal virtual Excel.Range GetResizedRange(Excel.Range range, int rows, int columns) =>
        range.Resize[rows, columns];

    /// <summary>
    /// Returns a table's body range. Test seam over the COM property.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <returns>The body range, or <see langword="null"/> when empty.</returns>
    internal virtual Excel.Range? GetTableBody(ListObject table) => table.DataBodyRange;

    /// <summary>
    /// Returns the bulk <c>Value2</c> payload of a range. Test seam so
    /// contract tests inject <c>object[,]</c> without COM.
    /// </summary>
    /// <param name="body">The range.</param>
    /// <returns>The <c>Value2</c> payload.</returns>
    internal virtual object? GetBodyValues(Excel.Range body) => body.Value2;
}
