using System.Globalization;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IConfigCatalogueReader"/> over the Excel application
/// object. Bulk-reads the five catalogue tables on the
/// <c>_GanttCreatorConfig</c> worksheet via one <c>DataBodyRange.Value2</c>
/// call per table and validates them against the code-owned Core catalogue
/// (ADR-0007 D2/D6/D7). Never writes cells, creates sheets, or changes
/// application state.
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
/// member expressions, matching <see cref="ExcelGanttTableReader"/>. The
/// <c>internal virtual</c> accessors isolate the COM parameterised
/// properties (indexers), so contract tests can substitute them (CS0855);
/// the real behaviour is exercised by the tagged live-Office integration
/// test.
/// </para>
/// <para>
/// Validation order (checklist D): locate every expected table and its
/// headers first, then validate row counts, then validate cell contents,
/// then compare the stored hash. The first mismatch short-circuits with the
/// exact typed refusal.
/// </para>
/// </remarks>
public class ExcelConfigCatalogueReader(object? application) : IConfigCatalogueReader
{
    private readonly Application? _application = application as Application;

    /// <inheritdoc />
    public ConfigReadOutcome Read()
    {
        Application? application = _application;
        if (application is null)
        {
            return ConfigReadOutcome.Refused(ConfigReadRefusalReason.NoActiveWorkbook);
        }

        // One proxy per local: no chained member expressions
        // (docs/02-ARCHITECTURE.md COM ownership).
        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return ConfigReadOutcome.Refused(ConfigReadRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        Worksheet? config = FindConfigSheet(sheets);
        if (config is null)
        {
            return ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing);
        }

        // Bulk-read every table first through the seams; validation below
        // is pure (no COM) and short-circuits on the first mismatch.
        Dictionary<string, string>? settings = null;
        string? workbookId = null;
        if (ReadTable(config, GanttCatalogues.TypesTableName, GanttCatalogues.TypesHeaders, out List<object?[]> typeRows) is { } found)
        {
            return ConfigReadOutcome.Refused(found);
        }

        if (ReadTable(config, GanttCatalogues.StylesTableName, GanttCatalogues.StylesHeaders, out List<object?[]> styleRows) is { } stylesFound)
        {
            return ConfigReadOutcome.Refused(stylesFound);
        }

        if (ReadTable(config, GanttCatalogues.MetricsTableName, GanttCatalogues.MetricsHeaders, out List<object?[]> metricRows) is { } metricsFound)
        {
            return ConfigReadOutcome.Refused(metricsFound);
        }

        if (ReadTable(config, GanttCatalogues.SettingsTableName, GanttCatalogues.SettingsHeaders, out List<object?[]> settingRows) is { } settingsFound)
        {
            return ConfigReadOutcome.Refused(settingsFound);
        }

        if (ReadTable(config, GanttCatalogues.ConfigTableName, GanttCatalogues.ConfigHeaders, out List<object?[]> configRows) is { } configFound)
        {
            return ConfigReadOutcome.Refused(configFound);
        }

        ConfigReadRefusalReason? refusal = ValidateTypes(typeRows)
            ?? ValidateStyles(styleRows)
            ?? ValidateMetrics(metricRows)
            ?? ValidateSettings(settingRows, out settings)
            ?? ValidateConfig(configRows, out workbookId);
        return refusal is not null
            ? ConfigReadOutcome.Refused(refusal.Value)
            : ConfigReadOutcome.Ok(workbookId!, settings!);
    }

    /// <summary>
    /// Finds <c>_GanttCreatorConfig</c> by exact name. Chart sheets are
    /// skipped via the worksheet cast.
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
    /// Bulk-reads one catalogue table's headers and body through the seams
    /// into 0-based row arrays. Validates the table's existence and headers;
    /// row counts and cell contents are validated later against the Core
    /// catalogue.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="tableName">The contract table name.</param>
    /// <param name="expectedHeaders">The contract header row.</param>
    /// <param name="rows">The body rows; empty on a refusal.</param>
    /// <returns>
    /// The refusal reason when the table is missing or header-mismatched;
    /// otherwise <see langword="null"/> with the rows populated.
    /// </returns>
    private ConfigReadRefusalReason? ReadTable(
        Worksheet config,
        string tableName,
        string[] expectedHeaders,
        out List<object?[]> rows)
    {
        rows = [];
        ListObject? table = FindTable(config, tableName);
        if (table is null)
        {
            return ConfigReadRefusalReason.TableMissing;
        }

        ListColumns columns = table.ListColumns;
        var columnCount = columns.Count;
        if (columnCount != expectedHeaders.Length)
        {
            return ConfigReadRefusalReason.HeaderMismatch;
        }

        for (var index = 1; index <= columnCount; index++)
        {
            ListColumn column = GetColumnAt(columns, index);
            if (!string.Equals(column.Name, expectedHeaders[index - 1], StringComparison.Ordinal))
            {
                return ConfigReadRefusalReason.HeaderMismatch;
            }
        }

        Excel.Range? body = GetTableBody(table);
        if (body is null)
        {
            return null;
        }

        var raw = GetBodyValues(body);
        if (raw is object[,] matrix)
        {
            // Live Excel returns DataBodyRange.Value2 as a one-based
            // SAFEARRAY (the R2.7 live gate crashed indexing [0, ...]; the
            // R2.4 reader and both contract fakes pin the same shape), but
            // iterate by lower bound so zero-based payloads convert too and
            // the two shapes stay indistinguishable to callers.
            var rowLower = matrix.GetLowerBound(0);
            var rowUpper = matrix.GetUpperBound(0);
            var columnLower = matrix.GetLowerBound(1);
            var payloadColumns = matrix.GetLength(1);
            for (var row = rowLower; row <= rowUpper; row++)
            {
                var cells = new object?[payloadColumns];
                for (var column = 0; column < payloadColumns; column++)
                {
                    cells[column] = matrix[row, columnLower + column];
                }

                rows.Add(cells);
            }
        }

        return null;
    }

    /// <summary>
    /// Validates the <c>tblGanttTypes</c> rows against the projected
    /// <see cref="EntityTypeCatalog"/> rows (ADR-0007 D2/D7).
    /// </summary>
    /// <param name="rows">The body rows.</param>
    /// <returns>The first mismatch, or <see langword="null"/> when valid.</returns>
    private static ConfigReadRefusalReason? ValidateTypes(List<object?[]> rows)
    {
        if (rows.Count != GanttCatalogues.TypeRows.Count)
        {
            return ConfigReadRefusalReason.RowCountMismatch;
        }

        for (var index = 0; index < rows.Count; index++)
        {
            GanttTypeCatalogueRow expected = GanttCatalogues.TypeRows[index];
            string[] expectedCells =
            [
                expected.TypeName,
                expected.DisplayName,
                expected.Kind,
                expected.DateMode,
                expected.DefaultStyleKey,
                expected.ColourCapability,
                expected.RequiresStyleKey,
                expected.AllowedLabelPositions,
            ];
            if (!RowMatches(rows[index], expectedCells))
            {
                return ConfigReadRefusalReason.CatalogueMismatch;
            }
        }

        return null;
    }

    /// <summary>
    /// Validates the <c>tblGanttStyles</c> built-in rows against the
    /// resolved Core presets (ADR-0007 D7). Additional user rows after the
    /// built-ins are permitted (ADR-0007 D4).
    /// </summary>
    /// <param name="rows">The body rows.</param>
    /// <returns>The first mismatch, or <see langword="null"/> when valid.</returns>
    private static ConfigReadRefusalReason? ValidateStyles(List<object?[]> rows)
    {
        if (rows.Count < GanttCatalogues.StylePresets.Count)
        {
            return ConfigReadRefusalReason.RowCountMismatch;
        }

        for (var index = 0; index < GanttCatalogues.StylePresets.Count; index++)
        {
            GanttStylePreset preset = GanttCatalogues.StylePresets[index];
            if (rows[index].Length < GanttCatalogues.StylesHeaders.Length)
            {
                return ConfigReadRefusalReason.RowCountMismatch;
            }

            // Colour format first. A malformed colour is the more specific
            // finding, and with exact-match first BadColourFormat would be
            // unreachable for every style row (a malformed colour can never
            // equal a valid catalogue colour). Empty is permitted: the
            // stroke and fill tokens are optional for several presets.
            foreach (var colourColumn in new[] { 2, 3, 7 })
            {
                var colourText = ToText(rows[index][colourColumn]);
                if (colourText.Length > 0 && !GanttColourToken.IsValidHex(colourText))
                {
                    return ConfigReadRefusalReason.BadColourFormat;
                }
            }

            string[] textCells =
            [
                preset.StyleKey,
                preset.DisplayName,
                preset.FillColour,
                preset.StrokeColour,
                preset.HatchPattern.ToString(),
            ];
            for (var column = 0; column < textCells.Length; column++)
            {
                if (!CellMatches(rows[index][column], textCells[column]))
                {
                    return ConfigReadRefusalReason.CatalogueMismatch;
                }
            }

            double[] expectedNumbers =
            [
                preset.HatchPitchPt,
                preset.HatchLinePt,
                preset.StandardOutlinePt,
                preset.ActivityHeightPt,
                preset.MilestoneSizePt,
            ];

            // The numeric columns are not contiguous: TextColour (7) sits
            // between HatchLinePt (6) and StandardOutlinePt (8), and
            // MilestoneSizePt (10) closes the original row. See StylesHeaders.
            int[] numberColumns = [5, 6, 8, 9, 10];
            for (var column = 0; column < expectedNumbers.Length; column++)
            {
                if (!CellMatchesDouble(rows[index][numberColumns[column]], expectedNumbers[column]))
                {
                    return ConfigReadRefusalReason.CatalogueMismatch;
                }
            }

            string[] capabilityCells =
            [
                preset.DefaultLabelPosition.ToString(),
                string.Join(
                    " ",
                    preset.AllowedLabelPositions
                        .Select(position => position.ToString())
                        .OrderBy(name => name, StringComparer.Ordinal)),
                preset.ColourCapability.ToString(),
            ];
            for (var column = 0; column < capabilityCells.Length; column++)
            {
                if (!CellMatches(rows[index][11 + column], capabilityCells[column]))
                {
                    return ConfigReadRefusalReason.CatalogueMismatch;
                }
            }
        }

        for (var index = GanttCatalogues.StylePresets.Count; index < rows.Count; index++)
        {
            if (ValidateUserStyleCapabilities(rows[index]) is { } userStyleRefusal)
            {
                return userStyleRefusal;
            }
        }

        return null;
    }

    private static ConfigReadRefusalReason? ValidateUserStyleCapabilities(object?[] row)
    {
        if (row.Length < GanttCatalogues.StylesHeaders.Length)
        {
            return ConfigReadRefusalReason.RowCountMismatch;
        }

        var defaultText = ToText(row[11]);
        if (!Enum.TryParse(defaultText, ignoreCase: false, out GanttLabelPosition defaultPosition)
            || !Enum.IsDefined(defaultPosition))
        {
            return ConfigReadRefusalReason.ValueOutOfRange;
        }

        var allowedTexts = ToText(row[12])
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (allowedTexts.Length == 0)
        {
            return ConfigReadRefusalReason.ValueOutOfRange;
        }

        var allowed = new HashSet<GanttLabelPosition>();
        foreach (var allowedText in allowedTexts)
        {
            if (!Enum.TryParse(allowedText, ignoreCase: false, out GanttLabelPosition position)
                || !Enum.IsDefined(position)
                || !allowed.Add(position))
            {
                return ConfigReadRefusalReason.ValueOutOfRange;
            }
        }

        return !allowed.Contains(defaultPosition)
            || !Enum.TryParse(ToText(row[13]), ignoreCase: false, out EntityColourCapability colourCapability)
            || !IsKnownColourCapability(colourCapability)
            ? ConfigReadRefusalReason.ValueOutOfRange
            : null;
    }

    private static bool IsKnownColourCapability(EntityColourCapability capability)
    {
        const EntityColourCapability known =
            EntityColourCapability.Fill | EntityColourCapability.Stroke |
            EntityColourCapability.Hatch | EntityColourCapability.StyleDefined;
        return (capability & ~known) == EntityColourCapability.None;
    }

    /// <summary>
    /// Validates the <c>tblGanttMetrics</c> rows against the code-owned
    /// tokens: exact names and defaults, every value finite and inside its
    /// valid range (ADR-0007 D2).
    /// </summary>
    /// <param name="rows">The body rows.</param>
    /// <returns>The first mismatch, or <see langword="null"/> when valid.</returns>
    private static ConfigReadRefusalReason? ValidateMetrics(List<object?[]> rows)
    {
        if (rows.Count != GanttCatalogues.Metrics.Count)
        {
            return ConfigReadRefusalReason.RowCountMismatch;
        }

        for (var index = 0; index < rows.Count; index++)
        {
            GanttMetricToken metric = GanttCatalogues.Metrics[index];
            var row = rows[index];
            if (row.Length < GanttCatalogues.MetricsHeaders.Length)
            {
                return ConfigReadRefusalReason.RowCountMismatch;
            }

            if (!CellMatches(row[0], metric.Name))
            {
                return ConfigReadRefusalReason.CatalogueMismatch;
            }

            if (!TryCellDouble(row[1], out var @default)
                || !TryCellDouble(row[2], out var minimum)
                || !TryCellDouble(row[3], out var maximum))
            {
                return ConfigReadRefusalReason.ValueOutOfRange;
            }

            // Internal consistency first (a tampered range is a range
            // refusal, not a drift finding), then exact catalogue equality.
            if (minimum > maximum || @default < minimum || @default > maximum)
            {
                return ConfigReadRefusalReason.ValueOutOfRange;
            }

            if (@default != metric.DefaultValue
                || minimum != metric.Minimum
                || maximum != metric.Maximum)
            {
                return ConfigReadRefusalReason.CatalogueMismatch;
            }
        }

        return null;
    }

    /// <summary>
    /// Validates the <c>tblGanttSettings</c> rows: the key set must be the
    /// approved keys (ADR-0007 D3) and every boolean-typed value must be
    /// <c>TRUE</c>/<c>FALSE</c>. Returns the effective map.
    /// </summary>
    /// <param name="rows">The body rows.</param>
    /// <param name="settings">The effective key/value map on success.</param>
    /// <returns>The first mismatch, or <see langword="null"/> when valid.</returns>
    private static ConfigReadRefusalReason? ValidateSettings(
        List<object?[]> rows,
        out Dictionary<string, string>? settings)
    {
        settings = null;
        if (rows.Count != GanttCatalogues.Settings.Count)
        {
            return ConfigReadRefusalReason.RowCountMismatch;
        }

        Dictionary<string, string> map = new(StringComparer.Ordinal);
        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            if (row.Length < GanttCatalogues.SettingsHeaders.Length)
            {
                return ConfigReadRefusalReason.RowCountMismatch;
            }

            GanttSettingDefinition expected = GanttCatalogues.Settings[index];
            if (!CellMatches(row[0], expected.Key))
            {
                return ConfigReadRefusalReason.CatalogueMismatch;
            }

            var value = ToText(row[1]);
            if (IsBooleanSetting(expected.Key) && value != "TRUE" && value != "FALSE")
            {
                return ConfigReadRefusalReason.ValueOutOfRange;
            }

            map[expected.Key] = value;
        }

        settings = map;
        return null;
    }

    /// <summary>
    /// Validates the <c>tblGanttConfig</c> metadata rows: the schema version
    /// must match <see cref="GanttSchemaVersion.CurrentSchemaVersion"/>, the
    /// stored hash must match the code-owned catalogue hash (ADR-0007 D6),
    /// and the workbook ID must be present. Returns the workbook ID.
    /// </summary>
    /// <param name="rows">The body rows.</param>
    /// <param name="workbookId">The workbook ID on success.</param>
    /// <returns>The first mismatch, or <see langword="null"/> when valid.</returns>
    private static ConfigReadRefusalReason? ValidateConfig(
        List<object?[]> rows,
        out string? workbookId)
    {
        workbookId = null;
        if (rows.Count != 4)
        {
            return ConfigReadRefusalReason.RowCountMismatch;
        }

        Dictionary<string, string> map = new(StringComparer.Ordinal);
        foreach (var row in rows)
        {
            if (row.Length < GanttCatalogues.ConfigHeaders.Length)
            {
                return ConfigReadRefusalReason.RowCountMismatch;
            }

            map[ToText(row[0])] = ToText(row[1]);
        }

        return !map.TryGetValue(GanttCatalogues.ConfigSchemaVersionKey, out var schemaText)
            || schemaText != GanttSchemaVersion.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture)
            ? ConfigReadRefusalReason.CatalogueHashMismatch
            : ValidateStoredHash(map, ref workbookId);
    }

    /// <summary>
    /// Validates the stored catalogue hash against the code-owned catalogue
    /// hash (ADR-0007 D6).
    /// </summary>
    /// <param name="map">The parsed config rows.</param>
    /// <param name="workbookId">The assigned workbook ID.</param>
    /// <returns>The mismatch reason, or the workbook-ID validation outcome.</returns>
    private static ConfigReadRefusalReason? ValidateStoredHash(
        Dictionary<string, string> map,
        ref string? workbookId)
    {
        if (!map.TryGetValue(GanttCatalogues.ConfigCatalogueHashKey, out var hash)
            || !string.Equals(hash, GanttCatalogues.ComputeCatalogueHash(), StringComparison.OrdinalIgnoreCase))
        {
            return ConfigReadRefusalReason.CatalogueHashMismatch;
        }

        if (!map.TryGetValue(GanttCatalogues.ConfigWorkbookIdKey, out workbookId)
            || string.IsNullOrWhiteSpace(workbookId))
        {
            workbookId = null;
            return ConfigReadRefusalReason.CatalogueMismatch;
        }

        return null;
    }

    /// <summary>
    /// Determines whether a setting key carries a boolean value
    /// (<c>TRUE</c>/<c>FALSE</c>). The free-text keys are <c>ChartTitle</c>,
    /// <c>TimeScale</c>, <c>LegendPosition</c>, <c>PlotStartMode</c>, and
    /// <c>PlotFinishMode</c> (ADR-0007 D3).
    /// </summary>
    /// <param name="key">The setting key.</param>
    /// <returns><see langword="true"/> for the boolean-valued keys.</returns>
    private static bool IsBooleanSetting(string key) =>
        key is "ShowTitle"
            or "ExportIncludeDataPanel"
            or "ExportIncludeLegend"
            or "AlternateBanding"
            or "ShowMinorGrid"
            or "ShowMajorGrid";

    /// <summary>
    /// Determines whether a row matches the expected cells exactly
    /// (Ordinal text comparison).
    /// </summary>
    /// <param name="row">The read row.</param>
    /// <param name="expected">The expected cell texts.</param>
    /// <returns><see langword="true"/> when every cell matches.</returns>
    private static bool RowMatches(object?[] row, string[] expected)
    {
        if (row.Length < expected.Length)
        {
            return false;
        }

        for (var index = 0; index < expected.Length; index++)
        {
            if (!CellMatches(row[index], expected[index]))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Determines whether a cell holds the expected text: an exact Ordinal
    /// match for strings, an exact invariant match for finite numbers.
    /// </summary>
    /// <param name="cell">The cell value.</param>
    /// <param name="expected">The expected text.</param>
    /// <returns><see langword="true"/> when the cell matches.</returns>
    private static bool CellMatches(object? cell, string expected) =>
        cell switch
        {
            null => expected.Length == 0,
            string text => string.Equals(text, expected, StringComparison.Ordinal),
            double number => string.Equals(
                number.ToString("R", CultureInfo.InvariantCulture),
                expected,
                StringComparison.Ordinal),
            bool flag => string.Equals(
                flag ? "TRUE" : "FALSE",
                expected,
                StringComparison.Ordinal),
            _ => string.Equals(
                Convert.ToString(cell, CultureInfo.InvariantCulture),
                expected,
                StringComparison.Ordinal),
        };

    /// <summary>
    /// Determines whether a cell holds the expected finite double: Excel
    /// doubles compare exactly; other values degrade through invariant
    /// parsing.
    /// </summary>
    /// <param name="cell">The cell value.</param>
    /// <param name="expected">The expected value.</param>
    /// <returns><see langword="true"/> when the cell matches.</returns>
    private static bool CellMatchesDouble(object? cell, double expected) =>
        cell is double number
            ? !double.IsNaN(number) && !double.IsInfinity(number) && number == expected
            : TryCellDouble(cell, out var parsed) && parsed == expected;

    /// <summary>
    /// Parses a cell as a finite double.
    /// </summary>
    /// <param name="cell">The cell value.</param>
    /// <param name="value">The parsed value.</param>
    /// <returns><see langword="true"/> when the cell holds a finite double.</returns>
    private static bool TryCellDouble(object? cell, out double value) =>
        cell switch
        {
            double number when !double.IsNaN(number) && !double.IsInfinity(number) => True(number, out value),
            string text when double.TryParse(
                text,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var parsed)
                && !double.IsNaN(parsed)
                && !double.IsInfinity(parsed) => True(parsed, out value),
            _ => FalseDouble(out value),
        };

    private static bool True(double parsed, out double value)
    {
        value = parsed;
        return true;
    }

    private static bool FalseDouble(out double value)
    {
        value = 0;
        return false;
    }

    /// <summary>
    /// Converts cell text: <see langword="null"/> and NaN/infinity become
    /// empty; everything else is invariant text.
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
    /// Returns a worksheet's list objects. Test seam over the COM property.
    /// </summary>
    /// <param name="worksheet">The configuration worksheet.</param>
    /// <returns>The list objects.</returns>
    internal virtual ListObjects GetListObjects(Worksheet worksheet) => worksheet.ListObjects;

    /// <summary>
    /// Returns the sheet at the one-based index. Test seam over the COM
    /// parameterised <c>Sheets.Item</c> property.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="index">The one-based sheet index.</param>
    /// <returns>The sheet at the index.</returns>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>
    /// Returns the list object at the one-based index. Test seam over the
    /// COM parameterised <c>ListObjects.Item</c> property.
    /// </summary>
    /// <param name="listObjects">The worksheet's list objects.</param>
    /// <param name="index">The one-based table index.</param>
    /// <returns>The list object at the index.</returns>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) => listObjects[index];

    /// <summary>
    /// Returns the list column at the one-based index. Test seam over the
    /// COM parameterised <c>ListColumns.Item</c> property.
    /// </summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column at the index.</returns>
    internal virtual ListColumn GetColumnAt(ListColumns columns, int index) => columns[index];

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
