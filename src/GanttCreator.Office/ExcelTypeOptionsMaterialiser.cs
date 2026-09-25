using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>Live Excel adapter for the TypeOptions name and Type-column list rule.</summary>
public class ExcelTypeOptionsMaterialiser(
    object? application,
    IConfigCatalogueReader? catalogueReader = null,
    IWorksheetProtectionGuard? protectionGuard = null) : ITypeOptionsMaterialiser
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IConfigCatalogueReader _catalogueReader =
        catalogueReader ?? new ExcelConfigCatalogueReader(application);
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public TypeOptionsMaterialiseOutcome Materialise()
    {
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return TypeOptionsMaterialiseOutcome.Refused(
                activeProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? TypeOptionsRefusalReason.NoActiveWorkbook
                    : TypeOptionsRefusalReason.TargetProtected);
        }

        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NoActiveWorkbook);
        }

        ConfigReadOutcome catalogue = _catalogueReader.Read();
        if (!catalogue.Succeeded)
        {
            return TypeOptionsMaterialiseOutcome.Refused(
                catalogue.Refusal == ConfigReadRefusalReason.ConfigSheetMissing
                    ? TypeOptionsRefusalReason.ConfigSheetMissing
                    : catalogue.Refusal == ConfigReadRefusalReason.TableMissing
                        ? TypeOptionsRefusalReason.TableMissing
                        : TypeOptionsRefusalReason.CatalogueHashMismatch);
        }

        Excel.Sheets sheets = workbook.Sheets;
        Excel.Worksheet? config = FindSheet(sheets, GanttWorkbookContract.ConfigSheetName);
        Excel.Worksheet? gantt = FindTableSheet(sheets, GanttTableSchema.TableName);
        if (config is null || gantt is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(
                config is null
                    ? TypeOptionsRefusalReason.ConfigSheetMissing
                    : TypeOptionsRefusalReason.TableMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(config);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return TypeOptionsMaterialiseOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? TypeOptionsRefusalReason.NoActiveWorkbook
                    : TypeOptionsRefusalReason.TargetProtected);
        }

        Excel.ListObject? types = FindTable(config, GanttCatalogues.TypesTableName);
        Excel.ListObject? data = FindTable(gantt, GanttTableSchema.TableName);
        if (types is null || data is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.TableMissing);
        }

        Excel.ListColumn? typeColumn = FindColumn(data, "Type");
        Excel.ListColumn? displayNameColumn = FindColumn(types, "DisplayName");
        if (typeColumn is null || displayNameColumn is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.TableMissing);
        }

        Excel.Range? optionsRange = displayNameColumn.DataBodyRange;
        if (optionsRange is null)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch);
        }

        Excel.Names names = workbook.Names;
        Excel.Name? existingName = FindName(names, GanttWorkbookContract.TypeOptionsDefinedName);
        var expectedRefersTo = BuildRefersTo(config.Name, optionsRange);
        if (existingName is not null && !RefersToMatches(existingName.RefersTo, expectedRefersTo))
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NameTargetInvalid);
        }

        SetName(names, GanttWorkbookContract.TypeOptionsDefinedName, expectedRefersTo);
        Excel.Range? typeRange = typeColumn.DataBodyRange;
        if (typeRange is not null)
        {
            Excel.Validation validation = typeRange.Validation;
            validation.Delete();
            validation.Add(
                Excel.XlDVType.xlValidateList,
                Excel.XlDVAlertStyle.xlValidAlertStop,
                Type.Missing,
                $"={GanttWorkbookContract.TypeOptionsDefinedName}",
                Type.Missing);
            validation.InCellDropdown = true;
            validation.IgnoreBlank = false;
            validation.ShowError = true;
        }

        return TypeOptionsMaterialiseOutcome.Ok();
    }

    private static string BuildRefersTo(string sheetName, Excel.Range range)
    {
        var address = range.Address[true, true]
            ?? throw new InvalidOperationException("TypeOptions range has no address.");
        return $"='{sheetName.Replace("'", "''", StringComparison.Ordinal)}'!{address}";
    }

    /// <summary>
    /// Returns whether an existing name designates the same sheet and address this
    /// materialiser would write.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Excel normalises a stored name's <c>RefersTo</c>: it drops the quotes around
    /// a sheet name that does not need them, so a name this adapter wrote as
    /// <c>='_GanttCreatorConfig'!$B$2:$B$17</c> is read back as
    /// <c>=_GanttCreatorConfig!$B$2:$B$17</c>. Comparing the raw strings with
    /// <see cref="StringComparison.Ordinal"/> therefore rejected the adapter's own
    /// correct value on every call and refused every Add Row.
    /// </para>
    /// <para>
    /// The sheet name is compared case-insensitively (Excel treats sheet names that
    /// way) after unquoting and unescaping; the range address is compared Ordinally,
    /// so a name pointing at a different range is still refused.
    /// </para>
    /// </remarks>
    /// <param name="existingRefersTo">The stored name's target as Excel reports it.</param>
    /// <param name="expectedRefersTo">The target this materialiser would write.</param>
    /// <returns>Whether both designate the same sheet and address.</returns>
    internal static bool RefersToMatches(string? existingRefersTo, string expectedRefersTo)
    {
        if (existingRefersTo is null)
        {
            return false;
        }

        NameTarget existing = ParseRefersTo(existingRefersTo);
        NameTarget expected = ParseRefersTo(expectedRefersTo);

        // Both targets must parse; an unparsable one is a mismatch, not a match.
        return existing.IsParsed
            && expected.IsParsed
            && string.Equals(existing.Sheet, expected.Sheet, StringComparison.OrdinalIgnoreCase)
            && string.Equals(existing.Address, expected.Address, StringComparison.Ordinal);
    }

    /// <summary>A <c>=Sheet!Address</c> name target split into its parts.</summary>
    /// <param name="Sheet">The unquoted, unescaped sheet name.</param>
    /// <param name="Address">The range address.</param>
    private readonly record struct NameTarget(string Sheet, string Address)
    {
        /// <summary>Gets whether the target carried both a sheet and an address.</summary>
        public bool IsParsed => Sheet.Length > 0 && Address.Length > 0;
    }

    /// <summary>
    /// Splits a <c>=Sheet!Address</c> name target into its unquoted, unescaped
    /// sheet name and its range address.
    /// </summary>
    /// <param name="refersTo">The name target.</param>
    /// <returns>The parsed parts, both empty when the target cannot be split.</returns>
    private static NameTarget ParseRefersTo(string refersTo)
    {
        var text = refersTo.TrimStart('=').Trim();
        var separator = text.LastIndexOf('!');
        if (separator <= 0 || separator == text.Length - 1)
        {
            return new NameTarget(string.Empty, string.Empty);
        }

        var sheetPart = text[..separator].Trim();
        var address = text[(separator + 1)..].Trim();
        var sheet = sheetPart.Length >= 2 && sheetPart[0] == '\'' && sheetPart[^1] == '\''
            ? sheetPart[1..^1].Replace("''", "'", StringComparison.Ordinal)
            : sheetPart;

        return new NameTarget(sheet, address);
    }

    private static Excel.Name? FindName(Excel.Names names, string nameText)
    {
        var count = names.Count;
        for (var index = 1; index <= count; index++)
        {
            Excel.Name candidate = names.Item(index);
            if (string.Equals(candidate.Name, nameText, StringComparison.OrdinalIgnoreCase))
            {
                return candidate;
            }
        }

        return null;
    }

    private static void SetName(Excel.Names names, string nameText, string refersTo)
    {
        Excel.Name? existing = FindName(names, nameText);
        if (existing is null)
        {
            _ = names.Add(nameText, refersTo);
        }
        else
        {
            existing.RefersTo = refersTo;
        }
    }

    private static Excel.Worksheet? FindSheet(Excel.Sheets sheets, string name)
    {
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            if (sheets[index] is Excel.Worksheet sheet
                && string.Equals(sheet.Name, name, StringComparison.Ordinal))
            {
                return sheet;
            }
        }

        return null;
    }

    private static Excel.Worksheet? FindTableSheet(Excel.Sheets sheets, string tableName)
    {
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            if (sheets[index] is Excel.Worksheet sheet && FindTable(sheet, tableName) is not null)
            {
                return sheet;
            }
        }

        return null;
    }

    private static Excel.ListObject? FindTable(Excel.Worksheet sheet, string name)
    {
        Excel.ListObjects objects = sheet.ListObjects;
        var count = objects.Count;
        for (var index = 1; index <= count; index++)
        {
            Excel.ListObject table = objects[index];
            if (string.Equals(table.Name, name, StringComparison.Ordinal))
            {
                return table;
            }
        }

        return null;
    }

    private static Excel.ListColumn? FindColumn(Excel.ListObject table, string name)
    {
        Excel.ListColumns columns = table.ListColumns;
        var count = columns.Count;
        for (var index = 1; index <= count; index++)
        {
            Excel.ListColumn column = columns[index];
            if (string.Equals(column.Name, name, StringComparison.Ordinal))
            {
                return column;
            }
        }

        return null;
    }
}
