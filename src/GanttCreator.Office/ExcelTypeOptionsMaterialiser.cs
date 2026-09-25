using System.Runtime.InteropServices;
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
        // ADR-0008 D4: the shared protection guard is intentionally the first
        // operation in this mutating entry point, before any workbook access.
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        return Apply(activeProtection, allowStaleNameTargetReplacement: false, skipWhenCurrent: false);
    }

    /// <inheritdoc />
    public TypeOptionsMaterialiseOutcome MaterialiseForRepair()
    {
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        return Apply(activeProtection, allowStaleNameTargetReplacement: true, skipWhenCurrent: false);
    }

    /// <summary>
    /// Returns success without writing when the stored TypeOptions name and the
    /// Type-column validation already match what this adapter would write, so a
    /// scaffold row does not re-delete and re-add the same validation on every
    /// Add Row.
    /// </summary>
    /// <remarks>
    /// Currency is only ever a reason to skip work. Anything that is missing,
    /// stale, or unreadable falls through to the ordinary
    /// <see cref="Materialise"/>, including the
    /// <see cref="TypeOptionsRefusalReason.NameTargetInvalid"/> refusal, so this
    /// never repairs a name it did not write.
    /// </remarks>
    /// <returns>The typed ensure result.</returns>
    public TypeOptionsMaterialiseOutcome EnsureCurrent()
    {
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        return Apply(activeProtection, allowStaleNameTargetReplacement: false, skipWhenCurrent: true);
    }

    private TypeOptionsMaterialiseOutcome Apply(
        ProtectionGuardOutcome activeProtection,
        bool allowStaleNameTargetReplacement,
        bool skipWhenCurrent)
    {
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
        var expectedRefersTo = GetRefersTo(config.Name, optionsRange);
        bool nameIsCurrent =
            existingName is not null && RefersToMatches(existingName.RefersTo, expectedRefersTo);

        if (existingName is not null && !nameIsCurrent && !allowStaleNameTargetReplacement)
        {
            return TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NameTargetInvalid);
        }

        Excel.Range? typeRange = typeColumn.DataBodyRange;
        if (skipWhenCurrent && nameIsCurrent && ValidationIsCurrent(typeRange))
        {
            return TypeOptionsMaterialiseOutcome.Ok();
        }

        SetName(names, GanttWorkbookContract.TypeOptionsDefinedName, expectedRefersTo);
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

    /// <summary>
    /// Builds the <c>=Sheet!Address</c> target for a catalogue range.
    /// </summary>
    /// <remarks>
    /// Virtual so contract tests can substitute it: <c>Range.Address</c> is a
    /// COM parameterised property, which an expression tree cannot contain
    /// (CS0855), so it cannot be mocked directly.
    /// </remarks>
    /// <param name="sheetName">The configuration worksheet name.</param>
    /// <param name="range">The catalogue range to reference.</param>
    /// <returns>The name target this adapter would write.</returns>
    internal virtual string GetRefersTo(string sheetName, Excel.Range range) =>
        BuildRefersTo(sheetName, range);

    /// <summary>
    /// Returns whether the Type body already carries this adapter's list
    /// validation against the TypeOptions name.
    /// </summary>
    /// <remarks>
    /// A read failure is reported as "not current" so the caller re-applies the
    /// validation rather than assuming a state it could not confirm. The
    /// re-application is idempotent, so the fallback is safe.
    /// </remarks>
    /// <param name="typeRange">The Type column's data body, or null when empty.</param>
    /// <returns><see langword="true"/> only when the stored validation is current.</returns>
    private static bool ValidationIsCurrent(Excel.Range? typeRange)
    {
        if (typeRange is null)
        {
            return false;
        }

        try
        {
            Excel.Validation validation = typeRange.Validation;
            // The PIA surfaces Validation.Type as the raw int; the integration
            // test asserts against the same cast.
            if ((Excel.XlDVType)validation.Type != Excel.XlDVType.xlValidateList)
            {
                return false;
            }

            var formula = validation.Formula1;
            return string.Equals(
                formula?.Trim(),
                $"={GanttWorkbookContract.TypeOptionsDefinedName}",
                StringComparison.OrdinalIgnoreCase);
        }
        catch (COMException)
        {
            return false;
        }
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
