using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Read-only adapter that observes every worksheet's name, visibility, and
/// contract-table count and hands them to the pure
/// <see cref="ConfigSheetLocator"/>. It never mutates the workbook.
/// </summary>
/// <remarks>
/// The identity decision lives in Core so every state has a positive test; this
/// adapter owns only the COM enumeration. A table name is compared against
/// <see cref="GanttCatalogues.ContractTableNames"/> rather than a literal.
/// </remarks>
public class ExcelConfigSheetInspector
{
    /// <summary>Locates the configuration worksheet in the active workbook's sheets.</summary>
    /// <param name="sheets">The workbook's sheets collection.</param>
    /// <returns>The location and its identity state.</returns>
    public ConfigSheetLocation Locate(Excel.Sheets sheets)
    {
        ArgumentNullException.ThrowIfNull(sheets);

        int count = GetSheetCount(sheets);
        List<ConfigSheetObservation> observations = new(count);
        for (int index = 1; index <= count; index++)
        {
            if (GetSheetAt(sheets, index) is not Excel.Worksheet worksheet)
            {
                continue;
            }

            observations.Add(
                new ConfigSheetObservation(
                    GetSheetName(worksheet),
                    GetSheetVisibility(worksheet) == Excel.XlSheetVisibility.xlSheetVeryHidden,
                    CountContractTables(worksheet)));
        }

        return ConfigSheetLocator.Locate(observations);
    }

    private int CountContractTables(Excel.Worksheet worksheet)
    {
        Excel.ListObjects objects = worksheet.ListObjects;
        int count = GetTableCount(objects);
        int contractTables = 0;
        for (int index = 1; index <= count; index++)
        {
            string name = GetTableName(GetTableAt(objects, index));
            foreach (string contractName in GanttCatalogues.ContractTableNames)
            {
                if (string.Equals(name, contractName, StringComparison.OrdinalIgnoreCase))
                {
                    contractTables++;
                    break;
                }
            }
        }

        return contractTables;
    }

    /// <summary>Reads the sheet count. Test seam over <c>Sheets.Count</c>.</summary>
    /// <param name="sheets">The sheets collection.</param>
    /// <returns>The number of sheets.</returns>
    internal virtual int GetSheetCount(Excel.Sheets sheets) => sheets.Count;

    /// <summary>Reads one sheet by index. Test seam over <c>Sheets.Item</c>.</summary>
    /// <param name="sheets">The sheets collection.</param>
    /// <param name="index">The 1-based index.</param>
    /// <returns>The sheet object.</returns>
    internal virtual object GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    /// <summary>Reads a worksheet name. Test seam over <c>Worksheet.Name</c>.</summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <returns>The live sheet name.</returns>
    internal virtual string GetSheetName(Excel.Worksheet worksheet) => worksheet.Name;

    /// <summary>Reads a worksheet visibility. Test seam over <c>Worksheet.Visible</c>.</summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <returns>The visibility state.</returns>
    internal virtual Excel.XlSheetVisibility GetSheetVisibility(Excel.Worksheet worksheet) => worksheet.Visible;

    /// <summary>Reads a table count. Test seam over <c>ListObjects.Count</c>.</summary>
    /// <param name="objects">The sheet's tables.</param>
    /// <returns>The number of tables.</returns>
    internal virtual int GetTableCount(Excel.ListObjects objects) => objects.Count;

    /// <summary>Reads one table. Test seam over <c>ListObjects.Item</c>.</summary>
    /// <param name="objects">The sheet's tables.</param>
    /// <param name="index">The 1-based index.</param>
    /// <returns>The table.</returns>
    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) => objects[index];

    /// <summary>Reads a table name. Test seam over <c>ListObject.Name</c>.</summary>
    /// <param name="table">The table.</param>
    /// <returns>The live table name.</returns>
    internal virtual string GetTableName(Excel.ListObject table) => table.Name;
}
