using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Shared, test-seamed table traversal for the Office adapters.
/// </summary>
/// <remarks>
/// <para>
/// <b>One authority for "find <c>tblGanttData</c>".</b> The adapters used to carry
/// their own copy of this walk, and the copies had already diverged. This type is
/// the single traversal; a derived adapter keeps only the thin virtual method its
/// contract tests already override, which delegates here.
/// </para>
/// <para>
/// <b>COM ownership (D1, approved 2026-11-24).</b> Releases every proxy it acquires
/// and does not return: a skipped <c>ListObject</c>, a skipped <c>ListObjects</c>
/// collection, and a non-owning <c>Worksheet</c>. It never releases the returned
/// table or sheet. <c>ReleaseComObject</c>, not <c>FinalReleaseComObject</c>,
/// because these proxies can be shared with another holder in the process.
/// </para>
/// <para>
/// The index seams exist for the same reason as the adapter seams: expression trees
/// cannot contain COM indexed properties, and the contract tests substitute them.
/// </para>
/// </remarks>
public abstract class ExcelTableAdapterBase
{
    /// <summary>Reads the sheet count. Test seam over <c>Sheets.Count</c>.</summary>
    internal virtual int GetSheetCount(Sheets sheets) => sheets.Count;

    /// <summary>Reads a sheet by one-based index. Test seam over <c>Sheets.Item</c>.</summary>
    internal virtual object? GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>Reads the table count. Test seam over <c>ListObjects.Count</c>.</summary>
    internal virtual int GetListObjectCount(ListObjects objects) => objects.Count;

    /// <summary>Reads a table by one-based index. Test seam over <c>ListObjects.Item</c>.</summary>
    internal virtual ListObject? GetTableAt(ListObjects objects, int index) => objects[index];

    /// <summary>
    /// Finds the worksheet and table carrying the Gantt data table.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="worksheet">The resolved worksheet, or <see langword="null"/>.</param>
    /// <param name="table">The resolved table, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when the table was found.</returns>
    internal virtual bool TryLocateTable(
        Sheets sheets,
        out Worksheet? worksheet,
        out ListObject? table)
    {
        ArgumentNullException.ThrowIfNull(sheets);

        worksheet = null;
        table = null;
        int sheetCount = GetSheetCount(sheets);
        for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            // A chart sheet carries no list objects; the cast fails and it is skipped.
            if (GetSheetAt(sheets, sheetIndex) is not Worksheet candidate)
            {
                continue;
            }

            bool found = false;
            ListObjects listObjects = candidate.ListObjects;
            try
            {
                int tableCount = GetListObjectCount(listObjects);
                for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    ListObject? located = GetTableAt(listObjects, tableIndex);
                    if (located is null)
                    {
                        continue;
                    }

                    if (GanttChartContract.IsDataTableName(located.Name))
                    {
                        worksheet = candidate;
                        table = located;
                        found = true;
                        break;
                    }

                    // Skipped: this adapter acquired it and does not return it.
                    ReleaseAcquired(located);
                }
            }
            finally
            {
                ReleaseAcquired(listObjects);
                if (!found)
                {
                    // Not the owning sheet, and the caller never sees it.
                    ReleaseAcquired(candidate);
                }
            }

            if (found)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Releases one acquired proxy, or ignores a null or non-COM value.
    /// </summary>
    /// <param name="proxy">The proxy this traversal acquired and does not return.</param>
    /// <remarks>
    /// Best-effort by design: a release must never change the read's outcome or
    /// escape into Excel. A non-COM value (a test mock) is ignored, and a refused
    /// release is swallowed because the read has already produced its answer.
    /// </remarks>
    internal virtual void ReleaseAcquired(object? proxy)
    {
        if (proxy is null || !Marshal.IsComObject(proxy))
        {
            return;
        }

        try
        {
            _ = Marshal.ReleaseComObject(proxy);
        }
        catch (Exception ex) when (ex is ArgumentException or COMException)
        {
            _ = ex;
        }
    }
}
