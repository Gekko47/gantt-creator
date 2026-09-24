namespace GanttCreator.Office;

/// <summary>
/// Normalizes rectangular Excel <c>Value2</c> SAFEARRAY payloads into
/// zero-based row arrays for Office adapters. Excel commonly returns
/// one-based lower bounds, while test and non-Excel hosts may return zero-based
/// arrays; both shapes are equivalent after normalization.
/// </summary>
internal static class ExcelValue2Matrix
{
    /// <summary>
    /// Converts a two-dimensional <c>Value2</c> payload into zero-based rows.
    /// Unsupported shapes and non-two-dimensional arrays produce no rows.
    /// </summary>
    /// <param name="raw">The raw <c>Value2</c> payload.</param>
    /// <returns>Zero-based row arrays.</returns>
    public static List<object?[]> ReadRows(object? raw)
    {
        if (raw is not Array matrix || matrix.Rank != 2)
        {
            return [];
        }

        var rowLower = matrix.GetLowerBound(0);
        var rowUpper = matrix.GetUpperBound(0);
        var columnLower = matrix.GetLowerBound(1);
        var columnCount = matrix.GetLength(1);
        var result = new List<object?[]>(rowUpper >= rowLower ? rowUpper - rowLower + 1 : 0);
        for (var row = rowLower; row <= rowUpper; row++)
        {
            var cells = new object?[columnCount];
            for (var column = 0; column < columnCount; column++)
            {
                cells[column] = matrix.GetValue(row, columnLower + column);
            }

            result.Add(cells);
        }

        return result;
    }
}
