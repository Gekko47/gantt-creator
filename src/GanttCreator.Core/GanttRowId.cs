using System.Globalization;

namespace GanttCreator.Core;

/// <summary>
/// A stable identifier for one row of the visible <c>tblGanttData</c> table
/// (the <c>Id</c> column, and the same neutral ID space for <c>LaneId</c> and
/// <c>ParentId</c> references). Generated once, never derived from row number:
/// reordering, inserting, moving, or deleting rows never rewrites a surviving
/// row's identifier. Immutable.
/// </summary>
/// <remarks>
/// <para>
/// The workbook text is <c>G-</c> followed by 32 lowercase hexadecimal digits
/// (a <see cref="Guid"/> in <c>"N"</c> form). The leading letter keeps Excel
/// from coercing the cell to a number or a formula. The exact text is part of
/// the workbook schema; changing the format is a schema migration (see
/// <see cref="GanttSchemaVersion.CurrentSchemaVersion"/>).
/// </para>
/// <para>
/// Parsing mirrors <see cref="EntityTypeCatalog"/>: <c>TryParse</c> matches
/// Ordinal on trimmed input with no case-folding and never throws.
/// Cross-row policy (duplicate or missing IDs) belongs to later validation
/// work, not to this value object.
/// </para>
/// </remarks>
public sealed class GanttRowId : IEquatable<GanttRowId>
{
    /// <summary>The constant prefix that opens every identifier.</summary>
    public const string Prefix = "G-";

    /// <summary>The exact length of every identifier's workbook text.</summary>
    public const int TextLength = 34;

    /// <summary>Initialises an identifier from already-validated text.</summary>
    /// <param name="value">The exact workbook text; validated by the caller.</param>
    private GanttRowId(string value)
    {
        Value = value;
    }

    /// <summary>Gets the exact workbook text of the identifier.</summary>
    public string Value { get; }

    /// <summary>
    /// Generates a new unique row identifier.
    /// </summary>
    /// <returns>A new identifier that is distinct from every previously generated one.</returns>
    public static GanttRowId New()
    {
        var token = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture);
        return new GanttRowId(Prefix + token);
    }

    /// <summary>
    /// Parses workbook text into an identifier.
    /// </summary>
    /// <param name="text">The workbook text to parse.</param>
    /// <returns>The parsed identifier.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="text"/> is null.</exception>
    /// <exception cref="FormatException">Thrown when <paramref name="text"/> is not a well-formed identifier.</exception>
    public static GanttRowId Parse(string? text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return TryParse(text, out GanttRowId? id) && id is not null
            ? id
            : throw new FormatException($"'{text}' is not a valid Gantt row identifier.");
    }

    /// <summary>
    /// Tries to parse workbook text into an identifier. Matching is Ordinal on
    /// trimmed input: no case-folding, no whitespace rewriting, no guessing.
    /// Never throws.
    /// </summary>
    /// <param name="text">The workbook text to parse.</param>
    /// <param name="id">The parsed identifier when the method returns <see langword="true"/>.</param>
    /// <returns><see langword="true"/> when the text is a well-formed identifier.</returns>
    public static bool TryParse(string? text, out GanttRowId? id)
    {
        id = null;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var trimmed = text.Trim();
        if (trimmed.Length != TextLength)
        {
            return false;
        }

        if (!trimmed.StartsWith(Prefix, StringComparison.Ordinal))
        {
            return false;
        }

        for (var i = Prefix.Length; i < trimmed.Length; i++)
        {
            var c = trimmed[i];
            if (c is not (>= '0' and <= '9') and not (>= 'a' and <= 'f'))
            {
                return false;
            }
        }

        id = new GanttRowId(trimmed);
        return true;
    }

    /// <inheritdoc cref="IEquatable{T}.Equals(T)" />
    public bool Equals(GanttRowId? other) =>
        other is not null && string.Equals(Value, other.Value, StringComparison.Ordinal);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is GanttRowId other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(Value);

    /// <summary>Equality operator; compares by identifier text.</summary>
    /// <param name="left">The left operand.</param>
    /// <param name="right">The right operand.</param>
    /// <returns><see langword="true"/> when both hold the same identifier text.</returns>
    public static bool operator ==(GanttRowId? left, GanttRowId? right) =>
        left is null ? right is null : left.Equals(right);

    /// <summary>Inequality operator; compares by identifier text.</summary>
    /// <param name="left">The left operand.</param>
    /// <param name="right">The right operand.</param>
    /// <returns><see langword="true"/> when the identifier texts differ.</returns>
    public static bool operator !=(GanttRowId? left, GanttRowId? right) => !(left == right);

    /// <summary>Returns the exact workbook text of the identifier.</summary>
    /// <returns>The identifier text.</returns>
    public override string ToString() => Value;
}
