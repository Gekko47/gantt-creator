namespace GanttCreator.Core;

/// <summary>An immutable RGB or ARGB hexadecimal colour value.</summary>
public sealed class ColourHex : IEquatable<ColourHex>
{
    private const string _prefix = "#";

    private ColourHex(uint argb, bool hasExplicitAlpha)
    {
        ARGB = argb;
        HasExplicitAlpha = hasExplicitAlpha;
    }

    /// <summary>Gets the colour as an unsigned <c>AARRGGBB</c> value.</summary>
    public uint ARGB { get; }

    private bool HasExplicitAlpha { get; }

    /// <summary>Parses a colour in <c>#RRGGBB</c> or <c>#RRGGBBAA</c> form.</summary>
    /// <param name="text">The colour text to parse.</param>
    /// <returns>The parsed colour.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="text"/> is null.</exception>
    /// <exception cref="FormatException">Thrown when the text is not a supported uppercase colour.</exception>
    public static ColourHex Parse(string? text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return TryParse(text, out ColourHex? colour) && colour is not null
            ? colour
            : throw new FormatException($"'{text}' is not a valid colour.");
    }

    /// <summary>Tries to parse a colour without throwing.</summary>
    /// <param name="text">The colour text to parse. Surrounding whitespace is ignored.</param>
    /// <param name="colour">The parsed colour when successful.</param>
    /// <returns><see langword="true"/> when the text is a supported uppercase colour.</returns>
    public static bool TryParse(string? text, out ColourHex? colour)
    {
        colour = null;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var value = text.Trim();
        if ((value.Length != 7 && value.Length != 9) ||
            !value.StartsWith(_prefix, StringComparison.Ordinal) ||
            !HasUpperHexDigits(value, 1))
        {
            return false;
        }

        var red = ReadByte(value, 1);
        var green = ReadByte(value, 3);
        var blue = ReadByte(value, 5);
        var alpha = value.Length == 9 ? ReadByte(value, 7) : byte.MaxValue;
        colour = new ColourHex(
            ((uint)alpha << 24) | ((uint)red << 16) | ((uint)green << 8) | blue,
            value.Length == 9);
        return true;
    }

    /// <inheritdoc />
    public bool Equals(ColourHex? other) =>
        other is not null
        && ARGB == other.ARGB
        && HasExplicitAlpha == other.HasExplicitAlpha;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ColourHex other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(ARGB, HasExplicitAlpha);

    /// <summary>Equality operator comparing the numeric ARGB value and alpha representation.</summary>
    /// <param name="left">The first colour.</param>
    /// <param name="right">The second colour.</param>
    /// <returns><see langword="true"/> when both colours have the same ARGB value and alpha representation.</returns>
    public static bool operator ==(ColourHex? left, ColourHex? right) => left is null ? right is null : left.Equals(right);

    /// <summary>Inequality operator comparing the numeric ARGB value.</summary>
    /// <param name="left">The first colour.</param>
    /// <param name="right">The second colour.</param>
    /// <returns><see langword="true"/> when the colours differ.</returns>
    public static bool operator !=(ColourHex? left, ColourHex? right) => !(left == right);

    /// <summary>Returns the parsed colour form without changing its alpha representation.</summary>
    /// <returns>The canonical six- or eight-digit colour text.</returns>
    public override string ToString()
    {
        var red = (byte)(ARGB >> 16);
        var green = (byte)(ARGB >> 8);
        var blue = (byte)ARGB;
        if (!HasExplicitAlpha)
        {
            return $"{_prefix}{red:X2}{green:X2}{blue:X2}";
        }

        var alpha = (byte)(ARGB >> 24);
        return $"{_prefix}{red:X2}{green:X2}{blue:X2}{alpha:X2}";
    }

    private static bool HasUpperHexDigits(string value, int start)
    {
        for (var index = start; index < value.Length; index++)
        {
            var c = value[index];
            if (c is not ((>= '0' and <= '9') or (>= 'A' and <= 'F')))
            {
                return false;
            }
        }

        return true;
    }

    private static byte ReadByte(string value, int index) =>
        Convert.ToByte(value.Substring(index, 2), 16);
}
