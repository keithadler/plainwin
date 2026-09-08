namespace Plain.Core;

/// <summary>A cell reference such as "A1" or "AA10", and the arithmetic for moving around a grid.</summary>
public readonly record struct CellRef(int Column, int Row)
{
    /// <summary>Columns and rows are one-based, the way the reference on screen reads.</summary>
    public static CellRef Parse(string text)
    {
        if (!TryParse(text, out var r)) throw new FormatException($"\"{text}\" is not a cell reference.");
        return r;
    }

    public static bool TryParse(string text, out CellRef cell)
    {
        cell = default;
        if (string.IsNullOrEmpty(text)) return false;
        int i = 0, col = 0;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
        {
            col = col * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            i++;
            if (col > 16384) return false;
        }
        if (col == 0 || i == text.Length) return false;
        int row = 0;
        for (; i < text.Length; i++)
        {
            if (!char.IsAsciiDigit(text[i])) return false;
            row = row * 10 + (text[i] - '0');
            if (row > 1048576) return false;
        }
        if (row == 0) return false;
        cell = new CellRef(col, row);
        return true;
    }

    public static string ColumnName(int column)
    {
        Span<char> buf = stackalloc char[4];
        int i = 4;
        while (column > 0)
        {
            int rem = (column - 1) % 26;
            buf[--i] = (char)('A' + rem);
            column = (column - 1) / 26;
        }
        return new string(buf[i..]);
    }

    public override string ToString() => ColumnName(Column) + Row.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
