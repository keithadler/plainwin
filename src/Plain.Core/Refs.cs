namespace Plain.Core;

/// <summary>A rectangle of cells a formula refers to, on a named sheet or on the formula's own sheet.</summary>
public readonly record struct RefRange(string? Sheet, int ColumnMin, int RowMin, int ColumnMax, int RowMax)
{
    public bool Contains(CellRef cell) =>
        cell.Column >= ColumnMin && cell.Column <= ColumnMax && cell.Row >= RowMin && cell.Row <= RowMax;

    public override string ToString() =>
        (Sheet is null ? "" : Sheet + "!") +
        new CellRef(ColumnMin, RowMin) +
        (ColumnMin == ColumnMax && RowMin == RowMax ? "" : ":" + new CellRef(ColumnMax, RowMax));
}

/// <summary>
/// Which cells a formula reads. Plain needs this to know which cached totals an edit made untrue, and only those:
/// clearing every formula on a sheet because one label changed would empty a spreadsheet of its numbers.
///
/// This reads references, it does not understand formulas. Text in quotes is skipped, a name followed by "(" is a
/// function and not a reference (so LOG10 is not column LOG row 10), and a whole-column reference like A:A covers
/// every row. Anything it cannot make sense of it simply does not report, and the caller falls back to caution.
/// </summary>
public static class Refs
{
    private const int MaxColumn = 16384, MaxRow = 1048576;

    public static IReadOnlyList<RefRange> Parse(string formula)
    {
        var found = new List<RefRange>();
        if (string.IsNullOrEmpty(formula)) return found;
        if (formula[0] == '=') formula = formula[1..];

        int i = 0;
        while (i < formula.Length)
        {
            char ch = formula[i];

            if (ch == '"') { i = SkipQuoted(formula, i, '"'); continue; }

            if (!(char.IsAsciiLetterOrDigit(ch) || ch is '$' or '\'' or '_'))
            {
                i++;
                continue;
            }

            int start = i;
            string? sheet = null;

            // A sheet name, either bare (Sheet2!) or quoted ('My Sheet'!).
            if (ch == '\'')
            {
                int end = SkipQuoted(formula, i, '\'');
                if (end < formula.Length && end - 1 > i && formula[end] == '!')
                {
                    sheet = formula[(i + 1)..(end - 1)].Replace("''", "'");
                    i = end + 1;
                }
                else { i = end; continue; }
            }
            else if (char.IsAsciiLetter(ch) || ch == '_')
            {
                int j = i;
                while (j < formula.Length && (char.IsAsciiLetterOrDigit(formula[j]) || formula[j] is '_' or '.')) j++;
                if (j < formula.Length && formula[j] == '!') { sheet = formula[i..j]; i = j + 1; }
            }

            if (TryRange(formula, ref i, sheet, out var range)) { found.Add(range); continue; }

            // Not a reference. Step over the whole name, never back into the middle of it, or "MyRange" would
            // be rescanned as the column NGE and a rejected row would come back as a smaller one.
            i = SkipName(formula, sheet is null ? start : i);
            if (i <= start) i = start + 1;
        }
        return found;
    }

    /// <summary>Move past a run of the characters a name can be made of, so a rejected name is never rescanned.</summary>
    private static int SkipName(string text, int at)
    {
        int i = at;
        while (i < text.Length && text[i] == '$') i++;
        while (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '_' or '.' or '$')) i++;
        return i > at ? i : at + 1;
    }

    private static int SkipQuoted(string text, int at, char quote)
    {
        int i = at + 1;
        while (i < text.Length)
        {
            if (text[i] == quote)
            {
                if (i + 1 < text.Length && text[i + 1] == quote) { i += 2; continue; }   // an escaped quote
                return i + 1;
            }
            i++;
        }
        return text.Length;
    }

    /// <summary>What one corner of a reference turned out to be.</summary>
    private enum Corner { None, Cell, WholeColumn, WholeRow }

    private static bool TryRange(string text, ref int i, string? sheet, out RefRange range)
    {
        range = default;
        int save = i;

        var first = ReadCorner(text, ref i, out int c1, out int r1);
        if (first == Corner.None) { i = save; return false; }

        // A name immediately followed by "(" is a function call, never a reference.
        if (i < text.Length && text[i] == '(') { i = save; return false; }

        int c2 = c1, r2 = r1;
        var second = Corner.None;
        if (i < text.Length && text[i] == ':')
        {
            int probe = i + 1;
            second = ReadCorner(text, ref probe, out int c3, out int r3);
            if (second != Corner.None) { i = probe; c2 = c3; r2 = r3; }
        }

        if (second == Corner.None)
        {
            // On its own, only a proper cell counts. A lone "A" is a name and a lone "42" is a number.
            if (first != Corner.Cell) { i = save; return false; }
        }
        else if (first != second) { i = save; return false; }
        else if (first == Corner.WholeColumn) { r1 = 1; r2 = MaxRow; }
        else if (first == Corner.WholeRow) { c1 = 1; c2 = MaxColumn; }

        range = new RefRange(sheet, Math.Min(c1, c2), Math.Min(r1, r2), Math.Max(c1, c2), Math.Max(r1, r2));
        return true;
    }

    /// <summary>One corner: letters then digits, either part optionally absolute with a dollar.</summary>
    private static Corner ReadCorner(string text, ref int i, out int column, out int row)
    {
        column = 1; row = 1;
        int save = i, letters = 0, digits = 0, col = 0, rw = 0;

        if (i < text.Length && text[i] == '$') i++;
        while (i < text.Length && char.IsAsciiLetter(text[i]) && letters < 4)
        {
            col = col * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            i++; letters++;
        }
        if (i < text.Length && char.IsAsciiLetter(text[i])) { i = save; return Corner.None; }   // too long for a column

        if (i < text.Length && text[i] == '$') i++;
        while (i < text.Length && char.IsAsciiDigit(text[i]))
        {
            rw = rw * 10 + (text[i] - '0');
            i++; digits++;
            if (rw > MaxRow) { i = save; return Corner.None; }
        }

        // A letter, digit or underscore straight after means this was part of a longer name.
        if (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '_' or '.')) { i = save; return Corner.None; }
        if (letters > 0 && col > MaxColumn) { i = save; return Corner.None; }

        if (letters > 0 && digits > 0) { column = col; row = rw; return Corner.Cell; }
        if (letters > 0) { column = col; return Corner.WholeColumn; }
        if (digits > 0) { row = rw; return Corner.WholeRow; }

        i = save;
        return Corner.None;
    }
}
