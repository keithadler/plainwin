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

    /// <summary>
    /// One reference exactly as it was written: where it sits in the formula, what it points at, and which parts were
    /// pinned with a dollar. Keeping the dollars matters because rewriting "$B$4" as "B5" would quietly unpin it.
    /// </summary>
    public readonly record struct Token(int Start, int Length, RefRange Range, Shape Kind,
                                        bool Column1Fixed, bool Row1Fixed, bool Column2Fixed, bool Row2Fixed, bool IsRange);

    public enum Shape { Cell, WholeColumn, WholeRow }

    public static IReadOnlyList<RefRange> Parse(string formula) => Scan(formula).Select(t => t.Range).ToList();

    /// <summary>
    /// Rewrite the references in a formula. The map is given each reference and returns the text to put in its place,
    /// or null to leave it exactly as it was written. Everything between references is copied through untouched, so a
    /// formula comes back the way its author wrote it apart from the references that had to move.
    /// </summary>
    public static string Rewrite(string formula, Func<Token, string?> map)
    {
        var tokens = Scan(formula);
        if (tokens.Count == 0) return formula;

        bool equals = formula.StartsWith('=');
        string body = equals ? formula[1..] : formula;

        var built = new System.Text.StringBuilder(body.Length + 16);
        int at = 0;
        foreach (var token in tokens)
        {
            var replacement = map(token);
            if (replacement is null) continue;
            built.Append(body, at, token.Start - at).Append(replacement);
            at = token.Start + token.Length;
        }
        if (at == 0) return formula;
        built.Append(body, at, body.Length - at);
        return (equals ? "=" : "") + built;
    }

    /// <summary>How a reference is written out again, with the dollars it came with.</summary>
    public static string Write(RefRange range, Shape kind, bool column1Fixed, bool row1Fixed,
                               bool column2Fixed, bool row2Fixed, bool isRange)
    {
        string Corner(int column, int row, bool columnFixed, bool rowFixed) => kind switch
        {
            Shape.WholeColumn => (columnFixed ? "$" : "") + CellRef.ColumnName(column),
            Shape.WholeRow => (rowFixed ? "$" : "") + row.ToString(System.Globalization.CultureInfo.InvariantCulture),
            _ => (columnFixed ? "$" : "") + CellRef.ColumnName(column) +
                 (rowFixed ? "$" : "") + row.ToString(System.Globalization.CultureInfo.InvariantCulture),
        };

        string sheet = range.Sheet is null ? "" : QuoteSheet(range.Sheet) + "!";
        string first = Corner(range.ColumnMin, range.RowMin, column1Fixed, row1Fixed);
        return isRange
            ? sheet + first + ":" + Corner(range.ColumnMax, range.RowMax, column2Fixed, row2Fixed)
            : sheet + first;
    }

    /// <summary>A sheet name needs quoting unless it is a plain word.</summary>
    public static string QuoteSheet(string name) =>
        name.All(c => char.IsAsciiLetterOrDigit(c) || c == '_') && name.Length > 0 && !char.IsAsciiDigit(name[0])
            ? name
            : "'" + name.Replace("'", "''") + "'";

    private static IReadOnlyList<Token> Scan(string formula)
    {
        var found = new List<Token>();
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

            if (TryRange(formula, ref i, sheet, out var token))
            {
                found.Add(token with { Start = start, Length = i - start });
                continue;
            }

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

    private static bool TryRange(string text, ref int i, string? sheet, out Token token)
    {
        token = default;
        int save = i;

        var first = ReadCorner(text, ref i, out int c1, out int r1, out bool cf1, out bool rf1);
        if (first == Corner.None) { i = save; return false; }

        // A name immediately followed by "(" is a function call, never a reference.
        if (i < text.Length && text[i] == '(') { i = save; return false; }

        int c2 = c1, r2 = r1;
        bool cf2 = cf1, rf2 = rf1;
        var second = Corner.None;
        if (i < text.Length && text[i] == ':')
        {
            int probe = i + 1;
            second = ReadCorner(text, ref probe, out int c3, out int r3, out bool cf3, out bool rf3);
            if (second != Corner.None) { i = probe; c2 = c3; r2 = r3; cf2 = cf3; rf2 = rf3; }
        }

        if (second == Corner.None)
        {
            // On its own, only a proper cell counts. A lone "A" is a name and a lone "42" is a number.
            if (first != Corner.Cell) { i = save; return false; }
        }
        else if (first != second) { i = save; return false; }
        else if (first == Corner.WholeColumn) { r1 = 1; r2 = MaxRow; }
        else if (first == Corner.WholeRow) { c1 = 1; c2 = MaxColumn; }

        bool swappedColumns = c1 > c2, swappedRows = r1 > r2;
        var range = new RefRange(sheet, Math.Min(c1, c2), Math.Min(r1, r2), Math.Max(c1, c2), Math.Max(r1, r2));
        token = new Token(0, 0, range,
            first switch { Corner.WholeColumn => Shape.WholeColumn, Corner.WholeRow => Shape.WholeRow, _ => Shape.Cell },
            swappedColumns ? cf2 : cf1, swappedRows ? rf2 : rf1,
            swappedColumns ? cf1 : cf2, swappedRows ? rf1 : rf2,
            second != Corner.None);
        return true;
    }

    /// <summary>One corner: letters then digits, either part optionally absolute with a dollar.</summary>
    private static Corner ReadCorner(string text, ref int i, out int column, out int row,
                                     out bool columnFixed, out bool rowFixed)
    {
        column = 1; row = 1;
        columnFixed = rowFixed = false;
        int save = i, letters = 0, digits = 0, col = 0, rw = 0;

        if (i < text.Length && text[i] == '$') { i++; columnFixed = true; }
        while (i < text.Length && char.IsAsciiLetter(text[i]) && letters < 4)
        {
            col = col * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            i++; letters++;
        }
        if (i < text.Length && char.IsAsciiLetter(text[i])) { i = save; return Corner.None; }   // too long for a column

        if (i < text.Length && text[i] == '$') { i++; rowFixed = true; }
        while (i < text.Length && char.IsAsciiDigit(text[i]))
        {
            rw = rw * 10 + (text[i] - '0');
            i++; digits++;
            if (rw > MaxRow) { i = save; return Corner.None; }
        }

        // A letter, digit or underscore straight after means this was part of a longer name.
        if (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '_' or '.')) { i = save; return Corner.None; }
        if (letters > 0 && col > MaxColumn) { i = save; return Corner.None; }
        if (letters == 0) columnFixed = false;
        if (digits == 0) rowFixed = false;

        if (letters > 0 && digits > 0) { column = col; row = rw; return Corner.Cell; }
        if (letters > 0) { column = col; return Corner.WholeColumn; }
        if (digits > 0) { row = rw; return Corner.WholeRow; }

        i = save;
        return Corner.None;
    }
}
