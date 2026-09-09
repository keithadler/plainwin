namespace Plain.Core;

/// <summary>
/// Sorting rows of a sheet.
///
/// Sorting is the one common spreadsheet job that can quietly ruin a file, because moving a row moves what every
/// formula in the workbook was pointing at. Plain will not guess its way through that. It sorts a block of plain
/// values and refuses, with a reason you can read, whenever a formula is involved: one inside the block, or one
/// anywhere else in the workbook that points into it. Refusing is the honest answer, and it is the same principle
/// as the rest of the app: it would rather do nothing than hand you back something subtly wrong.
/// </summary>
public static class Sort
{
    public abstract record Result;

    /// <summary>Sorted, and how many rows moved from where they were.</summary>
    public sealed record Sorted(int RowsMoved) : Result;

    /// <summary>Not sorted, and why, in a sentence meant to be shown to a person.</summary>
    public sealed record Refused(string Reason) : Result;

    /// <summary>
    /// Sort the rows from <paramref name="top"/> to <paramref name="bottom"/> by what is in
    /// <paramref name="keyColumn"/>, moving only the columns from <paramref name="left"/> to
    /// <paramref name="right"/> so a sort of one block never disturbs what sits beside it.
    /// </summary>
    public static Result Rows(Workbook book, Sheet sheet, int top, int bottom, int left, int right,
                              int keyColumn, bool ascending)
    {
        if (bottom <= top) return new Refused("Select more than one row to sort.");
        if (keyColumn < left || keyColumn > right) return new Refused("The column to sort by has to be inside the selection.");

        // A formula inside the block would be carried to a new row with its references unchanged, which is a
        // different sum than the one that was there. Refuse rather than move it.
        for (int r = top; r <= bottom; r++)
            for (int c = left; c <= right; c++)
                if (sheet.Read(new CellRef(c, r)).Formula is not null)
                    return new Refused($"There is a formula at {new CellRef(c, r)}. Plain does not sort formulas, "
                                     + "because moving one changes what it adds up.");

        // A formula elsewhere pointing into the block is the same problem seen from the other side, with one
        // important exception. A total underneath a table reads every row of it, and sorting does not change which
        // values are there, only their order, so the total is the same afterwards. That is the ordinary case and
        // refusing it would make sorting useless. A reference that covers only some of the rows is a different
        // matter: after the sort it would be reading different values, so that is still refused.
        foreach (var other in book.Sheets)
            foreach (var (at, formula) in other.Formulas())
                foreach (var token in Refs.Scan(formula))
                {
                    if (!Overlaps(token, sheet, other, top, bottom, left, right)) continue;

                    bool coversEveryRow = token.Range.RowMin <= top && token.Range.RowMax >= bottom;
                    if (coversEveryRow && !OrderMatters(formula)) continue;

                    return new Refused(coversEveryRow
                        ? $"The formula in {other.Name}!{at} depends on the order of these rows, so Plain has left "
                        + "them alone rather than change what it answers."
                        : $"The formula in {other.Name}!{at} reads some of these rows but not all of them. Sorting "
                        + "would change what it adds up, so Plain has left them alone.");
                }

        // Read the block out, sort it, and write it back. Everything is read before anything is written, so a
        // failure part way through cannot leave half a sorted sheet.
        var rows = new List<(string Key, List<Cell> Cells)>();
        for (int r = top; r <= bottom; r++)
        {
            var cells = new List<Cell>();
            for (int c = left; c <= right; c++) cells.Add(sheet.Read(new CellRef(c, r)));
            rows.Add((Value(cells[keyColumn - left]), cells));
        }

        var order = Enumerable.Range(0, rows.Count).ToList();
        order.Sort((a, b) =>
        {
            string ka = rows[a].Key, kb = rows[b].Key;
            // Blank rows go to the bottom whichever way the sort runs. Turning the order round should not float
            // the empty rows to the top, because nobody sorts in order to look at the gaps.
            bool ea = ka.Length == 0, eb = kb.Length == 0;
            if (ea || eb) return ea && eb ? a.CompareTo(b) : ea ? 1 : -1;

            int by = Compare(ka, kb);
            // A stable sort: rows that compare equal keep the order they were in, so sorting twice by two columns
            // does what people expect it to.
            return by != 0 ? (ascending ? by : -by) : a.CompareTo(b);
        });

        int moved = 0;
        for (int i = 0; i < order.Count; i++)
        {
            if (order[i] != i) moved++;
            var source = rows[order[i]].Cells;
            for (int c = left; c <= right; c++)
            {
                var cell = source[c - left];
                var target = new CellRef(c, top + i);
                sheet.Set(target, Value(cell));
            }
        }
        return new Sorted(moved);
    }

    /// <summary>
    /// What to put back, and what to sort on. A number's Display is how it is shown, thousands separators and all,
    /// so writing that back would turn 24000 into the text "24,000". The underlying value is what moves.
    /// </summary>
    private static string Value(Cell cell) => cell.Kind switch
    {
        CellKind.Empty => "",
        CellKind.Number or CellKind.Boolean => cell.Raw,
        _ => cell.Display,
    };

    /// <summary>Numbers before text, numbers by value, text without minding case. Blanks are dealt with above.</summary>
    private static int Compare(string a, string b)
    {
        bool na = double.TryParse(a, System.Globalization.NumberStyles.Any,
                                  System.Globalization.CultureInfo.InvariantCulture, out var da);
        bool nb = double.TryParse(b, System.Globalization.NumberStyles.Any,
                                  System.Globalization.CultureInfo.InvariantCulture, out var db);
        if (na && nb) return da.CompareTo(db);
        if (na != nb) return na ? -1 : 1;
        return string.Compare(a, b, StringComparison.CurrentCultureIgnoreCase);
    }

    /// <summary>
    /// Functions whose answer depends on where a value sits rather than only on which values are there. A total is
    /// the same however the rows are shuffled; a lookup is not. When one of these reads the block, Plain refuses
    /// even though every row is covered.
    /// </summary>
    private static readonly string[] OrderSensitive =
        { "INDEX", "MATCH", "VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP", "OFFSET", "INDIRECT", "ROW", "COLUMN" };

    private static bool OrderMatters(string formula)
    {
        foreach (var name in OrderSensitive)
        {
            int i = formula.IndexOf(name, StringComparison.OrdinalIgnoreCase);
            while (i >= 0)
            {
                // Only when it reads as a call, so a cell called INDEXED or a word in text is not mistaken for one.
                bool startsWord = i == 0 || !char.IsLetterOrDigit(formula[i - 1]);
                int after = i + name.Length;
                while (after < formula.Length && formula[after] == ' ') after++;
                if (startsWord && after < formula.Length && formula[after] == '(') return true;
                i = formula.IndexOf(name, i + 1, StringComparison.OrdinalIgnoreCase);
            }
        }
        return false;
    }

    /// <summary>Does this reference touch the block being sorted? A reference with no sheet means its own sheet.</summary>
    private static bool Overlaps(Refs.Token token, Sheet target, Sheet owner, int top, int bottom, int left, int right)
    {
        var named = token.Range.Sheet;
        bool sameSheet = named is null or "" ? ReferenceEquals(owner, target)
                                             : named.Equals(target.Name, StringComparison.OrdinalIgnoreCase);
        if (!sameSheet) return false;

        var r = token.Range;
        return r.RowMin <= bottom && r.RowMax >= top && r.ColumnMin <= right && r.ColumnMax >= left;
    }
}
