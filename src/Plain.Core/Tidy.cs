namespace Plain.Core;

/// <summary>
/// The two jobs everyone does to a list of data and nobody enjoys: taking out the rows that say the same thing
/// twice, and splitting one column into several.
///
/// Both change more than one cell at once, which is exactly when a tool needs to be careful, so both refuse in the
/// same circumstances as sorting does: a formula inside the block would be carried somewhere it does not mean, and
/// a formula elsewhere reading only part of the block would end up reading different values. Refusing with a
/// reason is better than a quietly wrong spreadsheet.
/// </summary>
public static class Tidy
{
    public abstract record Result;
    public sealed record Done(string What, int Rows) : Result;
    public sealed record Refused(string Reason) : Result;

    /// <summary>
    /// Take out rows whose cells across the block say the same as an earlier row's. The first of each set stays
    /// where it is, so the order of what is left is the order it was in.
    /// </summary>
    public static Result RemoveDuplicates(Workbook book, Sheet sheet, int top, int bottom, int left, int right)
    {
        if (bottom <= top) return new Refused("Select more than one row.");
        if (Guard(book, sheet, top, bottom, left, right) is Refused blocked) return blocked;

        var rows = new List<List<string>>();
        for (int r = top; r <= bottom; r++)
        {
            var values = new List<string>();
            for (int c = left; c <= right; c++) values.Add(Value(sheet.Read(new CellRef(c, r))));
            rows.Add(values);
        }

        var seen = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
        var keep = new List<List<string>>();
        foreach (var row in rows)
        {
            // A tab is a fair separator here: it cannot appear in a cell's value.
            if (seen.Add(string.Join('\t', row))) keep.Add(row);
        }

        int gone = rows.Count - keep.Count;
        if (gone == 0) return new Done("Every row was different, so nothing was taken out.", 0);

        for (int i = 0; i < rows.Count; i++)
            for (int c = left; c <= right; c++)
                sheet.Set(new CellRef(c, top + i), i < keep.Count ? keep[i][c - left] : "");

        return new Done($"Took out {gone} row{(gone == 1 ? "" : "s")} that said the same as an earlier one.", gone);
    }

    /// <summary>
    /// Split one column into several at a separator, writing the pieces into the columns to its right.
    ///
    /// It refuses rather than writing over anything: if the columns it would need are not empty, it says so, so
    /// splitting a column can never quietly destroy the one beside it.
    /// </summary>
    public static Result SplitColumn(Workbook book, Sheet sheet, int column, int top, int bottom, string separator)
    {
        if (separator.Length == 0) return new Refused("Say what to split on: a comma, a space, or whatever it is.");
        if (bottom < top) return new Refused("Select some rows.");

        var pieces = new List<string[]>();
        int widest = 0;
        for (int r = top; r <= bottom; r++)
        {
            var text = Value(sheet.Read(new CellRef(column, r)));
            var parts = text.Split(separator);
            pieces.Add(parts);
            widest = Math.Max(widest, parts.Length);
        }
        if (widest <= 1) return new Done($"Nothing in that column has a {Describe(separator)} in it.", 0);

        int lastNeeded = column + widest - 1;
        if (Guard(book, sheet, top, bottom, column, lastNeeded) is Refused blocked) return blocked;

        // Never write over something. The columns to the right have to be empty for as far as this needs them.
        for (int c = column + 1; c <= lastNeeded; c++)
            for (int r = top; r <= bottom; r++)
                if (sheet.Read(new CellRef(c, r)).Kind != CellKind.Empty)
                    return new Refused($"Column {CellRef.ColumnName(c)} has something in it at row {r}. "
                                     + "Splitting would write over it, so Plain has left everything alone. "
                                     + "Insert some empty columns first.");

        int changed = 0;
        for (int i = 0; i < pieces.Count; i++)
        {
            if (pieces[i].Length <= 1) continue;
            for (int p = 0; p < widest; p++)
                sheet.Set(new CellRef(column + p, top + i), p < pieces[i].Length ? pieces[i][p].Trim() : "");
            changed++;
        }

        return new Done($"Split {changed} row{(changed == 1 ? "" : "s")} into {widest} columns.", changed);
    }

    private static string Describe(string separator) => separator switch
    {
        "," => "comma",
        " " => "space",
        ";" => "semicolon",
        "\t" => "tab",
        _ => $"\"{separator}\"",
    };

    /// <summary>What a cell holds, as the thing to move or split: the value, not how it is shown.</summary>
    private static string Value(Cell cell) => cell.Kind switch
    {
        CellKind.Empty => "",
        CellKind.Number or CellKind.Boolean => cell.Raw,
        _ => cell.Display,
    };

    /// <summary>
    /// The same care sorting takes. A formula inside the block would be moved somewhere it does not mean; one
    /// outside reading part of the block would end up reading different values. A formula reading every row of it
    /// is fine for removing duplicates only in the sense that it will recompute, so it is refused here too: unlike
    /// a sort, this changes which values are there, not only their order.
    /// </summary>
    private static Refused? Guard(Workbook book, Sheet sheet, int top, int bottom, int left, int right)
    {
        for (int r = top; r <= bottom; r++)
            for (int c = left; c <= right; c++)
                if (sheet.Read(new CellRef(c, r)).Formula is not null)
                    return new Refused($"There is a formula at {new CellRef(c, r)}. Plain does not move formulas "
                                     + "about, because moving one changes what it means.");

        foreach (var other in book.Sheets)
            foreach (var (at, formula) in other.Formulas())
                foreach (var token in Refs.Scan(formula))
                {
                    var named = token.Range.Sheet;
                    bool same = named is null or "" ? ReferenceEquals(other, sheet)
                                                    : named.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase);
                    if (!same) continue;
                    var range = token.Range;
                    if (range.RowMin > bottom || range.RowMax < top || range.ColumnMin > right || range.ColumnMax < left) continue;

                    return new Refused($"The formula in {other.Name}!{at} reads these cells. Changing them would "
                                     + "change what it answers, so Plain has left them alone.");
                }

        return null;
    }
}
